from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

import pytest
from tkn_genai_bridge import (
    CliSettings,
    Profile,
    ProviderError,
    TokenCounts,
    Usage,
)
from tkn_genai_bridge.models import ResponseMetadata
from tkn_genai_bridge.providers import cli

import excel_catalog_pipeline.context_provider as provider
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.models import ContextConfig

LOGGER = logging.getLogger("bridge-tests")
PNG = b"\x89PNG\r\n\x1a\nfixture"


@pytest.fixture
def setup_provider(tmp_path, monkeypatch):
    profile = Profile(model="fixture-model", cli=CliSettings(executable=sys.executable))
    monkeypatch.setattr(provider, "load_profile", lambda *args, **kwargs: profile)
    image = tmp_path / "001.png"
    image.write_bytes(PNG)
    return image, tmp_path / "usage.json"


def generate(setup_provider, **kwargs):
    image, journal = setup_provider
    return provider.generate_markdown(
        ContextConfig(), {"sheet": "Data"}, [image], journal, LOGGER, **kwargs
    )


def test_real_bridge_receives_ordered_images_schema_and_journals_usage(setup_provider, monkeypatch):
    image, journal = setup_provider
    other = image.with_name("002.png")
    other.write_bytes(PNG + b"second")
    calls = []

    def run(command, prompt, cwd, timeout, **kwargs):
        calls.append(command)
        assert json.loads(journal.read_text())["status"] == "started"
        assert "SOURCE CONTENT" in prompt
        assert "--ignore-user-config" in command
        assert command[command.index("--sandbox") + 1] == "read-only"
        images = [
            Path(command[i + 1]).read_bytes() for i, arg in enumerate(command) if arg == "--image"
        ]
        assert images == [PNG, PNG + b"second"]
        schema = Path(command[command.index("--output-schema") + 1])
        assert json.loads(schema.read_text()) == provider.SCHEMA
        Path(command[command.index("--output-last-message") + 1]).write_text(
            '{"markdown":"### Good"}'
        )
        return '{"type":"turn.completed","usage":{"input_tokens":200,"output_tokens":30,"cached_input_tokens":150}}'

    monkeypatch.setattr(cli, "run_process", run)
    result = provider.generate_markdown(
        ContextConfig(), {"sheet": "Data"}, [image, other], journal, LOGGER
    )
    assert result == "### Good" and len(calls) == 1
    record = json.loads(journal.read_text())
    assert record["inputTokens"] == 200 and record["cachedInputTokens"] == 150
    assert record["usageComplete"] is True
    assert record["generationRecord"]["bridge_version"] == "0.8.0"
    assert len(record["generationRecord"]["images"]) == 2
    assert "### Good" not in journal.read_text()
    assert "SOURCE CONTENT" not in journal.read_text()
    assert str(image) not in journal.read_text()


@pytest.mark.parametrize("code", ["timeout", "process_exit"])
def test_failed_bridge_call_keeps_known_subtotal_and_does_not_retry(
    setup_provider, monkeypatch, code
):
    calls = []

    def run(*args, **kwargs):
        calls.append(True)
        error = ProviderError("provider stopped", code=code)
        error.metadata = ResponseMetadata(
            usage=Usage(
                completeness="partial", known_subtotal=TokenCounts(input_tokens=40, output_tokens=5)
            )
        )
        raise error

    monkeypatch.setattr(cli, "run_process", run)
    records = []
    with pytest.raises(ContextError, match=code):
        generate(setup_provider, on_usage=records.append)
    assert len(calls) == len(records) == 1
    record = json.loads(setup_provider[1].read_text())
    assert record["status"] == ("timeout" if code == "timeout" else "failed")
    assert record["inputTokens"] is None
    assert record["knownInputTokens"] == 40
    assert record["usageComplete"] is False
    assert record["generationRecord"]["error_code"] == code


@pytest.mark.parametrize(
    "output",
    [
        "not json",
        '{"unexpected":1}',
        '{"markdown":""}',
        '{"markdown":"   "}',
        '{"markdown":"<!-- excel-catalog:begin test -->"}',
    ],
)
def test_invalid_output_preserves_reported_usage(setup_provider, monkeypatch, output):
    def run(command, *args, **kwargs):
        Path(command[command.index("--output-last-message") + 1]).write_text(output)
        return '{"type":"turn.completed","usage":{"input_tokens":20,"output_tokens":3}}'

    monkeypatch.setattr(cli, "run_process", run)
    with pytest.raises(ContextError):
        generate(setup_provider)
    record = json.loads(setup_provider[1].read_text())
    assert record["status"] == "failed"
    assert record["inputTokens"] == 20
    assert record["outputTokens"] == 3


def test_no_call_if_usage_journal_cannot_be_written(setup_provider, monkeypatch):
    def unavailable(*args, **kwargs):
        raise OSError("fixture journal unavailable")

    monkeypatch.setattr(provider, "atomic_json", unavailable)
    monkeypatch.setattr(cli, "run_process", lambda *a, **k: pytest.fail("AI must not run"))
    with pytest.raises(OSError, match="journal unavailable"):
        generate(setup_provider)


def test_invalid_image_fails_before_journal_or_provider(setup_provider, monkeypatch):
    setup_provider[0].write_bytes(b"invalid")
    monkeypatch.setattr(cli, "run_process", lambda *a, **k: pytest.fail("AI must not run"))
    with pytest.raises(ContextError, match="invalid_image"):
        generate(setup_provider)
    assert not setup_provider[1].exists()


@pytest.mark.parametrize("images", [[], [Path("missing.png")] * 25])
def test_image_count_is_bounded_before_reading(tmp_path, images):
    with pytest.raises(ContextError, match="max_images"):
        provider.generate_markdown(
            ContextConfig(),
            {"sheet": "Data"},
            images,
            tmp_path / "usage.json",
            LOGGER,
            profile=Profile(),
        )
    assert not list(tmp_path.iterdir())


def test_unsupported_image_profile_stops_before_execution(monkeypatch):
    monkeypatch.setattr(provider, "load_profile", lambda *a, **k: Profile(provider="claude-code"))
    with pytest.raises(ContextError, match="Sheet images require"):
        provider.resolve_profile(ContextConfig())


def test_shared_profile_and_overrides_are_passed_to_bridge(monkeypatch):
    calls = []

    def load(name, *, overrides):
        calls.append((name, overrides))
        return Profile(provider="ollama", model="fixture-vision", local_only=True)

    monkeypatch.setattr(provider, "load_profile", load)
    profile = provider.resolve_profile(
        ContextConfig(bridge_profile="local-vision", overrides={"timeout_seconds": 45})
    )
    assert calls == [("local-vision", {"timeout_seconds": 45})]
    plan = provider.generation_plan(profile)
    assert plan["provider"] == "ollama" and plan["local_only"] is True
    assert "token_estimate" not in plan and "cost_estimate" not in plan
    assert plan["will_call_provider"] is False


def test_unknown_usage_does_not_become_zero():
    values = provider.usage_fields(Usage())
    assert values["inputTokens"] is None and values["knownInputTokens"] is None
    assert values["usageComplete"] is False

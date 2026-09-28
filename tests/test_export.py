from __future__ import annotations

import json
import logging
from dataclasses import replace

import pytest

import excel_catalog_pipeline.export as exporter
from excel_catalog_pipeline.cli import main
from excel_catalog_pipeline.context_profiles import load_prompt, prompt_digest
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.models import AppConfig, ContextConfig, SyncConfig
from tests.helpers import create_workbook
from tests.test_context import rewrite_package


@pytest.fixture
def setup(tmp_path, monkeypatch):
    book = create_workbook(tmp_path / "book.xlsx")
    output = tmp_path / "out" / "result.md"
    state = tmp_path / "state"
    monkeypatch.setattr(exporter, "state_root", lambda: state)
    monkeypatch.setattr(exporter, "resolve_profile", lambda config: object())
    monkeypatch.setattr(exporter, "generation_plan", lambda *args, **kwargs: {})
    calls = []

    def render(snapshot, evidence, output, config):
        output.mkdir(parents=True)
        (output / "001.png").write_bytes(b"image")
        return [{"image": "001.png", "range": "$A$1:$C$9", "overview": False}]

    def generate(config, evidence, images, usage, logger, *, on_usage, stage="sheet", **kwargs):
        calls.append((stage, evidence))
        on_usage({"inputTokens": 12, "outputTokens": 4})
        return "### Overview\n\nCombined." if stage == "workbook" else "### Details\n\nExact facts."

    monkeypatch.setattr(exporter, "render_sheet", render)
    monkeypatch.setattr(exporter, "generate_markdown", generate)
    return book, output, state, calls


def run(setup, **kwargs):
    book, output, _, _ = setup
    return exporter.export_workbook(
        book,
        output,
        ContextConfig(),
        logging.getLogger("test"),
        sheet_names=kwargs.pop("sheet_names", []),
        dry_run=kwargs.pop("dry_run", False),
        force=kwargs.pop("force", False),
        **kwargs,
    )


def test_standalone_export_preserves_workbook_and_publishes_portable_assets(setup):
    book, output, state, calls = setup
    original = book.read_bytes()
    result = run(setup)
    assert result["status"] == "success"
    assert result["usage"]["calls"] == 2
    assert result["usage"]["inputTokens"] == 24
    assert book.read_bytes() == original
    assert not state.exists()
    text = output.read_text(encoding="utf-8")
    assert "Combined." in text and "Exact facts." in text
    assert "excel-catalog:begin" not in text
    image = calls[0][1]["images"][0]["relativePath"]
    assert (output.parent / image).read_bytes() == b"image"
    manifest = json.loads(next(output.parent.glob("excel-assets-*/manifest.json")).read_text())
    assert manifest["sheets"] == ["Data"]
    assert not list(output.parent.glob("*.lock"))


def test_dry_run_no_writes_render_or_ai(setup, monkeypatch):
    book, output, state, calls = setup
    monkeypatch.setattr(exporter, "render_sheet", lambda *args: pytest.fail("render"))
    assert run(setup, dry_run=True)["status"] == "planned"
    assert not output.parent.exists() and not state.exists() and not calls


def test_existing_output_refused_before_ai_and_force_supported(setup):
    _, output, _, calls = setup
    output.parent.mkdir()
    output.write_text("reviewed")
    with pytest.raises(ContextError, match="--force"):
        run(setup)
    assert not calls and output.read_text() == "reviewed"
    run(setup, force=True)
    assert "Combined." in output.read_text(encoding="utf-8")


def test_synthesis_failure_preserves_existing_output_and_cleans_staging(setup, monkeypatch):
    _, output, _, _ = setup
    output.parent.mkdir()
    output.write_text("reviewed")
    generate = exporter.generate_markdown

    def fail(*args, **kwargs):
        if kwargs.get("stage") == "workbook":
            raise ContextError("synthesis failed")
        return generate(*args, **kwargs)

    monkeypatch.setattr(exporter, "generate_markdown", fail)
    with pytest.raises(ContextError, match="synthesis failed"):
        run(setup, force=True)
    assert output.read_text() == "reviewed"
    assert list(output.parent.iterdir()) == [output]


def test_concurrent_edit_is_preserved(setup, monkeypatch):
    _, output, _, _ = setup
    generate = exporter.generate_markdown

    def edit(*args, **kwargs):
        output.write_text("external edit")
        return generate(*args, **kwargs)

    monkeypatch.setattr(exporter, "generate_markdown", edit)
    with pytest.raises(ContextError, match="changed during"):
        run(setup)
    assert output.read_text() == "external edit"
    assert not list(output.parent.glob("excel-assets-*"))


def test_unknown_and_hidden_sheet_selection(setup):
    book, _, _, calls = setup
    with pytest.raises(ContextError, match="Unknown sheet"):
        run(setup, sheet_names=["missing"])
    assert not calls
    import zipfile

    with zipfile.ZipFile(book) as archive:
        xml = archive.read("xl/workbook.xml").replace(b'name="Data"', b'name="Data" state="hidden"')
    book.write_bytes(rewrite_package(book.read_bytes(), {"xl/workbook.xml": xml}))
    with pytest.raises(ContextError, match="No worksheets"):
        run(setup)
    assert run(setup, sheet_names=["Data"], dry_run=True)["sheets"] == ["Data"]


def test_cli_export_needs_no_sources(setup, monkeypatch, capsys):
    book, output, _, _ = setup
    import excel_catalog_pipeline.cli as cli

    monkeypatch.setattr(cli, "load_config", lambda **kwargs: AppConfig("1.0.0", (), SyncConfig()))
    assert (
        main(["export", str(book), "--output", str(output), "--profile", "default-en", "--dry-run"])
        == 0
    )
    payload = json.loads(capsys.readouterr().out)
    assert payload["profile"] == "default-en"


def test_prompt_profiles_language_and_content_hash():
    ja = ContextConfig(prompt_profile="default-ja")
    en = replace(ja, prompt_profile="default-en")
    assert "Japanese" in load_prompt(ja)
    assert "English" in load_prompt(en)
    assert "Relationships between sheets" in load_prompt(en, stage="workbook")
    assert prompt_digest(ja) != prompt_digest(en)
    with pytest.raises(ContextError, match="Unknown context profile"):
        load_prompt(replace(ja, prompt_profile="../unknown"))

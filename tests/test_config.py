from __future__ import annotations

from pathlib import Path

import pytest

import excel_catalog_pipeline.config as config_module
from excel_catalog_pipeline.config import (
    ConfigError,
    config_as_dict,
    config_template_text,
    init_user_config,
    load_config,
)


def write_config(path: Path, source_id: str, max_chars: int) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"""schema_version: 1
sources:
  - id: {source_id}
    path: 'C:\\path\\to\\source'
    include: ["**/*.xlsx"]
    notes:
      root: 'C:\\path\\to\\notes'
sync:
  max_extracted_text_chars: {max_chars}
""",
        encoding="utf-8",
    )


def test_config_precedence(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    global_path = tmp_path / "global.yaml"
    cwd = tmp_path / "cwd"
    local_path = cwd / ".tkn" / "config.yaml"
    explicit = tmp_path / "explicit.yaml"
    write_config(global_path, "global", 100)
    write_config(local_path, "local", 200)
    write_config(explicit, "explicit", 300)
    monkeypatch.setattr(config_module, "global_config_path", lambda: global_path)
    loaded = load_config(explicit=explicit, cwd=cwd)
    assert loaded.sources[0].id == "explicit"
    assert loaded.sync.max_extracted_text_chars == 300
    assert loaded.loaded_files == (global_path, local_path, explicit)


def test_unknown_key_is_rejected(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    path = tmp_path / "bad.yaml"
    path.write_text("schema_version: 1\nsources: []\nunknown: true\n", encoding="utf-8")
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "missing")
    with pytest.raises(ConfigError, match="Unknown top-level"):
        load_config(explicit=path, cwd=tmp_path)


def test_source_traversal_defaults_are_non_recursive(tmp_path: Path) -> None:
    loaded = config_module.validate_config(
        {
            "schema_version": 1,
            "sync": dict(config_module.DEFAULT_CONFIG["sync"]),
            "sources": [
                {
                    "id": "example",
                    "path": str(tmp_path / "source"),
                    "notes": {"root": str(tmp_path / "notes")},
                }
            ],
        }
    )

    source = loaded.sources[0]
    assert source.recursive is False
    assert source.include == ("*.xlsx", "*.xlsm")
    assert source.ignore == ()
    rendered = config_as_dict(loaded)["sources"][0]
    assert rendered["recursive"] is False
    assert rendered["ignore"] == []


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        ("recursive", "yes", "recursive must be boolean"),
        ("ignore", "archive/**", "ignore must be a string list"),
    ],
)
def test_invalid_source_traversal_settings_are_rejected(
    tmp_path: Path, field: str, value: object, message: str
) -> None:
    source = {
        "id": "example",
        "path": str(tmp_path / "source"),
        "notes": {"root": str(tmp_path / "notes")},
        field: value,
    }
    with pytest.raises(ConfigError, match=message):
        config_module.validate_config(
            {
                "schema_version": 1,
                "sources": [source],
                "sync": dict(config_module.DEFAULT_CONFIG["sync"]),
            }
        )


def test_unknown_note_profile_is_rejected_in_config(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "bad-profile.yaml"
    path.write_text(
        r"""schema_version: 1
sources:
  - id: example
    path: 'C:\path\to\source'
    notes:
      root: 'C:\path\to\notes'
      profile: missing-profile
""",
        encoding="utf-8",
    )
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "missing")

    with pytest.raises(ConfigError, match="Unknown note profile"):
        load_config(explicit=path, cwd=tmp_path)


def test_packaged_template_matches_repository_example() -> None:
    repository_template = Path(".tkn/config.example.yaml").read_text(encoding="utf-8")
    assert config_template_text() == repository_template


def test_init_user_config_is_idempotent_and_protects_edits(tmp_path: Path) -> None:
    target = tmp_path / ".tkn" / "excel_catalog_pipeline" / "config.yaml"

    status, path = init_user_config(target=target)
    assert status == "created"
    assert path == target.resolve()
    assert target.read_text(encoding="utf-8") == config_template_text()

    status, _ = init_user_config(target=target)
    assert status == "unchanged"

    target.write_text("reviewed: true\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="already exists and differs"):
        init_user_config(target=target)
    assert target.read_text(encoding="utf-8") == "reviewed: true\n"

    status, _ = init_user_config(target=target, force=True)
    assert status == "replaced"
    assert target.read_text(encoding="utf-8") == config_template_text()

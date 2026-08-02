from __future__ import annotations

import json
from pathlib import Path

import excel_catalog_pipeline.config as config_module
from excel_catalog_pipeline.cli import main

from .helpers import create_workbook


def test_config_show_outputs_one_json_document(tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text("schema_version: 1\nsources: []\n", encoding="utf-8")
    result = main(["--config", str(config), "config", "show"])
    captured = capsys.readouterr()
    assert result == 0
    assert json.loads(captured.out)["command"] == "config show"
    assert captured.out.count("\n") == 1


def test_config_init_creates_user_config(monkeypatch, tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    target = tmp_path / "user" / "config.yaml"
    monkeypatch.setattr(config_module, "global_config_path", lambda: target)

    result = main(["config", "init"])
    payload = json.loads(capsys.readouterr().out)
    assert result == 0
    assert payload == {
        "status": "created",
        "command": "config init",
        "configPath": str(target.resolve()),
    }
    assert target.exists()

    result = main(["config", "init"])
    payload = json.loads(capsys.readouterr().out)
    assert result == 0
    assert payload["status"] == "unchanged"


def test_cli_pull_dry_run_writes_report_not_note(tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    create_workbook(workbooks / "book.xlsx")
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: 1
sources:
  - id: example
    path: "{workbooks.as_posix()}"
    include: ["**/*.xlsx"]
    notes:
      root: "{notes.as_posix()}"
''',
        encoding="utf-8",
    )
    result = main(["--config", str(config), "--report-dir", str(tmp_path / "reports"), "pull"])
    payload = json.loads(capsys.readouterr().out)
    assert result == 0
    assert payload["statusCounts"] == {"would-create": 1}
    assert Path(payload["reportPath"], "summary.json").exists()
    assert not notes.exists()

from __future__ import annotations

import json
from pathlib import Path

import excel_catalog_pipeline.cli as cli_module
import excel_catalog_pipeline.config as config_module
from excel_catalog_pipeline.cli import main
from excel_catalog_pipeline.models import Action

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


def test_cli_push_logs_non_unchanged_files_and_readable_summary(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
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

    def fake_run_push(*args, on_written, **kwargs):  # type: ignore[no-untyped-def]
        actions = [
            Action(status="written", source_root_id="example", source_path="one.xlsx"),
            Action(status="unchanged", source_root_id="example", source_path="two.xlsx"),
            Action(
                status="would-write",
                source_root_id="example",
                source_path="nested/three.xlsx",
                message="Metadata differs.",
            ),
            Action(
                status="missing-source",
                source_root_id="example",
                note_path=str(notes / "missing.xlsx.md"),
                message="No unique workbook matches this note.",
            ),
            Action(
                status="conflict",
                source_root_id="example",
                source_path="conflict.xlsx",
                conflict_fields=["title"],
                message="Source and note changed differently from the base state.",
            ),
        ]
        for action in actions:
            if action.status == "written":
                on_written(action)
        return actions, None

    monkeypatch.setattr(cli_module, "run_push", fake_run_push)
    result = main(
        [
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "push",
            "--write-excel",
        ]
    )
    captured = capsys.readouterr()

    payload = json.loads(captured.out)
    assert result == 2
    assert captured.out.count("\n") == 1
    assert "[INFO] Running push for 1 configured source(s): example." in captured.err
    assert "[SUCCESS] [written] sourcePath=one.xlsx | message=-" in captured.err
    assert (
        "[INFO] [would-write] sourcePath=nested/three.xlsx | message=Metadata differs."
        in captured.err
    )
    assert (
        f"[WARNING] [missing-source] sourcePath=- | notePath={notes / 'missing.xlsx.md'} "
        "| message=No unique workbook matches this note."
    ) in captured.err
    assert (
        "[WARNING] [conflict] sourcePath=conflict.xlsx | "
        "message=Source and note changed differently from the base state."
    ) in captured.err
    assert "two.xlsx" not in captured.err
    assert "[INFO] Summary:\n{" in captured.err
    assert '  "statusCounts": {' in captured.err
    assert '    "unchanged": 1' in captured.err
    assert payload["statusCounts"]["unchanged"] == 1

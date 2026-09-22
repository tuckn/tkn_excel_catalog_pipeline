from __future__ import annotations

import json
from pathlib import Path

import pytest

import excel_catalog_pipeline.cli as cli_module
import excel_catalog_pipeline.config as config_module
import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.cli import main
from excel_catalog_pipeline.models import Action

from .helpers import create_workbook


def test_cli_uses_tkn_prefixed_program_name() -> None:
    assert cli_module.build_parser().prog == "tkn-excel-catalog"


def test_mutating_command_help_explains_normal_write_and_dry_run(capsys) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SystemExit) as exc_info:
        main(["pull", "--help"])
    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "Normal execution writes proxy notes" in captured.out
    assert "--dry-run" in captured.out
    assert "Preview and validate planned changes without writing" in captured.out
    assert "workbooks, notes, state, cache, reports" in captured.out
    assert "or external" in captured.out
    assert "synchronization command uses no network" in " ".join(captured.out.split())
    assert "--write-notes" in captured.out
    assert "Deprecated compatibility option" in captured.out


def test_config_show_prints_path_before_indented_json(monkeypatch, tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "user.yaml")
    config = tmp_path / "config.yaml"
    config.write_text("schema_version: 1\nsources: []\n", encoding="utf-8")
    result = main(["--config", "config.yaml", "config", "show"])
    captured = capsys.readouterr()
    assert result == 0
    header, body = captured.out.split("\n\n", 1)
    assert header == f"Config file: {config.resolve()}"
    payload = json.loads(body)
    assert payload["command"] == "config show"
    assert payload["config"]["loadedConfigFiles"] == [str(config.resolve())]
    assert '\n  "config": {\n    "schema_version": 1,' in body
    assert captured.err == ""


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


def test_cli_pull_without_option_writes_note_state_and_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
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
    monkeypatch.setattr(pipeline_module, "state_path", lambda: tmp_path / "state.json")
    result = main(["--config", str(config), "--report-dir", str(tmp_path / "reports"), "pull"])
    captured = capsys.readouterr()
    summary_path = next((tmp_path / "reports").glob("*-pull/summary.json"))
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    assert result == 0
    assert captured.out == ""
    assert f"[SUCCESS] [created] notePath={notes / 'book.xlsx.md'}" in captured.err
    assert "sourcePath=book.xlsx" in captured.err
    assert "  mode: write" in captured.err
    assert payload["statusCounts"] == {"created": 1}
    assert Path(payload["reportPath"], "summary.json").exists()
    details = json.loads(Path(payload["reportPath"], "details.json").read_text(encoding="utf-8"))
    assert details[0]["sourceToNoteFields"] == []
    assert details[0]["noteToSourceFields"] == []
    assert (notes / "book.xlsx.md").exists()


def test_cli_pull_dry_run_changes_no_persistent_files(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    reports = tmp_path / "reports"
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
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)

    result = main(
        [
            "--config",
            str(config),
            "--report-dir",
            str(reports),
            "pull",
            "--dry-run",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert captured.out == ""
    assert f"[INFO] [would-create] notePath={notes / 'book.xlsx.md'}" in captured.err
    assert "  mode: dry-run" in captured.err
    assert "  report: -" in captured.err
    assert "  summary JSON: -" in captured.err
    assert "  differences CSV: -" in captured.err
    assert "no persistent report was written" in captured.err
    assert "--report-dir has no effect in dry-run mode" in captured.err
    assert not notes.exists()
    assert not reports.exists()
    assert not state.exists()


def test_cli_status_logs_grouped_counts_and_attention_items(
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

    def fake_run_status(*args, **kwargs):  # type: ignore[no-untyped-def]
        return [
            Action(status="tracked", source_root_id="example", source_path="tracked.xlsx"),
            Action(
                status="untracked-source",
                source_root_id="example",
                source_path="nested/new.xlsx",
                message="stateMatch=none",
            ),
            *(
                Action(
                    status="untracked-note",
                    source_root_id="example",
                    note_path=str(notes / f"orphan-{index:02}.xlsx.md"),
                )
                for index in range(21)
            ),
            Action(
                status="duplicate-id",
                source_root_id="example",
                source_path="duplicate.xlsx",
                workbook_id="duplicate-workbook-id",
            ),
        ]

    monkeypatch.setattr(cli_module, "run_status", fake_run_status)
    result = main(["--config", str(config), "--report-dir", str(tmp_path / "reports"), "status"])
    captured = capsys.readouterr()

    assert result == 0
    assert captured.out == ""
    assert "[INFO] Status results:\n  example:" in captured.err
    assert "    tracked workbooks: 1" in captured.err
    assert "    untracked workbooks: 1" in captured.err
    assert "    untracked proxy notes: 21" in captured.err
    assert "    workbooks with duplicate IDs: 1" in captured.err
    assert "    items requiring attention:" in captured.err
    assert "      untracked workbooks:\n        - nested/new.xlsx" in captured.err
    assert "      untracked proxy notes:\n        - orphan-00.xlsx.md" in captured.err
    assert "        - orphan-19.xlsx.md" in captured.err
    assert "        ... 1 more; see the report below" in captured.err
    assert "orphan-20.xlsx.md" not in captured.err
    assert (
        "      workbooks with duplicate IDs:\n"
        "        - duplicate.xlsx | workbookId=duplicate-workbook-id" in captured.err
    )
    assert str(notes.resolve()) not in captured.err
    assert "tracked.xlsx" not in captured.err
    summary_path = next((tmp_path / "reports").glob("*-status/summary.json"))
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    assert payload["statusCounts"] == {
        "tracked": 1,
        "untracked-source": 1,
        "untracked-note": 21,
        "duplicate-id": 1,
    }


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

    def fake_run_push(*args, on_action, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_excel"] is True
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
            on_action(action)
        return actions, None

    monkeypatch.setattr(cli_module, "run_push", fake_run_push)
    result = main(
        [
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "push",
        ]
    )
    captured = capsys.readouterr()

    assert result == 2
    assert captured.out == ""
    assert "[INFO] Running push for 1 configured source(s): example." in captured.err
    assert "[INFO] Selected source configuration:\n  - id: example" in captured.err
    assert f"    path: '{workbooks.resolve()}'" in captured.err
    assert "    recursive: false" in captured.err
    assert '      - "**/*.xlsx"' in captured.err
    assert "    ignore: []" in captured.err
    assert f"      root: '{notes.resolve()}'" in captured.err
    assert "      profile: tkn-obsidian-v1" in captured.err
    assert "      frontmatter_term_format: obsidian-link" in captured.err
    assert "      rename_adapter: report-only" in captured.err
    assert "[SUCCESS] [written] fileName=one.xlsx | sourcePath=one.xlsx | message=-" in captured.err
    assert (
        "[INFO] [would-write] fileName=three.xlsx | sourcePath=nested/three.xlsx | "
        "message=Metadata differs." in captured.err
    )
    assert (
        "[WARNING] [missing-source] missing.xlsx.md | message=No unique workbook matches this note."
    ) in captured.err
    assert "[missing-source] fileName=" not in captured.err
    assert "[missing-source] sourcePath=-" not in captured.err
    assert (
        "[WARNING] [conflict] fileName=conflict.xlsx | sourcePath=conflict.xlsx | "
        "message=Source and note changed differently from the base state."
    ) in captured.err
    assert str(notes / "missing.xlsx.md") not in captured.err
    assert "two.xlsx" not in captured.err
    assert "[INFO] Summary:\n  command: push" in captured.err
    assert "  mode: write" in captured.err
    assert "  status counts:\n    written: 1\n    unchanged: 1" in captured.err
    assert "  differences CSV:" in captured.err
    summary_path = next((tmp_path / "reports").glob("*-push/summary.json"))
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    assert payload["statusCounts"]["unchanged"] == 1


def test_cli_pull_verbose_dry_run_logs_values_without_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: 1
sources:
  - id: example
    path: "{(tmp_path / "workbooks").as_posix()}"
    include: ["**/*.xlsx"]
    notes:
      root: "{(tmp_path / "notes").as_posix()}"
''',
        encoding="utf-8",
    )

    def fake_run_pull(*args, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_notes"] is False
        return [
            Action(
                status="would-update",
                source_root_id="example",
                source_path="nested/book.xlsx",
                note_path=str(tmp_path / "notes" / "nested" / "book.xlsx.md"),
                changed_fields=["author"],
                source_to_note_fields=["author"],
                field_differences=[
                    {
                        "field": "author",
                        "direction": "push",
                        "plannedDirection": "source-to-note",
                        "baseValue": "Example author",
                        "excelValue": "Example author",
                        "noteValue": "",
                    }
                ],
            ),
            Action(status="unchanged", source_root_id="example", source_path="same.xlsx"),
        ]

    monkeypatch.setattr(cli_module, "run_pull", fake_run_pull)
    result = main(
        [
            "-v",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "pull",
            "--dry-run",
            "--prefer-source",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert captured.out == ""
    assert (
        f"[INFO] [would-update] notePath={tmp_path / 'notes' / 'nested' / 'book.xlsx.md'} | "
        "sourcePath=nested/book.xlsx" in captured.err
    )
    assert "[DEBUG] Metadata differences:" in captured.err
    assert "[would-update] nested/book.xlsx" in captured.err
    assert (
        "author | Excel: Example author | Markdown: <empty> | "
        "base: Example author | apply: Excel -> Markdown" in captured.err
    )
    assert "[INFO] Summary:\n  command: pull\n  mode: dry-run" in captured.err
    assert "    would-update: 1\n    unchanged: 1" in captured.err

    assert "  summary JSON: -" in captured.err
    assert "  differences CSV: -" in captured.err
    assert not (tmp_path / "reports").exists()


def test_legacy_write_option_is_accepted_with_deprecation_warning(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        "schema_version: 1\nsources:\n"
        f"  - id: example\n    path: '{tmp_path.as_posix()}'\n"
        f"    notes:\n      root: '{(tmp_path / 'notes').as_posix()}'\n",
        encoding="utf-8",
    )

    def fake_run_pull(*args, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_notes"] is True
        return []

    monkeypatch.setattr(cli_module, "run_pull", fake_run_pull)
    result = main(
        [
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "pull",
            "--write-notes",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert "--write-notes is deprecated and no longer required" in captured.err


def test_cli_push_dry_run_validates_allowed_rename_without_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        "schema_version: 1\nsources:\n"
        f"  - id: example\n    path: '{tmp_path.as_posix()}'\n"
        f"    notes:\n      root: '{(tmp_path / 'notes').as_posix()}'\n",
        encoding="utf-8",
    )

    def fake_run_push(*args, on_action, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_excel"] is False
        assert kwargs["allow_rename"] is True
        action = Action(
            status="would-write",
            source_root_id="example",
            source_path="old.xlsx",
            changed_fields=["sourceFileName"],
        )
        on_action(action)
        return [action], None

    monkeypatch.setattr(cli_module, "run_push", fake_run_push)
    result = main(
        [
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "push",
            "--dry-run",
            "--allow-rename",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert "[INFO] [would-write] fileName=old.xlsx | sourcePath=old.xlsx" in captured.err
    assert "  mode: dry-run" in captured.err
    assert not (tmp_path / "reports").exists()


def test_cli_adopt_dry_run_lists_targets_without_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        "schema_version: 1\nsources:\n"
        f"  - id: example\n    path: '{tmp_path.as_posix()}'\n"
        f"    notes:\n      root: '{(tmp_path / 'notes').as_posix()}'\n",
        encoding="utf-8",
    )

    def fake_run_adopt(*args, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_excel"] is False
        return [Action(status="would-adopt", source_root_id="example", source_path="book.xlsx")], None

    monkeypatch.setattr(cli_module, "run_adopt", fake_run_adopt)
    result = main(
        [
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "adopt",
            "--dry-run",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert "[INFO] [would-adopt] sourcePath=book.xlsx" in captured.err
    assert "    would-adopt: 1" in captured.err
    assert not (tmp_path / "reports").exists()

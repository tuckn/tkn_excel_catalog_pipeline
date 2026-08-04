from __future__ import annotations

from pathlib import Path

import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.adapters.markdown import NoteError
from excel_catalog_pipeline.adapters.ooxml import write_properties
from excel_catalog_pipeline.config import validate_config
from excel_catalog_pipeline.pipeline import run_pull, run_push

from .helpers import create_workbook


def app_config(tmp_path: Path):  # type: ignore[no-untyped-def]
    return validate_config(
        {
            "schema_version": 1,
            "sources": [
                {
                    "id": "example",
                    "path": str(tmp_path / "workbooks"),
                    "include": ["**/*.xlsx", "**/*.xlsm"],
                    "notes": {"root": str(tmp_path / "notes"), "rename_adapter": "filesystem"},
                }
            ],
            "sync": {
                "pull_preserves_user_metadata": True,
                "delete_missing_notes": False,
                "delete_missing_workbooks": False,
                "allow_source_rename": False,
                "max_extracted_text_chars": 12000,
            },
        }
    )


def test_pull_dry_run_then_write_then_unchanged(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    dry = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in dry] == ["would-create"]
    assert not state.exists()
    written = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in written] == ["created"]
    assert state.exists()
    again = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in again] == ["unchanged"]


def test_push_detects_note_only_change_and_writes_backup(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    text = note_path.read_text(encoding="utf-8")
    note_path.write_text(
        text.replace("title: Example title", "title: Human title"), encoding="utf-8"
    )
    dry, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in dry] == ["would-write"]
    observed_actions = []
    written_actions = []
    applied, backup = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
        on_action=observed_actions.append,
        on_written=written_actions.append,
    )
    assert [action.status for action in applied] == ["written"]
    assert observed_actions == applied
    assert written_actions == applied
    assert backup is not None and any(backup.iterdir())
    from excel_catalog_pipeline.adapters.ooxml import inspect_workbook

    assert (
        inspect_workbook(workbook_path, config.sources[0], max_text_chars=1).core["title"]
        == "Human title"
    )


def test_push_preserves_repeated_title_whitespace(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    note_path.write_text(
        note_path.read_text(encoding="utf-8").replace(
            "title: Example title", 'title: "Human  title"'
        ),
        encoding="utf-8",
    )

    actions, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["written"]
    from excel_catalog_pipeline.adapters.ooxml import inspect_workbook

    assert (
        inspect_workbook(workbook_path, config.sources[0], max_text_chars=1).core["title"]
        == "Human  title"
    )
    again, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in again] == ["unchanged"]


def test_source_rename_is_detected_by_stable_id(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    old_path = create_workbook(
        tmp_path / "workbooks" / "book.xlsx",
        workbook_id="11111111-1111-1111-1111-111111111111",
    )
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    new_path = old_path.with_name("renamed.xlsx")
    old_path.replace(new_path)
    planned = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in planned] == ["would-update"]
    applied = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in applied] == ["updated"]
    assert (tmp_path / "notes" / "renamed.xlsx.md").exists()
    assert not (tmp_path / "notes" / "book.xlsx.md").exists()


def test_missing_note_is_reported_without_source_deletion(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    (tmp_path / "notes" / "book.xlsx.md").unlink()
    actions = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in actions] == ["missing-note"]
    assert workbook.exists()


def test_push_rolls_back_workbook_when_note_refresh_fails(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    note.write_text(
        note.read_text(encoding="utf-8").replace("title: Example title", "title: Human title"),
        encoding="utf-8",
    )
    workbook_before = workbook.read_bytes()
    note_before = note.read_text(encoding="utf-8")

    def fail_render(*args, **kwargs):  # type: ignore[no-untyped-def]
        raise NoteError("synthetic refresh failure")

    monkeypatch.setattr(pipeline_module, "render_note", fail_render)
    actions, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["write-error"]
    assert workbook.read_bytes() == workbook_before
    assert note.read_text(encoding="utf-8") == note_before


def test_push_stops_when_both_sides_changed_differently(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    note.write_text(
        note.read_text(encoding="utf-8").replace("title: Example title", "title: Note title"),
        encoding="utf-8",
    )
    write_properties(
        workbook,
        core={"title": "Source title"},
        backup_dir=tmp_path / "external-change-backup",
    )
    actions, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["conflict"]
    assert actions[0].conflict_fields == ["title"]

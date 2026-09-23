from __future__ import annotations

import logging
from pathlib import Path

import pytest

import excel_catalog_pipeline.context as context
import excel_catalog_pipeline.context_import as importer
from excel_catalog_pipeline.adapters.markdown import read_note, render_note, workbook_path
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.models import ContextConfig, SourceConfig
from excel_catalog_pipeline.note_layout import migrate_layout
from tests.helpers import create_workbook


@pytest.fixture
def setup_import(tmp_path, monkeypatch):
    source = SourceConfig("example", tmp_path / "workbooks", ("*.xlsx",), tmp_path / "notes")
    workbook = create_workbook(source.path / "book.xlsx", workbook_id="fixture-id")
    note = source.note_root / "book.xlsx.md"
    note.parent.mkdir()
    note.write_bytes(
        (
            "---\r\ntype: Excel\r\nschemaVersion: '2.0'\r\ntitle: Keep title\r\n"
            "description: '' # preserve comment\r\ncomments: Excel comment\r\nsourceRoot: example\r\n"
            "sourceId: example:fixture-id\r\nsourceFileName: book.xlsx\r\n"
            "date: 2025-01-02T00:00:00+09:00\r\nnoteId: keep-id\r\ncustom: [one, two]\r\n---\r\n\r\n"
            "# Keep title\r\n\r\n## Overview\r\n\r\nMy overview.\r\n\r\n"
            "<!-- excel-catalog:begin workbook-path -->\r\n## Workbook Path\r\n\r\n"
            f"`{workbook}`\r\n<!-- excel-catalog:end workbook-path -->\r\n\r\n"
            "<!-- excel-catalog:begin workbook-map -->\r\n## Workbook Map\r\n\r\n- Data\r\n"
            "<!-- excel-catalog:end workbook-map -->\r\n\r\n## My notes\r\n\r\nKeep this exactly.\r\n"
        ).encode()
    )
    evidence = tmp_path / "input"
    evidence.mkdir()
    (evidence / "image.png").write_bytes(b"image bytes")
    markdown = evidence / "Data.md"
    markdown.write_text(
        '---\ntype: ExcelContext\nsourceId: example:fixture-id\nsourceSheet: Data\nsourceSheetId: "1"\n'
        "sourceFileName: book.xlsx\ncapturedAt: 2026-01-01T00:00:00Z\n---\n\n"
        "# Sheet title\n\nIntroduction.\n\n## Topic\n\n### Detail\n\n"
        "[Figure](image.png)\n\n```markdown\n# Untouched code\n[Example](missing.png)\n```\n",
        encoding="utf-8",
    )
    state = tmp_path / "state"
    monkeypatch.setattr(context, "state_root", lambda: state)
    monkeypatch.setattr(importer, "state_root", lambda: state)

    def run(**kwargs):
        return importer.import_context(
            source,
            ContextConfig(),
            workbook_selector="book.xlsx",
            sheet_name="Data",
            markdown_path=markdown,
            dry_run=kwargs.get("dry_run", False),
            force=kwargs.get("force", False),
            description=kwargs.get("description"),
            logger=logging.getLogger("test-import"),
        )

    return source, workbook, note, markdown, state, run


def test_import_dry_run_makes_no_artifacts(setup_import):
    source, workbook, note, markdown, state, run = setup_import
    before = note.read_bytes()
    result = run(dry_run=True)
    assert result["result"] == "planned" and result["imageCount"] == 1
    assert note.read_bytes() == before
    assert not state.exists() and not (source.note_root / "img").exists()


def test_import_layout_provenance_images_and_idempotency(setup_import):
    source, workbook, note, markdown, state, run = setup_import
    before = note.read_bytes()
    source_before = workbook.read_bytes()
    input_before = markdown.read_bytes()
    result = run()
    assert result["result"] == "imported" and result["usage"]["calls"] == 0
    assert Path(result["backupPath"]).read_bytes() == before
    current = read_note(note)
    assert current.frontmatter["description"] == "My overview."
    assert current.frontmatter["sourceFullPath"] == str(workbook.resolve())
    assert current.frontmatter["type"] == "Excel" and current.frontmatter["noteId"] == "keep-id"
    assert current.frontmatter["schemaVersion"] == "2.1"
    assert b"custom: [one, two]\r\n" in note.read_bytes()
    assert b"# preserve comment\r\n" in note.read_bytes()
    assert "## Overview" not in current.body and "## Workbook Path" not in current.body
    assert (
        "## Data (sheetId: 1)" in current.body
        and "### Topic" in current.body
        and "#### Detail" in current.body
    )
    assert "# Untouched code\n[Example](missing.png)" in current.body
    assert current.body.index("## Workbook Map") < current.body.index("## Data (sheetId: 1)")
    assert current.body.endswith("## My notes\n\nKeep this exactly.\n")
    assert len(list((source.note_root / "img").rglob("*.png"))) == 1
    assert len(list((source.note_root / "img").rglob("provenance.json"))) == 1
    after = note.read_bytes()
    assert run()["result"] == "unchanged" and note.read_bytes() == after
    assert workbook.read_bytes() == source_before and markdown.read_bytes() == input_before
    assert workbook_path(current) == workbook.resolve()


def test_imported_context_retained_by_build_and_pull(setup_import, monkeypatch):
    source, workbook, note, markdown, state, run = setup_import
    run()
    before = note.read_bytes()

    def forbidden(*args, **kwargs):
        pytest.fail("AI/render must not run on imported content")

    monkeypatch.setattr(context, "render_sheet", forbidden)
    monkeypatch.setattr(context, "generate_markdown", forbidden)
    result = context.run_context(
        source,
        ContextConfig(),
        workbook_selector="book.xlsx",
        sheet_names=["Data"],
        all_sheets=False,
        list_only=False,
        dry_run=False,
        force=False,
        logger=logging.getLogger("test-import"),
    )
    assert result["results"][0]["status"] == "retained"
    assert result["usage"]["calls"] == 0 and note.read_bytes() == before
    current = read_note(note)
    regenerated = render_note(
        inspect_workbook(workbook, source, max_text_chars=100), source, existing=current
    )
    assert context.existing_block(regenerated, "1") == context.existing_block(current.body, "1")
    assert "## Overview" not in regenerated and "## Workbook Path" not in regenerated
    assert current.frontmatter["description"] in regenerated


def test_import_preserves_edited_block_unless_force(setup_import):
    _, _, note, _, _, run = setup_import
    run()
    note.write_bytes(note.read_bytes().replace(b"Introduction.", b"My revised introduction."))
    before = note.read_bytes()
    with pytest.raises(ContextError, match="edited"):
        run()
    assert note.read_bytes() == before
    run(force=True)
    assert b"Introduction." in note.read_bytes()


def test_import_source_identity_mismatch_fails_before_writes(setup_import):
    source, _, note, markdown, state, run = setup_import
    markdown.write_text(
        markdown.read_text(encoding="utf-8").replace("example:fixture-id", "example:other"),
        encoding="utf-8",
    )
    with pytest.raises(ContextError, match="sourceId"):
        run()
    assert not state.exists() and not (source.note_root / "img").exists()


@pytest.mark.parametrize(
    "target", ["../outside.png", "C:/outside.png", "file:///C:/outside.png", "missing.png"]
)
def test_import_rejects_missing_or_escaping_local_assets(setup_import, target):
    source, _, _, markdown, state, run = setup_import
    markdown.write_text(
        markdown.read_text(encoding="utf-8").replace("(image.png)", f"({target})"), encoding="utf-8"
    )
    with pytest.raises(ContextError):
        run()
    assert not state.exists() and not (source.note_root / "img").exists()


def test_failed_state_publication_rolls_back_note(setup_import, monkeypatch):
    _, _, note, _, _, run = setup_import
    before = note.read_bytes()
    real = importer.atomic_json

    def fail_state(path, payload):
        if path.name == "sheet-1.json":
            raise OSError("state unavailable")
        real(path, payload)

    monkeypatch.setattr(importer, "atomic_json", fail_state)
    with pytest.raises(OSError, match="unavailable"):
        run()
    assert note.read_bytes() == before


def test_description_override_and_preserving_existing_summary(setup_import):
    _, _, note, _, _, run = setup_import
    note.write_bytes(
        note.read_bytes().replace(b"description: ''", b"description: Existing summary")
    )
    run()
    assert read_note(note).frontmatter["description"] == "Existing summary\n\nMy overview."
    run(description="Chosen summary.")
    assert read_note(note).frontmatter["description"] == "Chosen summary."


def test_migration_does_not_delete_code_or_path_annotations():
    text = "---\nschemaVersion: '2.0'\ndescription: ''\ncomments: ''\n---\n\n# Title\n\n```markdown\n## Overview\nExample\n```\n\n## Workbook Path\n\n`old.xlsx`\n\nCustom path note.\n\n## Next\n\nKeep.\n"
    result = migrate_layout(text, "new.xlsx")
    assert "```markdown\n## Overview\nExample\n```" in result
    assert "Custom path note." in result and "## Next" in result
    assert "`old.xlsx`" not in result
    assert migrate_layout(result, "new.xlsx") == result

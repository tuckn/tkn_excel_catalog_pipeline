from __future__ import annotations

from pathlib import Path

from excel_catalog_pipeline.adapters.markdown import (
    metadata_to_nouns,
    note_metadata,
    read_note,
    render_note,
)
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook
from excel_catalog_pipeline.models import SourceConfig

from .helpers import create_workbook


def config(root: Path) -> SourceConfig:
    return SourceConfig(
        id="example",
        path=root,
        include=("**/*.xlsx",),
        note_root=root / "notes",
    )


def test_metadata_mapping_round_trip(tmp_path: Path) -> None:
    assert metadata_to_nouns("Engineer, Myself", "Excel; Catalog") == [
        "[[Engineer, Myself]]",
        "[[Excel]]",
        "[[Catalog]]",
    ]
    note_path = tmp_path / "note.md"
    note_path.write_text(
        "---\ntype: Excel\ntitle: T\ndescription: D\nnouns:\n  - '[[Engineer, Myself]]'\n  - '[[Excel]]'\nsourceFileName: book.xlsx\n---\n",
        encoding="utf-8",
    )
    assert note_metadata(read_note(note_path))["category"] == "Engineer; Myself"


def test_render_preserves_unknown_fields_and_handwritten_body(tmp_path: Path) -> None:
    workbook_path = create_workbook(tmp_path / "book.xlsx")
    source = config(tmp_path)
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    note_path.parent.mkdir()
    note_path.write_text(
        "---\ntype: Excel\ntitle: Old\ndescription: Old\nnouns: []\nsourceFileName: book.xlsx\ncustomField: keep-me\ndate: '2026-01-01T00:00:00+09:00'\nupdated: '2026-01-01T00:00:00+09:00'\nnoteId: fixed-id\n---\n\n# Manual title\n\nHandwritten text.\n\n## Workbook Path\n\n`old.xlsx`\n",
        encoding="utf-8",
    )
    existing = read_note(note_path)
    rendered = render_note(workbook, source, existing=existing)
    note_path.write_text(rendered, encoding="utf-8")
    updated = read_note(note_path)
    assert updated.frontmatter["customField"] == "keep-me"
    assert "Handwritten text." in updated.body
    assert "<!-- excel-catalog:begin workbook-path -->" in updated.body
    assert str(workbook_path) in updated.body
    assert render_note(workbook, source, existing=updated, touch_updated=False) == rendered

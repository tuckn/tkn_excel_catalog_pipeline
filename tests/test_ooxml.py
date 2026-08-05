from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from excel_catalog_pipeline.adapters.ooxml import (
    WorkbookError,
    discover_workbooks,
    inspect_workbook,
    read_core_properties,
    read_custom_properties,
    write_properties,
)
from excel_catalog_pipeline.models import SourceConfig

from .helpers import create_workbook


def source_config(root: Path) -> SourceConfig:
    return SourceConfig(
        id="example",
        path=root,
        include=("**/*.xlsx", "**/*.xlsm"),
        note_root=root / "notes",
    )


def test_inspect_reads_metadata_sheets_and_text(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=1000)
    assert info.core["title"] == "Example title"
    assert info.sheets == [
        {"name": "Data", "sheetId": "1", "state": "visible", "path": "xl/worksheets/sheet1.xml"}
    ]
    assert info.sheet_text["Data"] == ["Catalog entry", "42"]
    assert len(info.content_fingerprint) == 64


def test_write_preserves_vba_and_unrelated_entries(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsm", macro_enabled=True)
    with zipfile.ZipFile(path) as before:
        vba = before.read("xl/vbaProject.bin")
        styles = before.read("xl/styles.xml")
    backup = write_properties(
        path,
        core={"title": "Updated", "category": "Architecture"},
        custom={"TknExcelCatalogId": "11111111-1111-1111-1111-111111111111"},
        backup_dir=tmp_path / "backups",
    )
    assert backup.exists()
    with zipfile.ZipFile(path) as after:
        assert after.read("xl/vbaProject.bin") == vba
        assert after.read("xl/styles.xml") == styles
        assert read_core_properties(after)["title"] == "Updated"
        assert read_custom_properties(after)["TknExcelCatalogId"].startswith("1111")
        assert after.testzip() is None


def test_core_metadata_preserves_repeated_whitespace(tmp_path: Path) -> None:
    original_title = "Node.js  Promise result"
    original_description = "  Leading  and  repeated spaces  "
    path = create_workbook(
        tmp_path / "book.xlsx",
        title=original_title,
        description=original_description,
    )
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=1)
    assert info.core["title"] == original_title
    assert info.core["description"] == original_description

    updated_title = "Updated  title"
    updated_description = "  Updated  description  text  "
    write_properties(
        path,
        core={"title": updated_title, "description": updated_description},
        backup_dir=tmp_path / "backups",
    )
    with zipfile.ZipFile(path) as archive:
        properties = read_core_properties(archive)
    assert properties["title"] == updated_title
    assert properties["description"] == updated_description


def test_failed_write_leaves_source_unchanged(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    before = path.read_bytes()
    with pytest.raises(WorkbookError, match="Unsupported core property"):
        write_properties(
            path,
            core={"not-a-property": "value"},
            backup_dir=tmp_path / "backups",
        )
    assert path.read_bytes() == before


def test_non_zip_is_reported_as_read_error(tmp_path: Path) -> None:
    path = tmp_path / "broken.xlsx"
    path.write_text("not a zip", encoding="utf-8")
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=100)
    assert info.read_status == "read-error"


def test_discovery_is_non_recursive_by_default(tmp_path: Path) -> None:
    create_workbook(tmp_path / "root.xlsx")
    create_workbook(tmp_path / "2008" / "nested.xlsx")

    found = discover_workbooks(source_config(tmp_path), max_text_chars=1)

    assert [item.relative_path for item in found] == ["root.xlsx"]


def test_recursive_discovery_honors_relative_ignore_patterns(tmp_path: Path) -> None:
    create_workbook(tmp_path / "root.xlsx")
    create_workbook(tmp_path / "2008" / "included.xlsx")
    create_workbook(tmp_path / "archive" / "ignored.xlsx")
    create_workbook(tmp_path / "2009" / "private" / "ignored.xlsm", macro_enabled=True)
    source = SourceConfig(
        id="example",
        path=tmp_path,
        recursive=True,
        include=("*.xlsx", "*.xlsm"),
        ignore=("archive/**", "private"),
        note_root=tmp_path / "notes",
    )

    found = discover_workbooks(source, max_text_chars=1)

    assert [item.relative_path for item in found] == ["2008/included.xlsx", "root.xlsx"]

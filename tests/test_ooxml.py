from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from excel_catalog_pipeline.adapters.ooxml import (
    WorkbookError,
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

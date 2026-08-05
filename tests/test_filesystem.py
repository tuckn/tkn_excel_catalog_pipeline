from __future__ import annotations

from pathlib import Path

import pytest

from excel_catalog_pipeline.adapters.filesystem import (
    RenameError,
    validate_workbook_filename,
    validate_workbook_relative_path,
)


@pytest.mark.parametrize("name", ["bad?.xlsx", "CON.xlsx", "trailing .xlsx ", "convert.xlsm"])
def test_invalid_windows_rename_is_rejected(name: str) -> None:
    with pytest.raises(RenameError):
        validate_workbook_filename(Path("book.xlsx"), name)


def test_valid_rename_keeps_directory_and_extension() -> None:
    assert validate_workbook_filename(Path("folder/book.xlsx"), "renamed.xlsx") == Path(
        "folder/renamed.xlsx"
    )


def test_valid_relative_path_can_move_workbook_within_source(tmp_path: Path) -> None:
    current = tmp_path / "book.xlsx"
    assert (
        validate_workbook_relative_path(tmp_path, current, "2008/book.xlsx")
        == (tmp_path / "2008" / "book.xlsx").resolve()
    )


@pytest.mark.parametrize("requested", ["../book.xlsx", "2008/bad?.xlsx", "2008/book.xlsm"])
def test_invalid_relative_workbook_path_is_rejected(tmp_path: Path, requested: str) -> None:
    with pytest.raises(RenameError):
        validate_workbook_relative_path(tmp_path, tmp_path / "book.xlsx", requested)

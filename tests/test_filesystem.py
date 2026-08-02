from __future__ import annotations

from pathlib import Path

import pytest

from excel_catalog_pipeline.adapters.filesystem import RenameError, validate_workbook_filename


@pytest.mark.parametrize("name", ["bad?.xlsx", "CON.xlsx", "trailing .xlsx ", "convert.xlsm"])
def test_invalid_windows_rename_is_rejected(name: str) -> None:
    with pytest.raises(RenameError):
        validate_workbook_filename(Path("book.xlsx"), name)


def test_valid_rename_keeps_directory_and_extension() -> None:
    assert validate_workbook_filename(Path("folder/book.xlsx"), "renamed.xlsx") == Path(
        "folder/renamed.xlsx"
    )

from __future__ import annotations

import pytest

from excel_catalog_pipeline.note_layout import patch_frontmatter


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
@pytest.mark.parametrize("full_path", [r"C:\path\to\book.xlsx", r"C:\path\O'Brien 日本語.xlsx"])
def test_patch_frontmatter_quotes_path_and_preserves_comments(full_path, newline):
    import yaml

    original = newline.join(
        [
            "---",
            "schemaVersion: '2.1'",
            'description: ""',
            "sourceFullPath: old.xlsx # keep",
            "unknown: preserved",
            "---",
            "",
            "Body.",
            "",
        ]
    )
    result = patch_frontmatter(original, {"sourceFullPath": full_path})
    expected = "sourceFullPath: '" + full_path.replace("'", "''") + "'"
    assert result == original.replace("sourceFullPath: old.xlsx", expected)
    assert yaml.safe_load(result.split("---")[1])["sourceFullPath"] == full_path

from __future__ import annotations

import pytest

from excel_catalog_pipeline.note_layout import migrate_layout


def test_migration_does_not_delete_code_or_path_annotations():
    text = "---\nschemaVersion: '2.0'\ndescription: ''\ncomments: ''\n---\n\n# Title\n\n```markdown\n## Overview\nExample\n```\n\n## Workbook Path\n\n`old.xlsx`\n\nCustom path note.\n\n## Next\n\nKeep.\n"
    result = migrate_layout(text, "new.xlsx")
    assert "```markdown\n## Overview\nExample\n```" in result
    assert "Custom path note." in result and "## Next" in result
    assert "`old.xlsx`" not in result
    assert migrate_layout(result, "new.xlsx") == result


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
@pytest.mark.parametrize("full_path", [r"C:\path\to\book.xlsx", r"C:\path\O'Brien 日本語.xlsx"])
def test_migration_quotes_source_path(full_path, newline):
    import yaml

    from excel_catalog_pipeline.note_layout import patch_frontmatter

    original = newline.join([
        "---", "schemaVersion: '2.1'", 'description: ""',
        "sourceFullPath: old.xlsx # keep", "unknown: preserved", "---", "", "Body.", "",
    ])
    result = patch_frontmatter(original, {"sourceFullPath": full_path})
    expected = "sourceFullPath: '" + full_path.replace("'", "''") + "'"
    assert result == original.replace("sourceFullPath: old.xlsx", expected)
    assert yaml.safe_load(result.split("---")[1])["sourceFullPath"] == full_path
    migrated = migrate_layout(original, full_path)
    assert expected + " # keep" in migrated
    assert yaml.safe_load(migrated.split("---")[1])["sourceFullPath"] == full_path
    assert migrate_layout(migrated, full_path) == migrated

from __future__ import annotations

from excel_catalog_pipeline.sync import compare_metadata, direction_fields


def test_three_way_directions() -> None:
    base = {
        "title": "A",
        "description": "A",
        "category": "A",
        "keywords": "A",
        "sourceFileName": "a.xlsx",
    }
    source = {
        "title": "B",
        "description": "A",
        "category": "C",
        "keywords": "D",
        "sourceFileName": "a.xlsx",
    }
    note = {
        "title": "A",
        "description": "B",
        "category": "C",
        "keywords": "E",
        "sourceFileName": "a.xlsx",
    }
    decisions = compare_metadata(base, source, note)
    assert direction_fields(decisions, "pull") == ["title"]
    assert direction_fields(decisions, "push") == ["description"]
    assert direction_fields(decisions, "converged") == ["category"]
    assert direction_fields(decisions, "conflict") == ["keywords"]

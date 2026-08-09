from __future__ import annotations

from excel_catalog_pipeline.sync import compare_metadata, direction_fields, resolve_for_pull


def test_three_way_directions() -> None:
    base = {
        "title": "A",
        "subject": "A",
        "author": "A",
        "keywords": "A",
        "categories": "A",
        "comments": "A",
        "sourceFileName": "a.xlsx",
    }
    source = {
        "title": "B",
        "subject": "A",
        "author": "A",
        "keywords": "D",
        "categories": "C",
        "comments": "A",
        "sourceFileName": "a.xlsx",
    }
    note = {
        "title": "A",
        "subject": "A",
        "author": "A",
        "keywords": "E",
        "categories": "C",
        "comments": "B",
        "sourceFileName": "a.xlsx",
    }
    decisions = compare_metadata(base, source, note)
    assert direction_fields(decisions, "pull") == ["title"]
    assert direction_fields(decisions, "push") == ["comments"]
    assert direction_fields(decisions, "converged") == ["categories"]
    assert direction_fields(decisions, "conflict") == ["keywords"]

    resolved = resolve_for_pull(decisions, preference="source")
    assert resolved == source

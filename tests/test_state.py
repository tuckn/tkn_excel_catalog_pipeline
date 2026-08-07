from __future__ import annotations

import json
from pathlib import Path

from excel_catalog_pipeline.state import STATE_VERSION, load_state


def test_load_state_migrates_v1_metadata_field_names(tmp_path: Path) -> None:
    path = tmp_path / "state.json"
    path.write_text(
        json.dumps(
            {
                "schemaVersion": 1,
                "entries": {
                    "source:example:book.xlsx": {
                        "baseMetadata": {
                            "title": "T",
                            "description": "D",
                            "category": "C",
                            "keywords": "K",
                            "sourceFileName": "book.xlsx",
                        }
                    }
                },
            }
        ),
        encoding="utf-8",
    )

    migrated = load_state(path)

    assert migrated["schemaVersion"] == STATE_VERSION == 2
    assert migrated["entries"]["source:example:book.xlsx"]["baseMetadata"] == {
        "title": "T",
        "subject": "",
        "author": "",
        "keywords": "K",
        "categories": "C",
        "comments": "D",
        "sourceFileName": "book.xlsx",
    }
    assert json.loads(path.read_text(encoding="utf-8"))["schemaVersion"] == 1

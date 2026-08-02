"""Versioned synchronization base state with atomic persistence."""

from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Any

from .models import ProxyNote, WorkbookInfo

STATE_VERSION = 1


class StateError(ValueError):
    """Persistent state is malformed or newer than this application understands."""


def empty_state() -> dict[str, Any]:
    return {"schemaVersion": STATE_VERSION, "entries": {}}


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return empty_state()
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise StateError(f"Failed to read sync state {path}: {exc}") from exc
    if not isinstance(data, dict) or data.get("schemaVersion") != STATE_VERSION:
        raise StateError(
            f"Unsupported sync state schema: {data.get('schemaVersion') if isinstance(data, dict) else None!r}"
        )
    if not isinstance(data.get("entries"), dict):
        raise StateError("sync state entries must be a mapping")
    return data


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        temp.write_text(
            json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
            newline="\n",
        )
        temp.replace(path)
    except Exception:
        temp.unlink(missing_ok=True)
        raise


def state_key(workbook: WorkbookInfo) -> str:
    if workbook.workbook_id:
        return f"workbook:{workbook.workbook_id}"
    return f"source:{workbook.legacy_source_id}"


def find_entry(
    state: dict[str, Any], workbook: WorkbookInfo
) -> tuple[str, dict[str, Any] | None, str]:
    entries: dict[str, Any] = state["entries"]
    direct = state_key(workbook)
    value = entries.get(direct)
    if isinstance(value, dict):
        return direct, value, "identity"
    path_matches = [
        (key, item)
        for key, item in entries.items()
        if isinstance(item, dict)
        and item.get("sourceRootId") == workbook.source_root_id
        and str(item.get("currentPath", "")).casefold() == workbook.relative_path.casefold()
    ]
    if len(path_matches) == 1:
        return path_matches[0][0], path_matches[0][1], "path"
    if workbook.content_fingerprint:
        signature_matches = [
            (key, item)
            for key, item in entries.items()
            if isinstance(item, dict)
            and item.get("sourceRootId") == workbook.source_root_id
            and item.get("contentFingerprint") == workbook.content_fingerprint
        ]
        if len(signature_matches) == 1:
            return signature_matches[0][0], signature_matches[0][1], "fingerprint"
    return direct, None, "none"


def make_entry(
    workbook: WorkbookInfo,
    note: ProxyNote,
    metadata: dict[str, str],
    *,
    now: str,
    result: str,
    previous_paths: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "sourceRootId": workbook.source_root_id,
        "workbookId": workbook.workbook_id,
        "currentPath": workbook.relative_path,
        "previousPaths": previous_paths or [],
        "notePath": str(note.path),
        "noteId": note.note_id,
        "baseMetadata": dict(metadata),
        "contentFingerprint": workbook.content_fingerprint,
        "sourceSignature": {
            "sizeBytes": workbook.size_bytes,
            "modified": workbook.modified,
        },
        "lastScan": now,
        "lastPull": now if result == "pulled" else "",
        "lastPush": now if result == "pushed" else "",
        "lastResult": result,
        "conflictFields": [],
    }


def replace_entry(
    state: dict[str, Any],
    *,
    old_key: str,
    workbook: WorkbookInfo,
    entry: dict[str, Any],
) -> None:
    entries: dict[str, Any] = state["entries"]
    new_key = state_key(workbook)
    if old_key != new_key:
        entries.pop(old_key, None)
    entries[new_key] = entry

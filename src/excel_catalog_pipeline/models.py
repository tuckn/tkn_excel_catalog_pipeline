"""Domain models shared by inventory, note, and synchronization code."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

METADATA_FIELDS = (
    "title",
    "subject",
    "author",
    "keywords",
    "categories",
    "comments",
    "sourceFileName",
)
CORE_PROPERTY_BY_METADATA_FIELD = {
    "title": "title",
    "subject": "subject",
    "author": "creator",
    "keywords": "keywords",
    "categories": "category",
    "comments": "description",
}
Direction = Literal[
    "unchanged",
    "pull",
    "push",
    "converged",
    "conflict",
]


def _canonical_terms(value: str) -> str:
    normalized = re.sub(r"\s+", " ", value).strip()
    if not normalized:
        return ""
    parts = re.split(r"[;\r\n]+", normalized)
    return "; ".join(dict.fromkeys(part.strip() for part in parts if part.strip()))


@dataclass(frozen=True)
class SourceConfig:
    id: str
    path: Path
    include: tuple[str, ...]
    note_root: Path
    recursive: bool = False
    ignore: tuple[str, ...] = ()
    profile: str = "tkn-obsidian-v1"
    rename_adapter: str = "report-only"


@dataclass(frozen=True)
class SyncConfig:
    pull_preserves_user_metadata: bool = True
    delete_missing_notes: bool = False
    delete_missing_workbooks: bool = False
    allow_source_rename: bool = False
    max_extracted_text_chars: int = 12000


@dataclass(frozen=True)
class AppConfig:
    schema_version: int
    sources: tuple[SourceConfig, ...]
    sync: SyncConfig
    loaded_files: tuple[Path, ...] = ()


@dataclass
class WorkbookInfo:
    path: Path
    relative_path: str
    source_root_id: str
    extension: str
    size_bytes: int
    modified: str
    core: dict[str, str] = field(default_factory=dict)
    custom: dict[str, str] = field(default_factory=dict)
    sheets: list[dict[str, str]] = field(default_factory=list)
    sheet_text: dict[str, list[str]] = field(default_factory=dict)
    content_fingerprint: str = ""
    read_status: str = "ok"
    warnings: list[str] = field(default_factory=list)

    @property
    def workbook_id(self) -> str:
        return self.custom.get("TknExcelCatalogId", "")

    @property
    def legacy_source_id(self) -> str:
        return f"{self.source_root_id}:{self.relative_path}"

    def metadata(self) -> dict[str, str]:
        return {
            "title": self.core.get("title", ""),
            "subject": self.core.get("subject", ""),
            "author": self.core.get("creator", ""),
            "keywords": _canonical_terms(self.core.get("keywords", "")),
            "categories": _canonical_terms(self.core.get("category", "")),
            "comments": self.core.get("description", ""),
            "sourceFileName": self.relative_path,
        }


@dataclass
class ProxyNote:
    path: Path
    frontmatter: dict[str, Any]
    body: str

    @property
    def note_id(self) -> str:
        return str(self.frontmatter.get("noteId", "")).strip()

    @property
    def source_id(self) -> str:
        return str(self.frontmatter.get("sourceId", "")).strip()


@dataclass(frozen=True)
class FieldDecision:
    field: str
    base: str
    source: str
    note: str
    direction: Direction


@dataclass
class Action:
    status: str
    source_root_id: str
    source_path: str = ""
    note_path: str = ""
    workbook_id: str = ""
    changed_fields: list[str] = field(default_factory=list)
    source_to_note_fields: list[str] = field(default_factory=list)
    note_to_source_fields: list[str] = field(default_factory=list)
    conflict_fields: list[str] = field(default_factory=list)
    message: str = ""
    details: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status,
            "sourceRoot": self.source_root_id,
            "sourcePath": self.source_path,
            "notePath": self.note_path,
            "workbookId": self.workbook_id,
            "changedFields": self.changed_fields,
            "sourceToNoteFields": self.source_to_note_fields,
            "noteToSourceFields": self.note_to_source_fields,
            "conflictFields": self.conflict_fields,
            "message": self.message,
            "details": self.details,
        }

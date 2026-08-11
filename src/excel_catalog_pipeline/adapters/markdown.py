"""Proxy-note parsing and managed-section rendering."""

from __future__ import annotations

import hashlib
import json
import re
import uuid
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

from ..discovery import matches_any, normalize_relative_path
from ..models import ProxyNote, SourceConfig, WorkbookInfo
from ..note_resources import (
    ManagedBlock,
    NoteResourceError,
    load_note_template,
    managed_blocks,
)

WINDOWS_PATH_PATTERN = re.compile(r"^[A-Za-z]:[\\/]")
YAML_TIMESTAMP_TAG = "tag:yaml.org,2002:timestamp"
OBSOLETE_FRONTMATTER_FIELDS = {"noun", "nouns"}


class FrontmatterLoader(yaml.SafeLoader):
    """Load plain ISO timestamps as strings so proxy-note values remain stable."""


class FrontmatterDumper(yaml.SafeDumper):
    """Emit ISO timestamp strings without adding YAML quotes."""


def _without_timestamp_resolver(
    resolvers: dict[str | None, list[tuple[str, re.Pattern[str]]]],
) -> dict[str | None, list[tuple[str, re.Pattern[str]]]]:
    return {
        key: [(tag, pattern) for tag, pattern in values if tag != YAML_TIMESTAMP_TAG]
        for key, values in resolvers.items()
    }


FrontmatterLoader.yaml_implicit_resolvers = _without_timestamp_resolver(
    FrontmatterLoader.yaml_implicit_resolvers
)
FrontmatterDumper.yaml_implicit_resolvers = _without_timestamp_resolver(
    FrontmatterDumper.yaml_implicit_resolvers
)


class NoteError(ValueError):
    """A proxy note is malformed or unsafe to update."""


def now_iso() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def read_note(path: Path) -> ProxyNote:
    text = path.read_text(encoding="utf-8-sig")
    if not text.startswith("---"):
        return ProxyNote(path=path, frontmatter={}, body=text)
    match = re.match(r"\A---\s*\r?\n(.*?)\r?\n---\s*\r?\n?", text, re.DOTALL)
    if not match:
        raise NoteError(f"Invalid Frontmatter fence: {path}")
    try:
        loaded = yaml.load(match.group(1), Loader=FrontmatterLoader) or {}
    except yaml.YAMLError as exc:
        raise NoteError(f"Invalid Frontmatter YAML in {path}: {exc}") from exc
    if not isinstance(loaded, dict):
        raise NoteError(f"Frontmatter must be a mapping: {path}")
    return ProxyNote(path=path, frontmatter=loaded, body=text[match.end() :])


def discover_notes(source: SourceConfig) -> list[ProxyNote]:
    root = source.note_root
    if not root.exists():
        return []
    notes: list[ProxyNote] = []
    candidates = root.rglob("*.md") if source.recursive else root.glob("*.md")
    for path in sorted(candidates, key=lambda item: item.as_posix().casefold()):
        note = read_note(path)
        if str(note.frontmatter.get("type", "")).casefold() == "excel" or (
            note.frontmatter.get("fileKind") == "excelWorkbook"
        ):
            recorded = normalize_relative_path(str(note.frontmatter.get("sourceFileName", "")))
            note_relative = path.relative_to(root).as_posix()
            inferred = note_relative[:-3] if note_relative.casefold().endswith(".md") else ""
            if source.ignore and (
                matches_any(recorded, source.ignore) or matches_any(inferred, source.ignore)
            ):
                continue
            notes.append(note)
    return notes


def _frontmatter_list(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str) and value.strip():
        return [value.strip()]
    return []


def _frontmatter_string(value: Any) -> str:
    return "" if value is None else str(value)


def _link_target(value: str) -> str:
    text = value.strip()
    if text.startswith("[[") and text.endswith("]]"):
        text = text[2:-2]
    if "|" in text:
        text = text.split("|", 1)[0]
    return re.sub(r"\s+", " ", text).strip()


def _dedupe(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value and value.casefold() not in seen:
            result.append(value)
            seen.add(value.casefold())
    return result


def _split_terms(value: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", value).strip()
    if not normalized:
        return []
    parts = re.split(r"[;\r\n]+", normalized)
    return _dedupe(part.strip() for part in parts if part.strip())


def metadata_to_links(value: str) -> list[str]:
    return [f"[[{term}]]" for term in _split_terms(value)]


def metadata_to_frontmatter_terms(value: str, term_format: str) -> list[str]:
    terms = _split_terms(value)
    if term_format == "plain":
        return terms
    return [f"[[{term}]]" for term in terms]


def _frontmatter_terms(value: Any) -> str:
    terms: list[str] = []
    for item in _frontmatter_list(value):
        terms.extend(_split_terms(_link_target(item)))
    return "; ".join(_dedupe(terms))


def note_metadata(note: ProxyNote) -> dict[str, str]:
    schema_version = _frontmatter_string(note.frontmatter.get("schemaVersion"))
    uses_v2_metadata = schema_version == "2.0" or any(
        field in note.frontmatter for field in ("subject", "author", "categories", "comments")
    )
    if uses_v2_metadata:
        categories = _frontmatter_terms(note.frontmatter.get("categories"))
        keywords = _frontmatter_terms(note.frontmatter.get("keywords"))
    else:
        nouns = _frontmatter_list(note.frontmatter.get("nouns"))
        legacy = str(note.frontmatter.get("noun", "")).strip()
        if legacy:
            nouns = _dedupe([legacy, *nouns])
        targets = [_link_target(item) for item in nouns]
        categories = _frontmatter_terms(targets[:1])
        keywords = _frontmatter_terms(targets[1:])
    return {
        "title": _frontmatter_string(note.frontmatter.get("title")),
        "subject": _frontmatter_string(note.frontmatter.get("subject")),
        "author": _frontmatter_string(note.frontmatter.get("author")),
        "keywords": keywords,
        "categories": categories,
        "comments": _frontmatter_string(
            note.frontmatter.get("comments", note.frontmatter.get("description", ""))
        ),
        "sourceFileName": _frontmatter_string(note.frontmatter.get("sourceFileName")),
    }


def workbook_path(note: ProxyNote) -> Path | None:
    legacy = str(note.frontmatter.get("sourceFullPath", "")).strip()
    if legacy:
        return Path(legacy)
    marker = _marker_pattern("workbook-path")
    match = marker.search(note.body)
    section = match.group("content") if match else ""
    if not section:
        legacy_match = re.search(
            r"^##\s+(?:Workbook Path|ファイルを開く)\s*$\r?\n(?P<content>.*?)(?=^##\s+|\Z)",
            note.body,
            re.MULTILINE | re.DOTALL,
        )
        section = legacy_match.group("content") if legacy_match else ""
    code = re.search(r"`([^`\r\n]+)`", section)
    candidate = code.group(1).strip() if code else ""
    if not candidate:
        candidate = next((line.strip(" `") for line in section.splitlines() if line.strip()), "")
    return Path(candidate) if candidate else None


def frontmatter_hash(note: ProxyNote) -> str:
    payload = json.dumps(note_metadata(note), ensure_ascii=False, sort_keys=True).encode()
    return hashlib.sha256(payload).hexdigest()


def _marker_pattern(name: str) -> re.Pattern[str]:
    return re.compile(
        rf"<!-- excel-catalog:begin {re.escape(name)} -->\r?\n"
        rf"(?P<content>.*?)"
        rf"<!-- excel-catalog:end {re.escape(name)} -->",
        re.DOTALL,
    )


def _replace_managed(body: str, block: ManagedBlock) -> str:
    marker = _marker_pattern(block.name)
    if marker.search(body):
        return marker.sub(lambda _match: block.text, body, count=1)
    headings = [re.escape(block.heading)]
    if block.name == "workbook-path":
        headings.append(re.escape("ファイルを開く"))
    legacy = re.compile(
        rf"^##\s+(?:{'|'.join(headings)})\s*$\r?\n.*?(?=^##\s+|\Z)",
        re.MULTILINE | re.DOTALL,
    )
    if legacy.search(body):
        return legacy.sub(lambda _match: block.text + "\n\n", body, count=1)
    separator = "" if not body or body.endswith("\n\n") else "\n\n"
    return body + separator + block.text + "\n"


def _render_map(workbook: WorkbookInfo) -> str:
    if workbook.read_status != "ok":
        return f"- readStatus: {workbook.read_status}"
    if not workbook.sheets:
        return "- No sheet metadata extracted."
    return "\n".join(
        f"- {sheet.get('name', '')} (sheetId: {sheet.get('sheetId', '')}, "
        f"state: {sheet.get('state', 'visible')})"
        for sheet in workbook.sheets
    )


def _render_text(workbook: WorkbookInfo) -> str:
    if workbook.read_status != "ok":
        return f"Extraction unavailable: {workbook.read_status}."
    if not workbook.sheet_text:
        return "No text extracted."
    sections: list[str] = []
    for sheet_name, values in workbook.sheet_text.items():
        sections.extend([f"### {sheet_name}", "", *[f"- {value}" for value in values], ""])
    return "\n".join(sections).rstrip()


def note_filename(workbook: WorkbookInfo) -> str:
    invalid = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
    stem = invalid.sub(" ", workbook.path.stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .") or "Untitled Excel Workbook"
    return f"{stem[:150]}{workbook.extension}.md"


def note_path(workbook: WorkbookInfo, source: SourceConfig) -> Path:
    """Return the proxy-note path mirroring the workbook's relative parent folder."""
    parent = Path(workbook.relative_path).parent
    if parent == Path("."):
        return source.note_root / note_filename(workbook)
    return source.note_root / parent / note_filename(workbook)


def _merge_frontmatter(rendered: dict[str, Any], existing: dict[str, Any]) -> dict[str, Any]:
    """Apply the template contract while preserving unknown fields near their old position."""

    unknown_before: dict[str, list[tuple[str, Any]]] = {}
    pending: list[tuple[str, Any]] = []
    for key, value in existing.items():
        if key in OBSOLETE_FRONTMATTER_FIELDS:
            continue
        if key in rendered:
            if pending:
                unknown_before.setdefault(key, []).extend(pending)
                pending = []
        else:
            pending.append((key, value))

    merged: dict[str, Any] = {}
    for key, value in rendered.items():
        merged.update(unknown_before.get(key, ()))
        merged[key] = value
    merged.update(pending)
    return merged


def render_note(
    workbook: WorkbookInfo,
    source: SourceConfig,
    *,
    metadata: dict[str, str] | None = None,
    existing: ProxyNote | None = None,
    touch_updated: bool = True,
    refresh_source_properties: bool = True,
) -> str:
    timestamp = now_iso()
    values = metadata or workbook.metadata()
    existing_frontmatter = dict(existing.frontmatter) if existing else {}
    title = values["title"] or workbook.path.stem
    stable_part = workbook.workbook_id or workbook.relative_path

    try:
        template = load_note_template(source.profile)
        current_schema = _frontmatter_string(existing_frontmatter.get("schemaVersion"))
        description = (
            _frontmatter_string(existing_frontmatter.get("description"))
            if current_schema == template.schema_version
            else ""
        )
        source_created = (
            workbook.core.get("created", "")
            if refresh_source_properties
            else _frontmatter_string(existing_frontmatter.get("sourceCreated"))
        )
        source_modified = (
            workbook.core.get("modified", "")
            if refresh_source_properties
            else _frontmatter_string(existing_frontmatter.get("sourceModified"))
        )
        rendered_template = template.render(
            {
                "title": title,
                "description": description,
                "subject": values["subject"],
                "author": values["author"],
                "keywords": metadata_to_frontmatter_terms(
                    values["keywords"], source.frontmatter_term_format
                ),
                "categories": metadata_to_frontmatter_terms(
                    values["categories"], source.frontmatter_term_format
                ),
                "comments": values["comments"],
                "files": existing_frontmatter.get("files", []),
                "source_root": source.id,
                "source_file_name": values["sourceFileName"],
                "source_id": f"{source.id}:{stable_part}",
                "source_created": source_created,
                "source_modified": source_modified,
                "date": existing_frontmatter.get("date", timestamp),
                "updated": (
                    timestamp
                    if touch_updated or "updated" not in existing_frontmatter
                    else existing_frontmatter["updated"]
                ),
                "note_id": existing_frontmatter.get("noteId", str(uuid.uuid4())),
                "workbook_path": f"`{workbook.path}`",
                "workbook_map": _render_map(workbook),
                "extracted_text": _render_text(workbook),
            }
        )
        sections = managed_blocks(rendered_template.body, template)
    except NoteResourceError as exc:
        raise NoteError(str(exc)) from exc

    frontmatter = _merge_frontmatter(rendered_template.frontmatter, existing_frontmatter)

    if existing:
        body = _marker_pattern("excel-metadata").sub("", existing.body)
        body = re.sub(
            r"^##\s+Excel Metadata\s*$\r?\n.*?(?=^##\s+|\Z)",
            "",
            body,
            count=1,
            flags=re.MULTILINE | re.DOTALL,
        )
        for section in sections:
            body = _replace_managed(body, section)
    else:
        body = rendered_template.body.rstrip() + "\n"
    yaml_text = yaml.dump(
        frontmatter,
        Dumper=FrontmatterDumper,
        allow_unicode=True,
        sort_keys=False,
        default_flow_style=False,
        width=1000,
    )
    return f"---\n{yaml_text}---\n\n{body.lstrip()}"


def write_note_atomic(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        temp.write_text(content, encoding="utf-8", newline="\n")
        temp.replace(path)
    except Exception:
        temp.unlink(missing_ok=True)
        raise


def basename_collisions(
    vault_root: Path, note_name: str, *, target: Path | None = None
) -> list[Path]:
    stem = Path(note_name).stem.casefold()
    collisions: list[Path] = []
    if not vault_root.exists():
        return collisions
    for path in vault_root.rglob("*.md"):
        if target is not None and path.resolve() == target.resolve():
            continue
        if path.stem.casefold() == stem:
            collisions.append(path)
    return collisions


def find_obsidian_vault_root(note_root: Path) -> Path:
    """Find the nearest ancestor containing .obsidian, or conservatively use note_root."""
    resolved = note_root.resolve()
    for candidate in (resolved, *resolved.parents):
        if (candidate / ".obsidian").is_dir():
            return candidate
    return resolved

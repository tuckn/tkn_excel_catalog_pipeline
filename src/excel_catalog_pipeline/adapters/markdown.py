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

from ..models import ProxyNote, SourceConfig, WorkbookInfo
from ..note_resources import (
    ManagedBlock,
    NoteResourceError,
    load_note_template,
    managed_blocks,
)

WINDOWS_PATH_PATTERN = re.compile(r"^[A-Za-z]:[\\/]")


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
        loaded = yaml.safe_load(match.group(1)) or {}
    except yaml.YAMLError as exc:
        raise NoteError(f"Invalid Frontmatter YAML in {path}: {exc}") from exc
    if not isinstance(loaded, dict):
        raise NoteError(f"Frontmatter must be a mapping: {path}")
    return ProxyNote(path=path, frontmatter=loaded, body=text[match.end() :])


def discover_notes(root: Path) -> list[ProxyNote]:
    if not root.exists():
        return []
    notes: list[ProxyNote] = []
    for path in sorted(root.glob("*.md"), key=lambda item: item.name.casefold()):
        note = read_note(path)
        if str(note.frontmatter.get("type", "")).casefold() == "excel" or (
            note.frontmatter.get("fileKind") == "excelWorkbook"
        ):
            notes.append(note)
    return notes


def _frontmatter_list(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str) and value.strip():
        return [value.strip()]
    return []


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
    if len(parts) == 1:
        parts = re.split(r",+", normalized)
    return _dedupe(part.strip() for part in parts if part.strip())


def metadata_to_nouns(category: str, keywords: str) -> list[str]:
    category_terms = _split_terms(category)
    category_link = f"[[{', '.join(category_terms)}]]" if category_terms else ""
    keyword_links = [f"[[{term}]]" for term in _split_terms(keywords)]
    return _dedupe([category_link, *keyword_links])


def note_metadata(note: ProxyNote) -> dict[str, str]:
    nouns = _frontmatter_list(note.frontmatter.get("nouns"))
    legacy = str(note.frontmatter.get("noun", "")).strip()
    if legacy:
        nouns = _dedupe([legacy, *nouns])
    targets = [_link_target(item) for item in nouns]
    category = "; ".join(_split_terms(targets[0])) if targets else ""
    keywords = "; ".join(_dedupe(item for item in targets[1:] if item))
    return {
        "title": str(note.frontmatter.get("title", "")),
        "description": str(note.frontmatter.get("description", "")),
        "category": category,
        "keywords": keywords,
        "sourceFileName": str(note.frontmatter.get("sourceFileName", "")),
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


def _render_excel_metadata(workbook: WorkbookInfo) -> str:
    lines: list[str] = []
    if workbook.core:
        lines.extend(["### Core", ""])
        lines.extend(f"- {key}: {value}" for key, value in workbook.core.items())
    if workbook.custom:
        if lines:
            lines.append("")
        lines.extend(["### Custom", ""])
        lines.extend(f"- {key}: {value}" for key, value in workbook.custom.items())
    if workbook.warnings:
        if lines:
            lines.append("")
        lines.extend(["### Warnings", ""])
        lines.extend(f"- {warning}" for warning in workbook.warnings)
    return "\n".join(lines) if lines else "No Excel document metadata extracted."


def note_filename(workbook: WorkbookInfo) -> str:
    invalid = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
    stem = invalid.sub(" ", workbook.path.stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .") or "Untitled Excel Workbook"
    return f"{stem[:150]}{workbook.extension}.md"


def render_note(
    workbook: WorkbookInfo,
    source: SourceConfig,
    *,
    metadata: dict[str, str] | None = None,
    existing: ProxyNote | None = None,
    touch_updated: bool = True,
) -> str:
    timestamp = now_iso()
    values = metadata or workbook.metadata()
    frontmatter = dict(existing.frontmatter) if existing else {}
    frontmatter.pop("noun", None)
    frontmatter["type"] = "Excel"
    frontmatter["title"] = values["title"] or workbook.path.stem
    frontmatter["description"] = values["description"]
    frontmatter["nouns"] = metadata_to_nouns(values["category"], values["keywords"])
    frontmatter.setdefault("files", [])
    frontmatter["sourceRoot"] = source.id
    frontmatter["sourceFileName"] = values["sourceFileName"]
    stable_part = workbook.workbook_id or workbook.relative_path
    frontmatter["sourceId"] = f"{source.id}:{stable_part}"
    frontmatter.setdefault("date", timestamp)
    if touch_updated or "updated" not in frontmatter:
        frontmatter["updated"] = timestamp
    frontmatter.setdefault("noteId", str(uuid.uuid4()))

    try:
        template = load_note_template(source.profile)
        rendered_template = template.render(
            {
                "title": str(frontmatter["title"]),
                "description": str(frontmatter["description"]),
                "workbook_path": f"`{workbook.path}`",
                "workbook_map": _render_map(workbook),
                "extracted_text": _render_text(workbook),
                "excel_metadata": _render_excel_metadata(workbook),
            }
        )
        sections = managed_blocks(rendered_template, template)
    except NoteResourceError as exc:
        raise NoteError(str(exc)) from exc

    if existing:
        body = existing.body
        for section in sections:
            body = _replace_managed(body, section)
    else:
        body = rendered_template.rstrip() + "\n"
    yaml_text = yaml.safe_dump(
        frontmatter,
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

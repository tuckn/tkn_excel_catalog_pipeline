"""Migrate retired body sections without rewriting unrelated Frontmatter fields."""

from __future__ import annotations

import re
from typing import Any

import yaml

PLACEHOLDERS = {
    "Excel metadataから生成したExcel file entityの下書き。",
    "Excel workbookの検索・管理用代理ノート。",
}
FRONTMATTER = re.compile(r"\A(\ufeff?---[ \t]*\r?\n)(.*?)(\r?\n---[ \t]*\r?\n?)", re.DOTALL)


def _section_spans(body: str, names: set[str]) -> list[tuple[int, int, str]]:
    lines = body.splitlines(keepends=True)
    offsets = [0]
    for line in lines:
        offsets.append(offsets[-1] + len(line))
    boundaries: list[tuple[int, str]] = []
    fence = ""
    context = False
    for index, line in enumerate(lines):
        marker = re.match(r"^ {0,3}(`{3,}|~{3,})", line)
        if marker:
            run = marker.group(1)
            if not fence:
                fence = run
            elif run[0] == fence[0] and len(run) >= len(fence):
                fence = ""
            continue
        if fence:
            continue
        if line.startswith("<!-- excel-catalog:begin context-"):
            context = True
        if line.startswith("<!-- excel-catalog:end context-"):
            context = False
            continue
        if line.startswith("<!-- excel-catalog:begin "):
            boundaries.append((index, ""))
        heading = re.match(r"^(#{1,2})[ \t]+(.+?)[ \t]*\r?\n?$", line)
        if heading and not context:
            boundaries.append((index, heading.group(2) if len(heading.group(1)) == 2 else ""))
    result = []
    for pos, (index, name) in enumerate(boundaries):
        if name in names:
            end = boundaries[pos + 1][0] if pos + 1 < len(boundaries) else len(lines)
            result.append((offsets[index], offsets[end], "".join(lines[index + 1 : end]).strip()))
    return result


def retire_sections(body: str, description: str, full_path: str) -> tuple[str, str]:
    """Overview prose becomes description; keep annotations beside the old path."""
    overview = _section_spans(body, {"Overview"})
    additions = [text for _, _, text in overview if text and text not in PLACEHOLDERS]
    parts = [description] if description.strip() else []
    for text in additions:
        if text not in parts:
            parts.append(text)
    for start, end, _ in reversed(overview):
        body = body[:start] + body[end:]
    description = "\n\n".join(parts)
    managed = re.compile(
        r"<!-- excel-catalog:begin workbook-path -->\r?\n(.*?)<!-- excel-catalog:end workbook-path -->\r?\n?",
        re.DOTALL,
    )

    def path_notes(content: str) -> str:
        remaining = []
        for line in content.splitlines(keepends=True):
            stripped = line.strip()
            if re.match(r"^##\s+(?:Workbook Path|ファイルを開く)\s*$", stripped):
                continue
            candidate = stripped.strip("`")
            if candidate == full_path or (
                stripped.startswith("`")
                and stripped.endswith("`")
                and candidate.lower().endswith((".xlsx", ".xlsm", ".xls"))
            ):
                continue
            remaining.append(line)
        return "".join(remaining).lstrip("\r\n")

    body = managed.sub(lambda match: path_notes(match.group(1)), body)
    for start, end, text in reversed(_section_spans(body, {"Workbook Path", "ファイルを開く"})):
        body = body[:start] + path_notes(text) + body[end:]
    return body, description


def patch_frontmatter(text: str, values: dict[str, Any]) -> str:
    """Replace only selected YAML entries, preserving all other source text."""
    match = FRONTMATTER.match(text)
    if match is None:
        raise ValueError("A proxy note must have YAML Frontmatter")
    raw = match.group(2)
    root = yaml.compose(raw)
    if not isinstance(root, yaml.MappingNode):
        raise ValueError("Frontmatter must be a mapping")
    entries = {key.value: (key, value) for key, value in root.value}
    if len(entries) != len(root.value):
        raise ValueError("Duplicate Frontmatter keys are not supported")
    newline = "\r\n" if "\r\n" in text else "\n"
    changes = []
    additions = []
    for name, value in values.items():
        rendered = (
            yaml.safe_dump({name: value}, allow_unicode=True, sort_keys=False, width=1000)
            .rstrip("\n")
            .replace("\n", newline)
        )
        if name in entries:
            key, node = entries[name]
            start, end = key.start_mark.index, node.end_mark.index
            if raw[start:end].endswith("\n"):
                rendered += newline
            changes.append((start, end, rendered))
        else:
            additions.append(rendered)
    for start, end, rendered in sorted(changes, reverse=True):
        raw = raw[:start] + rendered + raw[end:]
    if additions:
        raw = raw.rstrip("\r\n") + newline + newline.join(additions)
    return match.group(1) + raw + match.group(3) + text[match.end() :]


def migrate_layout(text: str, full_path: str, *, description: str | None = None) -> str:
    match = FRONTMATTER.match(text)
    if match is None:
        raise ValueError("A proxy note must have YAML Frontmatter")
    metadata = yaml.safe_load(match.group(2))
    if not isinstance(metadata, dict):
        raise ValueError("Frontmatter must be a mapping")
    current = str(metadata.get("description") or "")
    if not str(metadata.get("schemaVersion", "")).startswith("2.") and "comments" not in metadata:
        raise ValueError("Run pull to migrate legacy metadata before importing context")
    body, summary = retire_sections(text[match.end() :], current, full_path)
    if description is not None:
        summary = description
    migrated = text[: match.end()] + body
    return patch_frontmatter(
        migrated, {"description": summary, "sourceFullPath": full_path, "schemaVersion": "2.1"}
    )

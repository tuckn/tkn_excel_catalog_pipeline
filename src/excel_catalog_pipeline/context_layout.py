"""Arrange managed context without rewriting text outside its markers."""

from __future__ import annotations

import re
from typing import Any

from .context_profiles import ContextProfile
from .context_source import ContextError

BLOCK = re.compile(
    r"<!-- excel-catalog:begin (?P<name>context-\d+|context-workbook|sheet-contexts|workbook-map|extracted-text) -->\r?\n"
    r".*?<!-- excel-catalog:end (?P=name) -->",
    re.DOTALL,
)


def sheet_heading(name: str) -> str:
    return name.replace("\r", " ").replace("\n", " ")


def migrate_legacy_block(
    block: str, sheet: dict[str, str], state: dict[str, Any], profile: ContextProfile
) -> str:
    """Mechanically nest a verified old block; never summarize or discard its prose."""
    if state.get("layoutVersion") == 1:
        return block
    newline = "\r\n" if "\r\n" in block else "\n"
    lines = block.replace("\r\n", "\n").splitlines()
    if len(lines) < 3 or not re.fullmatch(r"## .+ \(sheetId: \d+\)", lines[1]):
        return block
    body = lines[2:-1]
    fence = ""
    nested = []
    for line in body:
        marker = re.match(r"^ {0,3}(\x60{3,}|~{3,})", line)
        if marker:
            run = marker.group(1)
            if not fence:
                fence = run
            elif run[0] == fence[0] and len(run) >= len(fence):
                fence = ""
        elif not fence:
            heading = re.match(r"^(#{1,6}) (.*)", line)
            if heading:
                depth = len(heading.group(1)) + 2
                line = (
                    "#" * depth + " " + heading.group(2)
                    if depth <= 6
                    else "**" + heading.group(2) + "**"
                )
        nested.append(line)
    recorded_profile = state.get("promptProfile") or profile.labels["unknown_profile"]
    content_heading = profile.labels["content"]
    result = [
        lines[0],
        "### " + sheet_heading(sheet["name"]),
        "",
        profile.labels["profile"] + ": " + str(recorded_profile),
        "",
        profile.labels["legacy"],
        "",
        "#### " + content_heading,
        "",
        *nested,
        lines[-1],
    ]
    return newline.join(result)


def arrange_context(text: str, sheets: list[dict[str, str]], profile: ContextProfile) -> str:
    """Reassign managed blocks to their existing slots, preserving every other byte."""
    matches = list(BLOCK.finditer(text))
    by_name = {match["name"]: match.group() for match in matches}
    if len(by_name) != len(matches):
        raise ContextError("Duplicate managed body markers; repair the note before rebuilding")
    if "context-workbook" not in by_name:
        return text
    newline = "\r\n" if "\r\n" in text else "\n"
    by_name["sheet-contexts"] = (
        "<!-- excel-catalog:begin sheet-contexts -->"
        + newline
        + "## "
        + profile.labels["sheet_heading"]
        + newline
        + "<!-- excel-catalog:end sheet-contexts -->"
    )
    order = ["context-workbook", "sheet-contexts"]
    order += ["context-" + sheet["id"] for sheet in sheets]
    # Preserve any retired sections kept by the caller, after current sheets.
    order += [name for name in by_name if name.startswith("context-") and name not in order]
    order += ["extracted-text", "workbook-map"]
    blocks = [by_name[name] for name in order if name in by_name]
    # A new Sheets header shares the first slot. Later runs have equal slot counts.
    extra = len(blocks) - len(matches)
    replacements = [(newline * 2).join(blocks[: extra + 1])] + blocks[extra + 1 :]
    result = []
    offset = 0
    for match, replacement in zip(matches, replacements, strict=True):
        result.extend([text[offset : match.start()], replacement])
        offset = match.end()
    result.append(text[offset:])
    return "".join(result)

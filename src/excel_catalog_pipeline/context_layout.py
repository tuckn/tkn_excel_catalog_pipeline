"""Arrange managed context without rewriting text outside its markers."""

from __future__ import annotations

import re
from typing import Any

from .context_profiles import ContextProfile
from .context_source import ContextError
from .sheet_layout import arrange_sheets


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
    """Group each sheet's context and extraction without changing its context hash."""
    try:
        return arrange_sheets(text, sheets, heading=profile.labels["sheet_heading"])
    except ValueError as exc:
        raise ContextError(str(exc)) from exc

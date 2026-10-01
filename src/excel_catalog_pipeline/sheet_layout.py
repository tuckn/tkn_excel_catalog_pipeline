"""Compose sheet-local generated context and independently refreshed extraction."""

from __future__ import annotations

import re

BLOCK = re.compile(
    r"<!-- excel-catalog:begin (?P<name>context-\d+|context-workbook|sheet-contexts|"
    r"sheet-heading-\d+|sheet-text-\d+|workbook-map|extracted-text) -->\r?\n"
    r".*?<!-- excel-catalog:end (?P=name) -->",
    re.DOTALL,
)


def sheet_id(sheet: dict[str, str]) -> str:
    return sheet.get("id") or sheet["sheetId"]


def managed(name: str, body: str, newline: str = "\n") -> str:
    return (
        f"<!-- excel-catalog:begin {name} -->\n{body.rstrip()}\n<!-- excel-catalog:end {name} -->"
    ).replace("\n", newline)


def arrange_sheets(
    text: str,
    sheets: list[dict[str, str]],
    *,
    extracted: dict[str, str] | None = None,
    heading: str | None = None,
) -> str:
    """Keep context bytes and all unmanaged text intact while arranging managed slots."""
    matches = list(BLOCK.finditer(text))
    blocks = {match["name"]: match.group() for match in matches}
    if len(blocks) != len(matches):
        raise ValueError("Duplicate managed body markers; repair the note before rebuilding")
    newline = "\r\n" if "\r\n" in text else "\n"
    if heading is not None or "sheet-contexts" not in blocks:
        blocks["sheet-contexts"] = managed("sheet-contexts", "## " + (heading or "シート"), newline)
    if extracted is not None:
        # The old workbook-wide extraction is replaced only by fresh source extraction.
        blocks.pop("extracted-text", None)
        blocks = {key: value for key, value in blocks.items() if not key.startswith("sheet-text-")}
        for key, body in extracted.items():
            blocks["sheet-text-" + key] = managed(
                "sheet-text-" + key, "#### Extracted Text\n\n" + body, newline
            )
    order = ["workbook-map", "context-workbook", "sheet-contexts"]
    active = {sheet_id(sheet) for sheet in sheets}
    for sheet in sheets:
        key = sheet_id(sheet)
        header = "sheet-heading-" + key
        context = blocks.get("context-" + key, "")
        # Prior context owns its sheet heading. Retain it verbatim for hash/review protection.
        has_heading = re.match(r"[^\n]*\r?\n#{2,3} ", context) is not None
        if has_heading:
            blocks.pop(header, None)
        else:
            name = sheet["name"].replace("\r", " ").replace("\n", " ")
            blocks[header] = managed(header, "### " + name, newline)
        order.extend([header, "context-" + key, "sheet-text-" + key])
    # Retired context stays with its old sheet heading; only the context owner may remove it.
    for name in blocks:
        match = re.fullmatch(r"context-(\d+)", name)
        if match and match[1] not in active:
            order.extend(["sheet-heading-" + match[1], name, "sheet-text-" + match[1]])
    order.append("extracted-text")
    arranged = [blocks[name] for name in order if name in blocks]
    if not matches:
        return text.rstrip() + newline * 2 + (newline * 2).join(arranged) + newline
    # Place additional blocks together in the first slot; preserve bytes between every slot.
    extra = max(0, len(arranged) - len(matches))
    replacements = [(newline * 2).join(arranged[: extra + 1])] + arranged[extra + 1 :]
    replacements.extend([""] * (len(matches) - len(replacements)))
    result = []
    offset = 0
    for match, replacement in zip(matches, replacements, strict=True):
        result.extend([text[offset : match.start()], replacement])
        offset = match.end()
    result.append(text[offset:])
    return "".join(result)

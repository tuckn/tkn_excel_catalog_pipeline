"""Cover option validation shared by YAML, CLI and saved generation settings."""

from __future__ import annotations

import re
from dataclasses import asdict
from typing import Any

from .models import CoverConfig


def normalize_range(value: Any) -> str:
    if not isinstance(value, str):
        raise ValueError("cover.range must be a rectangular A1 range, for example A1:Q50")
    match = re.fullmatch(
        r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}):\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})",
        value.strip(),
    )
    if match is None:
        raise ValueError("cover.range must be a rectangular A1 range, for example A1:Q50")
    first_col, first_row, last_col, last_row = match.groups()

    def column(text: str) -> int:
        result = 0
        for char in text.upper():
            result = result * 26 + ord(char) - ord("A") + 1
        return result

    if not (
        column(first_col) <= column(last_col) <= 16384
        and int(first_row) <= int(last_row) <= 1048576
    ):
        raise ValueError("cover.range must be ordered and within Excel worksheet limits")
    return f"{first_col.upper()}{first_row}:{last_col.upper()}{last_row}"


def validate_cover(raw: Any) -> CoverConfig:
    if not isinstance(raw, dict):
        raise ValueError("cover must be a mapping")
    unknown = set(raw) - set(asdict(CoverConfig()))
    if unknown:
        raise ValueError("Unknown cover key(s): " + ", ".join(sorted(unknown)))
    values = {**asdict(CoverConfig()), **raw}
    if values["mode"] not in ("auto", "embedded", "sheet"):
        raise ValueError("cover.mode must be auto, embedded, or sheet")
    if values["sheet"] is not None and (
        not isinstance(values["sheet"], str) or not values["sheet"].strip()
    ):
        raise ValueError("cover.sheet must be a non-empty sheet name or null")
    if values["range"] is not None:
        values["range"] = normalize_range(values["range"])
    if values["width"] is not None and (
        type(values["width"]) is not int or not 600 <= values["width"] <= 4000
    ):
        raise ValueError("cover.width must be an integer between 600 and 4000 or null")
    return CoverConfig(**values)

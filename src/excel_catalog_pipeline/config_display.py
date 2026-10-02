"""Copyable, line-oriented rendering of resolved configuration data."""

from __future__ import annotations

import json
from typing import Any


def config_leaves(value: Any, prefix: str = "") -> dict[str, Any]:
    """Flatten mappings and sequences while retaining empty containers."""
    if isinstance(value, dict) and value:
        return {
            path: leaf
            for key, item in value.items()
            for path, leaf in config_leaves(item, f"{prefix}.{key}" if prefix else key).items()
        }
    if isinstance(value, (list, tuple)) and value:
        return {
            path: leaf
            for index, item in enumerate(value)
            for path, leaf in config_leaves(item, f"{prefix}[{index}]").items()
        }
    return {prefix: value}


def _display_text(value: str) -> str:
    escaped = {"\r": "\\r", "\n": "\\n", "\t": "\\t", "\b": "\\b", "\f": "\\f"}
    return "".join(
        escaped.get(char, f"\\u{ord(char):04x}")
        if ord(char) < 32 or 127 <= ord(char) <= 159 or char in "\u2028\u2029"
        else char
        for char in value
    )


def config_lines(value: Any) -> list[str]:
    """Render one key=value per line, with literal Windows path separators."""
    return [
        f"{_display_text(key)}="
        + (_display_text(item) if isinstance(item, str) else json.dumps(item, ensure_ascii=False))
        for key, item in config_leaves(value).items()
    ]

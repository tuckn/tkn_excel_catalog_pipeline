"""Arrange managed context without rewriting text outside its markers."""

from __future__ import annotations

from .context_profiles import ContextProfile
from .context_source import ContextError
from .sheet_layout import arrange_sheets


def sheet_heading(name: str) -> str:
    return name.replace("\r", " ").replace("\n", " ")


def arrange_context(text: str, sheets: list[dict[str, str]], profile: ContextProfile) -> str:
    """Group each sheet's context and extraction without changing its context hash."""
    try:
        return arrange_sheets(text, sheets, heading=profile.labels["sheet_heading"])
    except ValueError as exc:
        raise ContextError(str(exc)) from exc

"""Source-relative include and ignore pattern matching."""

from __future__ import annotations

from fnmatch import fnmatchcase
from pathlib import PurePosixPath

from .models import SourceConfig


def normalize_relative_path(value: str) -> str:
    """Normalize a configured or recorded relative path to portable POSIX form."""
    return value.replace("\\", "/").removeprefix("./").strip("/")


def matches_pattern(relative_path: str, pattern: str) -> bool:
    """Match a source-relative path using portable glob-like rules.

    Patterns without a slash match any single path component. Patterns containing a
    slash match the whole relative path. A leading ``**/`` also matches at the source
    root, and a trailing ``/**`` includes the named directory itself.
    """
    value = normalize_relative_path(relative_path)
    normalized = normalize_relative_path(pattern)
    if not value or not normalized:
        return False

    value_folded = value.casefold()
    pattern_folded = normalized.casefold()
    if "/" not in pattern_folded:
        return any(
            fnmatchcase(part.casefold(), pattern_folded) for part in PurePosixPath(value).parts
        )

    candidates = [pattern_folded]
    if pattern_folded.startswith("**/"):
        candidates.append(pattern_folded[3:])
    if any(fnmatchcase(value_folded, candidate) for candidate in candidates):
        return True

    for candidate in candidates:
        if candidate.endswith("/**"):
            prefix = candidate[:-3].rstrip("/")
            if value_folded == prefix or value_folded.startswith(prefix + "/"):
                return True
    return False


def matches_any(relative_path: str, patterns: tuple[str, ...]) -> bool:
    return any(matches_pattern(relative_path, pattern) for pattern in patterns)


def matches_include(relative_path: str, patterns: tuple[str, ...]) -> bool:
    """Match include patterns, treating slashless patterns as filename globs."""
    value = normalize_relative_path(relative_path)
    filename = PurePosixPath(value).name.casefold()
    for pattern in patterns:
        normalized = normalize_relative_path(pattern)
        if "/" not in normalized:
            if fnmatchcase(filename, normalized.casefold()):
                return True
        elif matches_pattern(value, normalized):
            return True
    return False


def is_source_path_in_scope(relative_path: str, source: SourceConfig) -> bool:
    """Return whether a recorded source-relative path belongs to this source scan."""
    value = normalize_relative_path(relative_path)
    if not value or (not source.recursive and "/" in value):
        return False
    return matches_include(value, source.include) and not matches_any(value, source.ignore)

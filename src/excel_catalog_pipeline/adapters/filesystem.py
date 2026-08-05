"""Filesystem safety checks and reversible workbook rename helpers."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from ..discovery import normalize_relative_path

INVALID_WINDOWS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *{f"COM{index}" for index in range(1, 10)},
    *{f"LPT{index}" for index in range(1, 10)},
}


class RenameError(ValueError):
    """Requested filename or rename target is unsafe."""


def validate_workbook_filename(current: Path, requested_name: str) -> Path:
    if not requested_name or requested_name != requested_name.strip():
        raise RenameError("Requested filename must be non-empty without surrounding whitespace")
    if requested_name.endswith((".", " ")) or INVALID_WINDOWS_CHARS.search(requested_name):
        raise RenameError("Requested filename contains Windows-invalid characters or suffix")
    requested = Path(requested_name)
    if requested.name != requested_name:
        raise RenameError("Requested filename must not contain a directory")
    if requested.stem.upper() in RESERVED_NAMES:
        raise RenameError("Requested filename uses a reserved Windows device name")
    if requested.suffix.casefold() != current.suffix.casefold():
        raise RenameError("Workbook extension changes are format conversions and are not allowed")
    if len(requested_name.encode("utf-8")) > 240:
        raise RenameError("Requested filename is too long")
    target = current.with_name(requested_name)
    if target.exists() and target.resolve() != current.resolve():
        raise RenameError(f"Rename target already exists: {target}")
    return target


def validate_workbook_relative_path(source_root: Path, current: Path, requested_path: str) -> Path:
    """Validate a source-relative workbook rename or move and return its target."""
    if not requested_path or requested_path != requested_path.strip():
        raise RenameError(
            "Requested source-relative path must be non-empty without surrounding whitespace"
        )
    normalized = normalize_relative_path(requested_path)
    requested = Path(normalized)
    if Path(requested_path).is_absolute() or not normalized:
        raise RenameError("Requested workbook path must be relative to the source root")
    if any(part in {"", ".", ".."} for part in requested.parts):
        raise RenameError("Requested workbook path must not contain empty, dot, or parent parts")
    for part in requested.parts:
        if part.endswith((".", " ")) or INVALID_WINDOWS_CHARS.search(part):
            raise RenameError(
                "Requested workbook path contains Windows-invalid characters or suffixes"
            )
        if Path(part).stem.upper() in RESERVED_NAMES:
            raise RenameError("Requested workbook path uses a reserved Windows device name")
        if len(part.encode("utf-8")) > 240:
            raise RenameError("Requested workbook path contains a component that is too long")
    if requested.suffix.casefold() != current.suffix.casefold():
        raise RenameError("Workbook extension changes are format conversions and are not allowed")

    root = source_root.resolve()
    target = (root / requested).resolve()
    try:
        target.relative_to(root)
    except ValueError as exc:
        raise RenameError("Requested workbook path escapes the source root") from exc
    if target.exists() and target != current.resolve():
        raise RenameError(f"Rename target already exists: {target}")
    return target


def rename_workbook(current: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    current.replace(target)


def backup_workbook(current: Path, backup_dir: Path) -> Path:
    backup_dir.mkdir(parents=True, exist_ok=True)
    candidate = backup_dir / current.name
    counter = 1
    while candidate.exists():
        candidate = backup_dir / f"{current.stem}-{counter}{current.suffix}"
        counter += 1
    shutil.copy2(current, candidate)
    return candidate

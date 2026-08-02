"""Filesystem safety checks and reversible workbook rename helpers."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

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


def rename_workbook(current: Path, target: Path) -> None:
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

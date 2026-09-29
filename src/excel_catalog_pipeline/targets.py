"""Resolve a single workbook/note onto the same source model used by batch sync."""

from __future__ import annotations

import hashlib
from dataclasses import replace
from pathlib import Path

from .adapters.markdown import read_note, workbook_path
from .config import ConfigError
from .discovery import is_source_path_in_scope
from .models import AppConfig, SourceConfig
from .paths import state_path
from .state import load_state


def single_source(config: AppConfig, workbook: Path, *, output: Path | None = None) -> SourceConfig:
    workbook = workbook.expanduser().resolve()
    if not workbook.is_file() or workbook.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ConfigError("Select an existing .xlsx or .xlsm workbook")
    if workbook.name.startswith("~$"):
        raise ConfigError("Excel lock files cannot be selected")
    matches = []
    for source in config.sources:
        if workbook.is_relative_to(source.path.resolve()):
            relative = workbook.relative_to(source.path.resolve()).as_posix()
            if is_source_path_in_scope(relative, source):
                matches.append(source)
    if len(matches) > 1:
        raise ConfigError("Workbook belongs to multiple sources; make the configured scopes unique")
    entries = load_state(state_path())["entries"]
    linked = []
    for entry in entries.values():
        if not isinstance(entry, dict):
            continue
        recorded = entry.get("workbookPath")
        if not recorded and entry.get("notePath") and Path(entry["notePath"]).is_file():
            recorded = workbook_path(read_note(Path(entry["notePath"])))
        if recorded and Path(recorded).resolve() == workbook:
            linked.append(entry)
    if len(linked) > 1:
        raise ConfigError("Multiple synchronization records match this workbook")
    recorded_note = Path(linked[0]["notePath"]).resolve() if linked else None
    if output is not None:
        raw_output = output.expanduser().absolute()
        if raw_output.is_symlink():
            raise ConfigError("Output must not be a symbolic link")
        target = raw_output.resolve()
        if recorded_note is not None and target != recorded_note:
            raise ConfigError("Workbook already has a linked note at a different path")
    elif recorded_note is not None:
        target = recorded_note
    elif matches:
        source = matches[0]
        target = source.note_root / workbook.relative_to(source.path)
        target = target.with_name(target.name + ".md")
    else:
        target = workbook.with_name(workbook.name + ".md")
    if target.suffix.lower() != ".md" or (target.exists() and not target.is_file()):
        raise ConfigError("Output must be a .md file")
    existing = read_note(target) if target.exists() else None
    if existing is not None:
        original = workbook_path(existing)
        if (
            str(existing.frontmatter.get("type", "")).lower() != "excel"
            or original is None
            or original.resolve() != workbook
        ):
            raise ConfigError(
                "Output is not the proxy note for this workbook; it was not overwritten"
            )
    if matches:
        source = matches[0]
    else:
        folder = workbook.parent
        source_id = "file-" + hashlib.sha256(str(folder).casefold().encode()).hexdigest()[:16]
        if existing is not None:
            source_id = str(existing.frontmatter.get("sourceRoot") or source_id)
            recorded_relative = Path(str(existing.frontmatter.get("sourceFileName", workbook.name)))
            if (
                recorded_relative.is_absolute()
                or ".." in recorded_relative.parts
                or not recorded_relative.parts
                or len(recorded_relative.parts) > len(workbook.parents)
            ):
                raise ConfigError("Invalid sourceFileName in the selected note")
            candidate = workbook.parents[len(recorded_relative.parts) - 1]
            if (candidate / recorded_relative).resolve() == workbook:
                folder = candidate
        source = SourceConfig(
            source_id,
            folder,
            ("*.xlsx", "*.xlsm"),
            target.parent,
            recursive=True,
            frontmatter_term_format="plain",
        )
    # An explicit file command touches one pair only, even when it belongs to a source.
    return replace(source, note_root=target.parent, single_workbook=workbook, single_note=target)


def source_for_note(config: AppConfig, note: Path) -> SourceConfig:
    note = note.expanduser().resolve()
    if not note.is_file():
        raise ConfigError("Select an existing proxy Markdown note")
    parsed = read_note(note)
    workbook = workbook_path(parsed)
    if workbook is None:
        raise ConfigError("The note does not identify its source workbook")
    return single_source(config, workbook, output=note)

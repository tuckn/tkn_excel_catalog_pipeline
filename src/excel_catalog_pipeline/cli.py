"""Command-line interface for the Excel catalog pipeline."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from . import __version__
from .config import ConfigError, config_as_dict, init_user_config, load_config, select_sources
from .models import Action, SourceConfig
from .pipeline import run_adopt, run_pull, run_push, run_status
from .reports import summarize_actions, write_report
from .state import StateError

SUCCESS = 25
logging.addLevelName(SUCCESS, "SUCCESS")
RESET = "\x1b[0m"
COLORS = {SUCCESS: "\x1b[32m", logging.ERROR: "\x1b[31m", logging.CRITICAL: "\x1b[31m"}
STATUS_ITEM_LIMIT = 20
STATUS_LABELS = {
    "tracked": "tracked workbooks",
    "untracked-source": "untracked workbooks",
    "untracked-note": "untracked proxy notes",
    "duplicate-id": "workbooks with duplicate IDs",
    "unsupported": "unsupported workbooks",
    "read-error": "read errors",
}


class ColorFormatter(logging.Formatter):
    def __init__(self, *, use_color: bool) -> None:
        super().__init__("[%(levelname)s] %(message)s")
        self.use_color = use_color

    def format(self, record: logging.LogRecord) -> str:
        rendered = super().format(record)
        color = COLORS.get(record.levelno) if self.use_color else None
        return f"{color}{rendered}{RESET}" if color else rendered


def supports_color(stream: Any) -> bool:
    try:
        if os.environ.get("NO_COLOR") is not None or os.environ.get("TERM") == "dumb":
            return False
        return bool(stream.isatty())
    except (AttributeError, OSError):
        return False


def configure_logging(*, quiet: bool, verbose: bool, no_color: bool) -> logging.Logger:
    level = logging.DEBUG if verbose else logging.ERROR if quiet else logging.INFO
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(ColorFormatter(use_color=not no_color and supports_color(sys.stderr)))
    logger = logging.getLogger("excel_catalog_pipeline")
    logger.handlers.clear()
    logger.addHandler(handler)
    logger.setLevel(level)
    logger.propagate = False
    return logger


def _preference(args: argparse.Namespace) -> str | None:
    if getattr(args, "prefer_source", False):
        return "source"
    if getattr(args, "prefer_note", False):
        return "note"
    return None


def _add_common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--source", help="Limit the command to one configured source id.")


def _add_preference(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--prefer-source",
        action="store_true",
        help=(
            "Use Excel for conflicts and, during pull, overwrite differing "
            "Markdown metadata with Excel values."
        ),
    )
    group.add_argument(
        "--prefer-note", action="store_true", help="Resolve conflicts from Markdown."
    )


def _add_execution_mode(parser: argparse.ArgumentParser, *, legacy_write_option: str) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Preview and validate planned changes without writing workbooks, notes, "
            "state, cache, reports, or external systems. This CLI uses no network, "
            "authentication, downloads, external services, or generative AI in either mode."
        ),
    )
    group.add_argument(
        legacy_write_option,
        dest="legacy_write",
        action="store_true",
        help=(
            "Deprecated compatibility option. Normal execution already writes; "
            "use --dry-run to preview."
        ),
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tkn-excel-catalog",
        description="Synchronize Excel workbook metadata with Markdown proxy notes.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--config", type=Path, help="Explicit YAML config file.")
    parser.add_argument(
        "--report-dir",
        type=Path,
        help="Override the run-report root for status and non-dry-run commands.",
    )
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument("-q", "--quiet", action="store_true")
    verbosity.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--no-color", action="store_true")

    commands = parser.add_subparsers(dest="command", required=True)
    config_parser = commands.add_parser("config", help="Create or inspect configuration.")
    config_commands = config_parser.add_subparsers(dest="config_command", required=True)
    config_commands.add_parser("show", help="Print the resolved non-secret configuration as JSON.")
    config_init = config_commands.add_parser(
        "init",
        help="Create the user-global config from the packaged template.",
    )
    config_init.add_argument(
        "--force",
        action="store_true",
        help="Replace a different existing user config. Reviewed edits will be lost.",
    )

    status = commands.add_parser(
        "status", help="Inventory sources and notes without changing them."
    )
    _add_common(status)

    pull = commands.add_parser(
        "pull",
        help="Apply Excel-to-Markdown changes; use --dry-run to preview.",
        description=(
            "Apply Excel-to-Markdown changes. Normal execution writes proxy notes and "
            "synchronization state; use --dry-run for a read-only preview."
        ),
    )
    _add_common(pull)
    _add_execution_mode(pull, legacy_write_option="--write-notes")
    _add_preference(pull)

    push = commands.add_parser(
        "push",
        help="Apply Markdown-to-Excel changes; use --dry-run to preview.",
        description=(
            "Apply Markdown-to-Excel changes. Normal execution backs up and updates "
            "workbooks and synchronization state; use --dry-run for a read-only preview."
        ),
    )
    _add_common(push)
    _add_execution_mode(push, legacy_write_option="--write-excel")
    push.add_argument(
        "--allow-rename",
        action="store_true",
        help="Allow validated source-relative path changes requested through sourceFileName.",
    )
    push.add_argument(
        "--note",
        action="append",
        default=[],
        help="Limit push to a note path/name, sourceFileName, or sourceId. Repeatable.",
    )
    _add_preference(push)

    adopt = commands.add_parser(
        "adopt",
        help="Assign stable workbook custom IDs; use --dry-run to preview.",
        description=(
            "Assign missing stable workbook custom IDs. Normal execution backs up and "
            "updates workbooks; use --dry-run for a read-only preview."
        ),
    )
    _add_common(adopt)
    _add_execution_mode(adopt, legacy_write_option="--write-excel")
    return parser


def _emit(payload: dict[str, Any]) -> None:
    print(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))


def _exit_code(summary: dict[str, Any]) -> int:
    if summary.get("errors", 0):
        return 1
    if summary.get("conflicts", 0):
        return 2
    return 0


def _one_line(value: str) -> str:
    return " ".join(value.split()) or "-"


def _yaml_single_quoted(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _log_source_configuration(logger: logging.Logger, sources: tuple[SourceConfig, ...]) -> None:
    lines: list[str] = []
    for source in sources:
        lines.extend(
            [
                f"  - id: {source.id}",
                f"    path: {_yaml_single_quoted(str(source.path))}",
                f"    recursive: {str(source.recursive).lower()}",
                "    include:",
                *(f"      - {json.dumps(pattern)}" for pattern in source.include),
                *(
                    [
                        "    ignore:",
                        *(f"      - {json.dumps(pattern)}" for pattern in source.ignore),
                    ]
                    if source.ignore
                    else ["    ignore: []"]
                ),
                "    notes:",
                f"      root: {_yaml_single_quoted(str(source.note_root))}",
                f"      profile: {source.profile}",
                f"      frontmatter_term_format: {source.frontmatter_term_format}",
                f"      rename_adapter: {source.rename_adapter}",
            ]
        )
    logger.info("Selected source configuration:\n%s", "\n".join(lines))


def _log_push_action(logger: logging.Logger, action: Action) -> None:
    if action.status == "unchanged":
        return
    if action.status == "written":
        level = SUCCESS
    elif action.status.endswith("error"):
        level = logging.ERROR
    elif action.conflict_fields or action.status in {
        "conflict",
        "duplicate-id",
        "missing-source",
        "pull-required",
        "rename-required",
    }:
        level = logging.WARNING
    else:
        level = logging.INFO
    file_name = Path(action.note_path).name if action.note_path else Path(action.source_path).name
    if action.status == "missing-source":
        fields = [_one_line(file_name)]
    else:
        fields = [
            f"fileName={_one_line(file_name)}",
            f"sourcePath={_one_line(action.source_path)}",
        ]
    fields.append(f"message={_one_line(action.message)}")
    logger.log(level, "[%s] %s", action.status, " | ".join(fields))


def _log_pull_action(logger: logging.Logger, action: Action) -> None:
    if action.status == "unchanged":
        return
    if action.status in {"created", "updated"}:
        level = SUCCESS
    elif action.status.endswith("error"):
        level = logging.ERROR
    elif action.conflict_fields or action.status in {
        "conflict",
        "duplicate-id",
        "missing-source",
        "push-required",
        "rename-required",
    }:
        level = logging.WARNING
    else:
        level = logging.INFO
    fields: list[str] = []
    if action.note_path:
        fields.append(f"notePath={_one_line(action.note_path)}")
    desired_note_path = action.details.get("desiredNotePath")
    if desired_note_path:
        fields.append(f"desiredNotePath={_one_line(str(desired_note_path))}")
    if action.source_path:
        fields.append(f"sourcePath={_one_line(action.source_path)}")
    if action.message and not action.message.startswith("stateMatch="):
        fields.append(f"message={_one_line(action.message)}")
    logger.log(level, "[%s] %s", action.status, " | ".join(fields) or "-")


def _log_readable_summary(logger: logging.Logger, summary: dict[str, Any]) -> None:
    mode = "write" if summary.get("writeEnabled") else "dry-run"
    report_path = summary.get("reportPath")
    summary_path = str(Path(str(report_path)) / "summary.json") if report_path else "-"
    lines = [
        f"  command: {summary.get('command', '-')}",
        f"  mode: {mode}",
        f"  result: {summary.get('status', '-')}",
        f"  changed files: {summary.get('changed', 0)}",
        f"  conflicts: {summary.get('conflicts', 0)}",
        f"  errors: {summary.get('errors', 0)}",
        "  status counts:",
    ]
    counts = summary.get("statusCounts", {})
    if isinstance(counts, dict) and counts:
        lines.extend(f"    {status}: {count}" for status, count in counts.items())
    else:
        lines.append("    none")
    lines.extend(
        [
            f"  report: {report_path or '-'}",
            f"  summary JSON: {summary_path}",
            f"  differences CSV: {summary.get('differencesPath') or '-'}",
        ]
    )
    logger.info("Summary:\n%s", "\n".join(lines))


def _log_adopt_action(logger: logging.Logger, action: Action) -> None:
    if action.status == "unchanged":
        return
    if action.status == "adopted":
        level = SUCCESS
    elif action.status.endswith("error"):
        level = logging.ERROR
    elif action.conflict_fields or action.status == "duplicate-id":
        level = logging.WARNING
    else:
        level = logging.INFO
    fields = [f"sourcePath={_one_line(action.source_path)}"]
    if action.workbook_id:
        fields.append(f"workbookId={_one_line(action.workbook_id)}")
    if action.message:
        fields.append(f"message={_one_line(action.message)}")
    logger.log(level, "[%s] %s", action.status, " | ".join(fields))


def _log_legacy_write_warning(logger: logging.Logger, args: argparse.Namespace) -> None:
    if getattr(args, "legacy_write", False):
        option = "--write-notes" if args.command == "pull" else "--write-excel"
        logger.warning(
            "%s is deprecated and no longer required; normal %s execution writes. "
            "Use --dry-run for a read-only preview.",
            option,
            args.command,
        )


def _display_difference_value(value: str) -> str:
    if not value:
        return "<empty>"
    rendered = " ".join(value.split())
    return rendered if len(rendered) <= 120 else rendered[:117] + "..."


def _log_action_differences(logger: logging.Logger, actions: list[Action]) -> None:
    lines: list[str] = []
    direction_labels = {
        "source-to-note": "Excel -> Markdown",
        "note-to-source": "Markdown -> Excel",
        "": "review only",
    }
    for action in actions:
        if not action.field_differences:
            continue
        target = action.source_path or Path(action.note_path).name
        lines.append(f"  [{action.status}] {target}")
        for difference in action.field_differences:
            planned = direction_labels.get(
                difference.get("plannedDirection", ""),
                difference.get("plannedDirection", "review only"),
            )
            lines.append(
                "    {field} | Excel: {excel} | Markdown: {note} | "
                "base: {base} | apply: {planned}".format(
                    field=difference.get("field", "-"),
                    excel=_display_difference_value(difference.get("excelValue", "")),
                    note=_display_difference_value(difference.get("noteValue", "")),
                    base=_display_difference_value(difference.get("baseValue", "")),
                    planned=planned,
                )
            )
    if lines:
        logger.debug("Metadata differences:\n%s", "\n".join(lines))


def _status_action_target(action: Action, source: SourceConfig) -> str:
    if action.source_path:
        return action.source_path
    if action.note_path:
        note_path = Path(action.note_path)
        try:
            return str(note_path.resolve().relative_to(source.note_root.resolve()))
        except (OSError, ValueError):
            return note_path.name
    return "source scan"


def _status_action_details(action: Action) -> str:
    details: list[str] = []
    if action.workbook_id:
        details.append(f"workbookId={_one_line(action.workbook_id)}")
    warnings = action.details.get("warnings", [])
    if isinstance(warnings, list):
        details.extend(_one_line(str(warning)) for warning in warnings if warning)
    if action.message and not action.message.startswith("stateMatch="):
        details.append(_one_line(action.message))
    return " | ".join(details)


def _ordered_statuses(counts: Counter[str]) -> list[str]:
    known = [status for status in STATUS_LABELS if counts[status]]
    return [*known, *sorted(status for status in counts if status not in STATUS_LABELS)]


def _log_status_results(
    logger: logging.Logger,
    actions: list[Action],
    sources: tuple[SourceConfig, ...],
) -> None:
    lines: list[str] = []
    for source in sources:
        source_actions = [action for action in actions if action.source_root_id == source.id]
        counts = Counter(action.status for action in source_actions)
        statuses = _ordered_statuses(counts)
        lines.append(f"  {source.id}:")
        if not statuses:
            lines.append("    no workbooks or proxy notes found")
            continue
        for status in statuses:
            lines.append(f"    {STATUS_LABELS.get(status, status)}: {counts[status]}")
        attention = [status for status in statuses if status != "tracked"]
        if not attention:
            lines.append("    items requiring attention: none")
            continue
        lines.append("    items requiring attention:")
        for status in attention:
            matching = [action for action in source_actions if action.status == status]
            lines.append(f"      {STATUS_LABELS.get(status, status)}:")
            for action in matching[:STATUS_ITEM_LIMIT]:
                item = f"        - {_status_action_target(action, source)}"
                details = _status_action_details(action)
                lines.append(f"{item} | {details}" if details else item)
            remaining = len(matching) - STATUS_ITEM_LIMIT
            if remaining > 0:
                lines.append(f"        ... {remaining} more; see the report below")
    logger.info("Status results:\n%s", "\n".join(lines))


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logger = configure_logging(
        quiet=args.quiet,
        verbose=args.verbose,
        no_color=args.no_color,
    )
    try:
        if args.command == "config" and args.config_command == "init":
            if args.config is not None:
                raise ConfigError("--config cannot be combined with config init")
            status, path = init_user_config(force=args.force)
            if status == "unchanged":
                logger.info("Config already matches the packaged template: %s", path)
            else:
                logger.log(SUCCESS, "Config %s: %s", status, path)
            _emit({"status": status, "command": "config init", "configPath": str(path)})
            return 0
        config = load_config(explicit=args.config)
        if args.command == "config":
            _emit({"status": "success", "command": "config show", "config": config_as_dict(config)})
            return 0
        sources = select_sources(config, args.source)
        if not sources:
            raise ConfigError("No sources are configured")
        report_root = args.report_dir.expanduser().resolve() if args.report_dir else None
        logger.info(
            "Running %s for %d configured source(s): %s.",
            args.command,
            len(sources),
            ", ".join(source.id for source in sources),
        )
        if args.command == "push":
            _log_source_configuration(logger, sources)
        if args.command in {"pull", "push", "adopt"}:
            _log_legacy_write_warning(logger, args)
        backup_dir: Path | None = None
        if args.command == "status":
            actions = run_status(config, sources)
            write_enabled = False
        elif args.command == "pull":
            write_enabled = not args.dry_run
            actions = run_pull(
                config,
                sources,
                write_notes=write_enabled,
                preference=_preference(args),
            )
            for action in actions:
                _log_pull_action(logger, action)
        elif args.command == "push":
            write_enabled = not args.dry_run
            actions, backup_dir = run_push(
                config,
                sources,
                write_excel=write_enabled,
                allow_rename=args.allow_rename,
                preference=_preference(args),
                note_filters=tuple(args.note),
                on_action=lambda action: _log_push_action(logger, action),
            )
        elif args.command == "adopt":
            write_enabled = not args.dry_run
            actions, backup_dir = run_adopt(
                config,
                sources,
                write_excel=write_enabled,
            )
            for action in actions:
                _log_adopt_action(logger, action)
        else:
            raise AssertionError(args.command)
        extra = {"backupPath": str(backup_dir) if backup_dir else ""}
        if args.command in {"pull", "push", "adopt"} and args.dry_run:
            summary = summarize_actions(
                args.command,
                actions,
                write_enabled=False,
                extra=extra,
            )
            if report_root is not None:
                logger.warning("--report-dir has no effect in dry-run mode; no report was written.")
        else:
            summary, _ = write_report(
                args.command,
                actions,
                write_enabled=write_enabled,
                report_root=report_root,
                extra=extra,
            )
        if args.verbose:
            _log_action_differences(logger, actions)
        if args.command == "status":
            _log_status_results(logger, actions, sources)
        result_suffix = (
            "no persistent report was written"
            if args.command in {"pull", "push", "adopt"} and args.dry_run
            else f"report: {summary['reportPath']}"
        )
        if summary["status"] == "success":
            logger.log(SUCCESS, "%s completed; %s", args.command, result_suffix)
        elif summary["status"] == "conflict":
            logger.warning("%s found conflicts; %s", args.command, result_suffix)
        else:
            logger.error(
                "%s failed for one or more targets; %s", args.command, result_suffix
            )
        if args.command in {"pull", "push", "adopt"}:
            _log_readable_summary(logger, summary)
        return _exit_code(summary)
    except (ConfigError, StateError) as exc:
        logger.error("%s", exc)
        if args.command == "config":
            _emit({"status": "config-error", "command": args.command, "message": str(exc)})
        return 3
    except Exception as exc:  # pragma: no cover - last-resort CLI boundary
        if args.verbose:
            logger.exception("Unexpected failure")
        else:
            logger.error("Unexpected failure: %s", exc)
        if args.command == "config":
            _emit({"status": "error", "command": args.command, "message": str(exc)})
        return 1

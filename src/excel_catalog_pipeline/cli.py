"""Command-line interface for the Excel catalog pipeline."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .config import ConfigError, config_as_dict, init_user_config, load_config, select_sources
from .models import Action, SourceConfig
from .pipeline import run_adopt, run_pull, run_push, run_status
from .reports import write_report
from .state import StateError

SUCCESS = 25
logging.addLevelName(SUCCESS, "SUCCESS")
RESET = "\x1b[0m"
COLORS = {SUCCESS: "\x1b[32m", logging.ERROR: "\x1b[31m", logging.CRITICAL: "\x1b[31m"}


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
    group.add_argument("--prefer-source", action="store_true", help="Resolve conflicts from Excel.")
    group.add_argument(
        "--prefer-note", action="store_true", help="Resolve conflicts from Markdown."
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-catalog",
        description="Synchronize Excel workbook metadata with Markdown proxy notes.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--config", type=Path, help="Explicit YAML config file.")
    parser.add_argument("--report-dir", type=Path, help="Override the application run-report root.")
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

    pull = commands.add_parser("pull", help="Plan or apply Excel-to-Markdown changes.")
    _add_common(pull)
    pull.add_argument(
        "--write-notes",
        action="store_true",
        help="Apply planned proxy-note writes. The default only reports.",
    )
    _add_preference(pull)

    push = commands.add_parser("push", help="Plan or apply Markdown-to-Excel changes.")
    _add_common(push)
    push.add_argument(
        "--write-excel",
        action="store_true",
        help="Back up and update workbooks. The default only reports.",
    )
    push.add_argument(
        "--allow-rename",
        action="store_true",
        help="Allow validated filename changes requested through sourceFileName.",
    )
    push.add_argument(
        "--note",
        action="append",
        default=[],
        help="Limit push to a note path/name, sourceFileName, or sourceId. Repeatable.",
    )
    _add_preference(push)

    adopt = commands.add_parser("adopt", help="Plan or assign stable workbook custom IDs.")
    _add_common(adopt)
    adopt.add_argument(
        "--write-excel",
        action="store_true",
        help="Back up workbooks and add missing TknExcelCatalogId values.",
    )
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


def _log_source_configuration(
    logger: logging.Logger, sources: tuple[SourceConfig, ...]
) -> None:
    lines: list[str] = []
    for source in sources:
        lines.extend(
            [
                f"  - id: {source.id}",
                f"    path: {_yaml_single_quoted(str(source.path))}",
                "    include:",
                *(f"      - {json.dumps(pattern)}" for pattern in source.include),
                "    notes:",
                f"      root: {_yaml_single_quoted(str(source.note_root))}",
                f"      profile: {source.profile}",
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


def _log_readable_summary(logger: logging.Logger, summary: dict[str, Any]) -> None:
    logger.info("Summary:\n%s", json.dumps(summary, ensure_ascii=False, indent=2))


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
        backup_dir: Path | None = None
        if args.command == "status":
            actions = run_status(config, sources)
            write_enabled = False
        elif args.command == "pull":
            actions = run_pull(
                config,
                sources,
                write_notes=args.write_notes,
                preference=_preference(args),
            )
            write_enabled = bool(args.write_notes)
        elif args.command == "push":
            if args.allow_rename and not args.write_excel:
                raise ConfigError("--allow-rename requires --write-excel")
            actions, backup_dir = run_push(
                config,
                sources,
                write_excel=args.write_excel,
                allow_rename=args.allow_rename,
                preference=_preference(args),
                note_filters=tuple(args.note),
                on_action=lambda action: _log_push_action(logger, action),
            )
            write_enabled = bool(args.write_excel)
        elif args.command == "adopt":
            actions, backup_dir = run_adopt(
                config,
                sources,
                write_excel=args.write_excel,
            )
            write_enabled = bool(args.write_excel)
        else:
            raise AssertionError(args.command)
        extra = {"backupPath": str(backup_dir) if backup_dir else ""}
        summary, _ = write_report(
            args.command,
            actions,
            write_enabled=write_enabled,
            report_root=report_root,
            extra=extra,
        )
        if summary["status"] == "success":
            logger.log(SUCCESS, "%s completed; report: %s", args.command, summary["reportPath"])
        elif summary["status"] == "conflict":
            logger.warning("%s found conflicts; report: %s", args.command, summary["reportPath"])
        else:
            logger.error(
                "%s failed for one or more targets; report: %s", args.command, summary["reportPath"]
            )
        if args.command == "push":
            _log_readable_summary(logger, summary)
        _emit(summary)
        return _exit_code(summary)
    except (ConfigError, StateError) as exc:
        logger.error("%s", exc)
        _emit({"status": "config-error", "command": args.command, "message": str(exc)})
        return 3
    except Exception as exc:  # pragma: no cover - last-resort CLI boundary
        if args.verbose:
            logger.exception("Unexpected failure")
        else:
            logger.error("Unexpected failure: %s", exc)
        _emit({"status": "error", "command": args.command, "message": str(exc)})
        return 1

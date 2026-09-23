"""YAML configuration loading with explicit precedence and strict validation."""

from __future__ import annotations

import uuid
from contextlib import suppress
from copy import deepcopy
from dataclasses import asdict
from importlib.resources import files
from pathlib import Path
from typing import Any, cast

import yaml

from .models import AppConfig, ContextConfig, FrontmatterTermFormat, SourceConfig, SyncConfig
from .note_resources import NoteResourceError, load_note_template
from .paths import global_config_path

SCHEMA_VERSION = 1
DEFAULT_CONFIG: dict[str, Any] = {
    "schema_version": SCHEMA_VERSION,
    "sources": {},
    "context": asdict(ContextConfig()),
    "sync": {
        "pull_preserves_user_metadata": True,
        "delete_missing_notes": False,
        "delete_missing_workbooks": False,
        "allow_source_rename": False,
        "max_extracted_text_chars": 12000,
    },
}
TOP_LEVEL_KEYS = {"schema_version", "sources", "sync", "context"}
SYNC_KEYS = set(DEFAULT_CONFIG["sync"])
SOURCE_KEYS = {"id", "path", "recursive", "include", "ignore", "notes"}
NOTES_KEYS = {"root", "profile", "frontmatter_term_format", "rename_adapter"}
FRONTMATTER_TERM_FORMATS = {"obsidian-link", "plain"}


class ConfigError(ValueError):
    """Configuration could not be loaded or validated."""


class ConfigLoader(yaml.SafeLoader):
    """Reject duplicate explicit YAML keys instead of silently losing settings."""

    def construct_mapping(self, node: yaml.MappingNode, deep: bool = False) -> dict[Any, Any]:
        seen: set[Any] = set()
        for key_node, _ in node.value:
            if key_node.tag == "tag:yaml.org,2002:merge":
                continue
            key = self.construct_object(key_node, deep=deep)
            try:
                if key in seen:
                    raise yaml.constructor.ConstructorError(
                        "while reading config",
                        node.start_mark,
                        f"Duplicate config key: {key!r}",
                        key_node.start_mark,
                    )
                seen.add(key)
            except TypeError as exc:
                raise yaml.constructor.ConstructorError(
                    "while reading config",
                    node.start_mark,
                    "Config keys must be hashable",
                    key_node.start_mark,
                ) from exc
        return super().construct_mapping(node, deep=deep)


def _source_items(raw: Any) -> dict[str, dict[str, Any]]:
    """Use mapping keys as IDs; accept legacy lists without rewriting user files."""
    if isinstance(raw, list):
        result: dict[str, dict[str, Any]] = {}
        for index, item in enumerate(raw):
            if not isinstance(item, dict):
                raise ConfigError(f"sources[{index}] must be a mapping")
            source_id = item.get("id")
            if not isinstance(source_id, str) or not source_id.strip():
                raise ConfigError(f"sources[{index}].id must be a non-empty string")
            if source_id in result:
                raise ConfigError(f"Duplicate source id: {source_id}")
            result[source_id] = item
        raw = result
    if not isinstance(raw, dict):
        raise ConfigError("sources must be a mapping keyed by source ID")
    for source_id, item in raw.items():
        if not isinstance(source_id, str) or not source_id.strip():
            raise ConfigError("sources keys must be non-empty strings")
        if not isinstance(item, dict):
            raise ConfigError(f"sources[{source_id!r}] must be a mapping")
        explicit_id = item.get("id")
        if explicit_id is not None and explicit_id != source_id:
            raise ConfigError(f"sources[{source_id!r}].id must match its source key or be null")
    return cast(dict[str, dict[str, Any]], raw)


def config_template_text() -> str:
    return (
        files("excel_catalog_pipeline")
        .joinpath("resources/config.example.yaml")
        .read_text(encoding="utf-8")
    )


def init_user_config(*, force: bool = False, target: Path | None = None) -> tuple[str, Path]:
    """Create the user-global config from the packaged template without implicit overwrite."""
    destination = (target or global_config_path()).expanduser().resolve()
    template = config_template_text()
    try:
        parsed = yaml.safe_load(template)
    except yaml.YAMLError as exc:  # pragma: no cover - packaged resource regression
        raise ConfigError(f"Packaged config template is invalid: {exc}") from exc
    if not isinstance(parsed, dict):  # pragma: no cover - packaged resource regression
        raise ConfigError("Packaged config template must contain a mapping")
    validate_config(parsed)

    if destination.exists():
        try:
            existing = destination.read_text(encoding="utf-8-sig")
        except OSError as exc:
            raise ConfigError(f"Failed to read existing config {destination}: {exc}") from exc
        if existing == template:
            return "unchanged", destination
        if not force:
            raise ConfigError(
                f"Config already exists and differs from the template: {destination}. "
                "Use config init --force only if replacing it is intentional."
            )
        status = "replaced"
    else:
        status = "created"

    temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary.write_text(template, encoding="utf-8", newline="\n")
        temporary.replace(destination)
    except OSError as exc:
        with suppress(OSError):
            temporary.unlink(missing_ok=True)
        raise ConfigError(f"Failed to write config {destination}: {exc}") from exc
    return status, destination


def _deep_merge(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in overlay.items():
        if key != "sources" and isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def _read_yaml(path: Path) -> dict[str, Any]:
    try:
        raw = yaml.load(path.read_text(encoding="utf-8-sig"), Loader=ConfigLoader)
    except (OSError, yaml.YAMLError) as exc:
        raise ConfigError(f"Failed to read config {path}: {exc}") from exc
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ConfigError(f"Config must contain a mapping: {path}")
    return raw


def _reject_unknown(data: dict[str, Any], allowed: set[str], where: str) -> None:
    unknown = sorted(set(data) - allowed)
    if unknown:
        raise ConfigError(f"Unknown {where} key(s): {', '.join(unknown)}")


def _path_value(value: Any, field: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{field} must be a non-empty path string")
    return Path(value).expanduser().resolve()


def load_config(*, explicit: Path | None = None, cwd: Path | None = None) -> AppConfig:
    cwd = (cwd or Path.cwd()).resolve()
    candidates = [global_config_path(), cwd / ".tkn" / "config.yaml"]
    if explicit is not None:
        candidates.append(explicit.expanduser().resolve())

    merged = deepcopy(DEFAULT_CONFIG)
    loaded: list[Path] = []
    for path in candidates:
        if path.exists():
            merged = _deep_merge(merged, _read_yaml(path))
            loaded.append(path)
        elif explicit is not None and path == explicit.expanduser().resolve():
            raise ConfigError(f"Explicit config not found: {path}")

    return validate_config(merged, loaded_files=tuple(loaded))


def validate_config(data: dict[str, Any], *, loaded_files: tuple[Path, ...] = ()) -> AppConfig:
    _reject_unknown(data, TOP_LEVEL_KEYS, "top-level")
    if data.get("schema_version") != SCHEMA_VERSION:
        raise ConfigError(
            f"Unsupported schema_version: {data.get('schema_version')!r}; expected {SCHEMA_VERSION}"
        )

    sync_raw = data.get("sync")
    if not isinstance(sync_raw, dict):
        raise ConfigError("sync must be a mapping")
    _reject_unknown(sync_raw, SYNC_KEYS, "sync")
    for name in (
        "pull_preserves_user_metadata",
        "delete_missing_notes",
        "delete_missing_workbooks",
        "allow_source_rename",
    ):
        if not isinstance(sync_raw.get(name), bool):
            raise ConfigError(f"sync.{name} must be boolean")
    max_chars = sync_raw.get("max_extracted_text_chars")
    if not isinstance(max_chars, int) or isinstance(max_chars, bool) or max_chars <= 0:
        raise ConfigError("sync.max_extracted_text_chars must be a positive integer")
    sync = SyncConfig(**sync_raw)

    sources_raw = _source_items(data.get("sources"))
    sources: list[SourceConfig] = []
    for source_id, item in sources_raw.items():
        where = f"sources[{source_id!r}]"
        _reject_unknown(item, SOURCE_KEYS, where)
        recursive = item.get("recursive", False)
        if not isinstance(recursive, bool):
            raise ConfigError(f"{where}.recursive must be boolean")
        include = item.get("include", ["*.xlsx", "*.xlsm"])
        if (
            not isinstance(include, list)
            or not include
            or not all(isinstance(value, str) and value for value in include)
        ):
            raise ConfigError(f"{where}.include must be a non-empty string list")
        ignore = item.get("ignore", [])
        if not isinstance(ignore, list) or not all(
            isinstance(value, str) and value for value in ignore
        ):
            raise ConfigError(f"{where}.ignore must be a string list")
        notes = item.get("notes")
        if not isinstance(notes, dict):
            raise ConfigError(f"{where}.notes must be a mapping")
        _reject_unknown(notes, NOTES_KEYS, f"{where}.notes")
        profile = notes.get("profile", "tkn-obsidian-v1")
        if not isinstance(profile, str) or not profile.strip():
            raise ConfigError(f"{where}.notes.profile must be a non-empty string")
        try:
            load_note_template(profile)
        except NoteResourceError as exc:
            raise ConfigError(f"{where}.notes.profile: {exc}") from exc
        frontmatter_term_format = notes.get("frontmatter_term_format", "obsidian-link")
        if (
            not isinstance(frontmatter_term_format, str)
            or frontmatter_term_format not in FRONTMATTER_TERM_FORMATS
        ):
            allowed = ", ".join(sorted(FRONTMATTER_TERM_FORMATS))
            raise ConfigError(f"{where}.notes.frontmatter_term_format must be one of: {allowed}")
        sources.append(
            SourceConfig(
                id=source_id,
                path=_path_value(item.get("path"), f"{where}.path"),
                include=tuple(include),
                note_root=_path_value(notes.get("root"), f"{where}.notes.root"),
                recursive=recursive,
                ignore=tuple(ignore),
                profile=profile,
                frontmatter_term_format=cast(FrontmatterTermFormat, frontmatter_term_format),
                rename_adapter=str(notes.get("rename_adapter", "report-only")),
            )
        )
    context_raw = data.get("context", {})
    if not isinstance(context_raw, dict):
        raise ConfigError("context must be a mapping")
    defaults = asdict(ContextConfig())
    _reject_unknown(context_raw, set(defaults), "context")
    context_values = {**defaults, **context_raw}
    for name, default in defaults.items():
        value = context_values[name]
        if isinstance(default, int):
            if type(value) is not int or value <= 0:
                raise ConfigError(f"context.{name} must be a positive integer")
        elif not isinstance(value, str) or not value.strip():
            raise ConfigError(f"context.{name} must be a non-empty string")
    if context_values["reasoning_effort"] not in {"low", "medium", "high", "xhigh"}:
        raise ConfigError("context.reasoning_effort must be low, medium, high or xhigh")
    if context_values["overlap_points"] * 2 >= min(
        context_values["tile_width_points"], context_values["tile_height_points"]
    ):
        raise ConfigError("context.overlap_points must be less than half the tile dimensions")
    if not 72 <= context_values["image_dpi"] <= 300:
        raise ConfigError("context.image_dpi must be between 72 and 300")
    if context_values["max_images"] > 100:
        raise ConfigError("context.max_images must be at most 100")
    return AppConfig(
        schema_version=SCHEMA_VERSION,
        sources=tuple(sources),
        sync=sync,
        loaded_files=loaded_files,
        context=ContextConfig(**context_values),
    )


def select_sources(config: AppConfig, selector: str | None) -> tuple[SourceConfig, ...]:
    if selector is None:
        return config.sources
    selected = tuple(source for source in config.sources if source.id == selector)
    if not selected:
        raise ConfigError(f"Unknown source id: {selector}")
    return selected


def config_as_dict(config: AppConfig) -> dict[str, Any]:
    return {
        "schema_version": config.schema_version,
        "sources": {
            source.id: {
                "path": str(source.path),
                "recursive": source.recursive,
                "include": list(source.include),
                "ignore": list(source.ignore),
                "notes": {
                    "root": str(source.note_root),
                    "profile": source.profile,
                    "frontmatter_term_format": source.frontmatter_term_format,
                    "rename_adapter": source.rename_adapter,
                },
            }
            for source in config.sources
        },
        "sync": {
            "pull_preserves_user_metadata": config.sync.pull_preserves_user_metadata,
            "delete_missing_notes": config.sync.delete_missing_notes,
            "delete_missing_workbooks": config.sync.delete_missing_workbooks,
            "allow_source_rename": config.sync.allow_source_rename,
            "max_extracted_text_chars": config.sync.max_extracted_text_chars,
        },
        "context": asdict(config.context),
        "loadedConfigFiles": [str(path) for path in config.loaded_files],
    }

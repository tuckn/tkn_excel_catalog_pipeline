"""Opt-in AI generation for unified proxy notes."""

from __future__ import annotations

import hashlib
import io
import json
import logging
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import asdict
from pathlib import Path
from typing import Any

from tkn_genai_bridge import Profile

from .adapters.markdown import discover_notes
from .adapters.ooxml import content_fingerprint, read_custom_properties
from .context_profiles import load_prompt, profile_name, prompt_digest
from .context_provider import (
    PROMPT_VERSION,
    atomic_json,
    generate_markdown,
    generation_plan,
    resolve_profile,
    utc_now,
)
from .context_render import render_sheet
from .context_source import ContextError, extract_sheet, rendering_snapshot, sheet_list
from .discovery import is_source_path_in_scope
from .models import ContextConfig, SourceConfig
from .note_layout import patch_frontmatter
from .paths import state_root
from .shared_read import read_shared


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def resolve_workbook(source: SourceConfig, selector: str, config: ContextConfig) -> Path:
    candidate = Path(selector).expanduser()
    path = (candidate if candidate.is_absolute() else source.path / candidate).resolve()
    try:
        relative = path.relative_to(source.path.resolve()).as_posix()
    except ValueError as exc:
        raise ContextError("Workbook must be inside the selected source root") from exc
    if not is_source_path_in_scope(relative, source) or path.name.startswith("~$"):
        raise ContextError("Workbook is excluded by the selected source configuration")
    if path.suffix.lower() not in {".xlsx", ".xlsm"} or not path.is_file():
        raise ContextError("Select an existing .xlsx or .xlsm workbook")
    if path.stat().st_size > config.max_workbook_mb * 1024 * 1024:
        raise ContextError("Workbook exceeds generation.max_workbook_mb")
    return path


def find_proxy(source: SourceConfig, workbook: Path, data: bytes) -> tuple[Path, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        workbook_id = read_custom_properties(archive).get("TknExcelCatalogId", "")
    relative = workbook.relative_to(source.path.resolve()).as_posix()
    source_id = f"{source.id}:{workbook_id or relative}"
    matches = [note for note in discover_notes(source) if note.source_id == source_id]
    if len(matches) != 1:
        raise ContextError(
            "Expected exactly one matching proxy note. Run pull first and resolve duplicate notes."
        )
    note = matches[0]
    if note.path.is_symlink() or not note.path.resolve().is_relative_to(source.note_root.resolve()):
        raise ContextError("Proxy note must remain inside the configured notes root")
    return note.path, digest(source_id.encode())[:20]


def block_pattern(sheet_id: str) -> re.Pattern[str]:
    return re.compile(
        rf"<!-- excel-catalog:begin context-{re.escape(sheet_id)} -->\r?\n.*?"
        rf"<!-- excel-catalog:end context-{re.escape(sheet_id)} -->",
        re.DOTALL,
    )


def existing_block(text: str, sheet_id: str) -> str | None:
    begin = f"<!-- excel-catalog:begin context-{sheet_id} -->"
    end = f"<!-- excel-catalog:end context-{sheet_id} -->"
    blocks = list(block_pattern(sheet_id).finditer(text))
    if text.count(begin) != len(blocks) or text.count(end) != len(blocks) or len(blocks) > 1:
        raise ContextError(
            "Malformed or duplicate context markers; repair the note before rebuilding"
        )
    return blocks[0].group() if blocks else None


def update_text(text: str, sheet_id: str, block: str) -> str:
    existing_block(text, sheet_id)
    if block_pattern(sheet_id).search(text):
        return block_pattern(sheet_id).sub(lambda _: block, text, count=1)
    newline = "\r\n" if "\r\n" in text else "\n"
    # Keep the workbook overview first, followed by sheet context and extracted text.
    offset = text.find("<!-- excel-catalog:begin extracted-text -->")
    if offset < 0:
        map_end = re.search(r"<!-- excel-catalog:end workbook-map -->(?:\r?\n){0,2}", text)
        offset = map_end.end() if map_end else len(text)
    prefix, suffix = text[:offset], text[offset:]
    separator = "" if prefix.endswith(newline * 2) else newline * 2
    return prefix + separator + block + newline * 2 + suffix


def _read_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ContextError(f"Cannot read context state: {path}") from exc
    if not isinstance(value, dict):
        raise ContextError("Invalid context state")
    return value


def _assets_intact(note: Path, state: dict[str, Any]) -> bool:
    assets = state.get("assets")
    if not isinstance(assets, dict):
        return False
    if not assets:
        return state.get("origin") == "import"
    for relative, checksum in assets.items():
        path = (note.parent / relative).resolve()
        if not path.is_relative_to((note.parent / "img").resolve()) or not path.is_file():
            return False
        if digest(path.read_bytes()) != checksum:
            return False
    return True


@contextmanager
def note_lock(note: Path) -> Iterator[None]:
    # A persistent lock left by an abruptly killed process is intentionally not stolen.
    location = state_root() / "context" / "locks" / (digest(str(note.resolve()).encode()) + ".lock")
    location.parent.mkdir(parents=True, exist_ok=True)
    try:
        stream = location.open("x", encoding="utf-8")
    except FileExistsError as exc:
        raise ContextError(
            f"Another build may be active. Inspect the process before removing stale lock: {location}"
        ) from exc
    try:
        with stream:
            stream.write(utc_now())
        yield
    finally:
        location.unlink(missing_ok=True)


def _validate_markdown(markdown: str, image_paths: set[str]) -> None:
    if "<!-- excel-catalog:" in markdown or markdown.lstrip().startswith("---"):
        raise ContextError("Generated Markdown contains reserved markers or unexpected Frontmatter")
    # All embedded images must be durable caller-provided assets, not temporary paths.
    for target in re.findall(r"!\[[^\]]*\]\(([^)]+)\)", markdown):
        if target.strip("<>") not in image_paths:
            raise ContextError("Generated Markdown contains an unrecognized image path")
    targets = re.findall(r"!?\[[^\]]*\]\(([^)]+)\)", markdown)
    if any(re.search(r"(?:[A-Za-z]:[\\/]|file://|/tmp/)", target) for target in targets):
        raise ContextError(
            "Generated Markdown contains an absolute local path; refusing temporary/nonportable links"
        )


def build_context(
    source: SourceConfig,
    workbook: Path,
    data: bytes,
    selected: list[dict[str, str]],
    config: ContextConfig,
    logger: logging.Logger,
    *,
    dry_run: bool,
    force: bool,
    note_path: Path | None = None,
    preview_text: str | None = None,
    include_overview: bool = False,
) -> dict[str, Any]:
    if note_path is None:
        note, book_key = find_proxy(source, workbook, data)
    else:
        note = note_path
        book_key = digest(str(workbook.resolve()).encode())[:20]
    results: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    state_dir = state_root() / "context" / digest(str(note.resolve()).encode())[:24]
    # Validate all requested sheets before any rendering, paid call, or persistent write.
    prepared = [
        extract_sheet(data, sheet, max_cells=config.max_cells, max_objects=config.max_objects)
        for sheet in selected
    ]
    for evidence in prepared:
        if len(json.dumps(evidence, ensure_ascii=False)) > config.max_input_chars:
            raise ContextError("Evidence exceeds generation.max_input_chars; no AI was called")
    logger.info(
        "Reading a saved-file snapshot; unsaved edits in Excel are not included. SHA256=%s",
        digest(data),
    )

    def run() -> None:
        connection: Profile | None = None
        plan: dict[str, Any] = {}
        for sheet, evidence in zip(selected, prepared, strict=True):
            result: dict[str, Any] = {"sheet": sheet["name"], "notePath": str(note)}
            results.append(result)
            original = note.read_bytes() if note.exists() else (preview_text or "").encode()
            text = original.decode("utf-8-sig")
            old_block = existing_block(text, sheet["id"])
            state_path = state_dir / f"sheet-{sheet['id']}.json"
            state = _read_state(state_path)
            if (
                old_block
                and digest(old_block.replace("\r\n", "\n").encode()) != state.get("blockSha256")
                and not force
            ):
                raise ContextError(
                    f"Context for {sheet['name']} was edited or has no matching state. Use --force only to replace it intentionally."
                )
            if old_block and state.get("origin") == "import" and not force:
                if not _assets_intact(note, state):
                    raise ContextError(
                        "Imported context assets are missing or modified; restore the saved assets or use --force to replace the context by generation."
                    )
                result["status"] = "retained"
                logger.info(
                    "Sheet %s: imported context retained; no AI call. Use --force to replace it by generation.",
                    sheet["name"],
                )
                continue
            if connection is None:
                connection = resolve_profile(config)
                plan = generation_plan(connection)
            result["bridgePlan"] = plan
            key = digest(
                json.dumps(
                    {
                        "sheet": evidence["fingerprint"],
                        "config": {
                            k: v
                            for k, v in asdict(config).items()
                            if k not in {"bridge_profile", "overrides"}
                        },
                        "bridge": {
                            k: plan[k]
                            for k in (
                                "bridge_version",
                                "generation_settings_sha256",
                                "schema_sha256",
                            )
                        },
                        "promptVersion": PROMPT_VERSION,
                        "promptSha256": prompt_digest(config),
                    },
                    sort_keys=True,
                ).encode()
            )
            if (
                not force
                and old_block
                and state.get("buildKey") == key
                and _assets_intact(note, state)
            ):
                result["status"] = "cached"
                logger.info("Sheet %s: cached; no rendering or AI call (tokens=0).", sheet["name"])
                continue
            result["status"] = "planned" if dry_run else "started"
            if dry_run:
                result["cellCount"] = len(evidence["cells"])
                result["objectCount"] = len(evidence["objects"])
                logger.info(
                    "Sheet %s: would render and generate context; token count unavailable until execution.",
                    sheet["name"],
                )
                continue
            generation_plan(connection, check_executable=True)
            with tempfile.TemporaryDirectory(prefix="excel-catalog-context-") as temporary:
                working = Path(temporary).resolve()
                snapshot = working / ("snapshot" + workbook.suffix)
                snapshot.write_bytes(rendering_snapshot(data))
                images_dir = working / "images"
                logger.info("Rendering sheet %s with native Excel.", sheet["name"])
                images = render_sheet(snapshot, evidence, images_dir, config)
                snapshot.unlink()
                generation = f"{key[:12]}-{uuid.uuid4().hex[:8]}"
                relative_dir = Path("img") / book_key / f"sheet-{sheet['id']}" / generation
                destination = note.parent / relative_dir
                if not destination.resolve().is_relative_to(note.parent.resolve()):
                    raise ContextError("Image output escapes the proxy note directory")
                for image in images:
                    image["relativePath"] = (relative_dir / image["image"]).as_posix()
                evidence["images"] = images
                usage_path = state_root() / "context" / "usage" / f"{uuid.uuid4().hex}.json"
                markdown = generate_markdown(
                    config,
                    evidence,
                    [images_dir / image["image"] for image in images],
                    usage_path,
                    logger,
                    profile=connection,
                    on_usage=records.append,
                )
                markdown = markdown.replace("\r\n", "\n")
                _validate_markdown(markdown, {image["relativePath"] for image in images})
                newline = "\r\n" if b"\r\n" in original else "\n"
                links = "\n".join(
                    f"- [{image['range']} ({'overview' if image['overview'] else 'detail'})]({image['relativePath']})"
                    for image in images
                )
                heading = sheet["name"].replace("\n", " ").replace("\r", " ")
                block = (
                    f"<!-- excel-catalog:begin context-{sheet['id']} -->\n"
                    f"## {heading} (sheetId: {sheet['id']})\n\n{markdown}\n\n"
                    f"### Source images\n\n{links}\n\n"
                    f"<!-- Saved-file snapshot SHA256: {digest(data)}; generated: {utc_now()}; "
                    f"provider: {plan['provider']}; requested model: {plan['model'] or 'provider default'}; prompt: {PROMPT_VERSION} -->\n"
                    f"<!-- excel-catalog:end context-{sheet['id']} -->"
                )
                block_hash = digest(block.encode())
                block = block.replace("\r\n", "\n").replace("\n", newline)
                new_text = update_text(text, sheet["id"], block)
                if include_overview:
                    new_text = patch_frontmatter(
                        new_text, {"contextStatus": "stale", "updated": utc_now()}
                    )
                if read_shared(workbook) != data:
                    raise ContextError("Workbook changed during generation; run pull again")
                encoded = (
                    b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b""
                ) + new_text.encode("utf-8")
                if note.read_bytes() != original:
                    raise ContextError(
                        "Proxy note changed during generation; generated context was not applied"
                    )
                destination.mkdir(parents=True, exist_ok=False)
                assets: dict[str, str] = {}
                for image in images:
                    image_path = images_dir / image["image"]
                    shutil.copyfile(image_path, destination / image_path.name)
                    assets[image["relativePath"]] = digest(image_path.read_bytes())
                atomic_json(destination / "evidence.json", evidence)
                temporary_note = note.with_name(f".{note.name}.{uuid.uuid4().hex}.tmp")
                try:
                    temporary_note.write_bytes(encoded)
                    if note.read_bytes() != original:
                        raise ContextError(
                            "Proxy note changed before publication; generated context was not applied"
                        )
                    temporary_note.replace(note)
                finally:
                    temporary_note.unlink(missing_ok=True)
                atomic_json(
                    state_path,
                    {
                        "buildKey": key,
                        "blockSha256": block_hash,
                        "assets": assets,
                        "snapshotSha256": digest(data),
                        "sheet": sheet["name"],
                        "usagePath": str(usage_path),
                        "generatedAt": utc_now(),
                    },
                )
                result.update(status="written", imageCount=len(images), usagePath=str(usage_path))
                logger.info("Sheet %s: context written to %s.", sheet["name"], note)

        if include_overview:
            _build_overview(
                note,
                workbook,
                data,
                selected,
                config,
                logger,
                state_dir,
                results,
                records,
                dry_run=dry_run,
                force=force,
                preview_text=preview_text,
            )

    try:
        if dry_run:
            run()
        else:
            with note_lock(note):
                run()
    except Exception as exc:
        if results:
            results[-1].update(status="error", message=str(exc))
        else:
            results.append({"status": "error", "message": str(exc)})
        logger.error("AI pull stopped: %s", exc)
    error = any(result["status"] == "error" for result in results)
    totals: dict[str, Any] = {"calls": len(records)}
    for field in ("inputTokens", "outputTokens", "cachedInputTokens", "reasoningTokens"):
        values = [record.get(field) for record in records]
        totals[field] = sum(values) if all(type(value) is int for value in values) else None
    totals["durationSeconds"] = round(
        sum(record.get("durationSeconds", 0) for record in records), 3
    )
    return {
        "command": "pull --context",
        "status": "error" if error else "success",
        "dryRun": dry_run,
        "results": results,
        "usage": totals,
    }


def _build_overview(
    note: Path,
    workbook: Path,
    data: bytes,
    selected: list[dict[str, str]],
    config: ContextConfig,
    logger: logging.Logger,
    state_dir: Path,
    results: list[dict[str, Any]],
    records: list[dict[str, Any]],
    *,
    dry_run: bool,
    force: bool,
    preview_text: str | None,
) -> None:
    original = note.read_bytes() if note.exists() else (preview_text or "").encode()
    text = original.decode("utf-8-sig")
    current_ids = {sheet["id"] for sheet in sheet_list(data)}
    removed_ids = sorted(
        set(re.findall(r"<!-- excel-catalog:begin context-([^ >]+) -->", text))
        - current_ids
        - {"workbook"}
    )
    for sheet_id in removed_ids:
        previous = existing_block(text, sheet_id)
        saved = _read_state(state_dir / f"sheet-{sheet_id}.json")
        if (
            previous
            and digest(previous.replace("\r\n", "\n").encode()) != saved.get("blockSha256")
            and not force
        ):
            raise ContextError(
                "Context for a removed sheet was edited or has no matching state; use --force only to replace it intentionally"
            )
        text = block_pattern(sheet_id).sub("", text, count=1)
    old = existing_block(text, "workbook")
    state_path = state_dir / "workbook.json"
    state = _read_state(state_path)
    if old and digest(old.replace("\r\n", "\n").encode()) != state.get("blockSha256") and not force:
        raise ContextError(
            "Workbook overview was edited or has no matching state; use --force only to replace it intentionally"
        )
    connection = resolve_profile(config)
    plan = generation_plan(connection)
    prompt = load_prompt(config, stage="workbook")
    needs_sheets = any(item["status"] == "planned" for item in results)
    retained = any(item["status"] == "retained" for item in results)
    result: dict[str, Any] = {"kind": "workbook", "status": "planned", "notePath": str(note)}
    result["removedSheetIds"] = removed_ids
    results.append(result)
    if dry_run and needs_sheets:
        return
    names = [sheet["name"] for sheet in selected]
    omitted = [sheet["name"] for sheet in sheet_list(data) if sheet["name"] not in names]
    notes = [
        {"sheet": sheet["name"], "markdown": existing_block(text, sheet["id"]) or ""}
        for sheet in selected
    ]
    evidence = {"sheet": "Workbook overview", "notes": notes, "omittedSheets": omitted}
    if len(json.dumps(evidence, ensure_ascii=False)) > config.max_input_chars:
        raise ContextError(
            "Workbook overview input exceeds generation.max_input_chars; no text was truncated"
        )
    key = digest(
        json.dumps(
            {"evidence": evidence, "prompt": prompt, "config": asdict(config), "bridge": plan},
            sort_keys=True,
        ).encode()
    )
    if not force and old and state.get("buildKey") == key:
        result["status"] = "cached"
    elif dry_run:
        return
    else:
        generation_plan(connection, check_executable=True)
        markdown = generate_markdown(
            config,
            evidence,
            [],
            state_root() / "context" / "usage" / f"{uuid.uuid4().hex}.json",
            logger,
            profile=connection,
            stage="workbook",
            on_usage=records.append,
        ).replace("\r\n", "\n")
        _validate_markdown(markdown, set())
        heading = "ブック全体の説明" if profile_name(config) == "default-ja" else "Workbook context"
        scope = json.dumps(names, ensure_ascii=False)
        omitted_text = json.dumps(omitted, ensure_ascii=False)
        block = (
            f"<!-- excel-catalog:begin context-workbook -->\n## {heading}\n\n"
            f"Analyzed sheets: {scope}\n\nOmitted sheets: {omitted_text}\n\n{markdown}\n"
            "<!-- excel-catalog:end context-workbook -->"
        )
        newline = "\r\n" if b"\r\n" in original else "\n"
        block = block.replace("\n", newline)
        if old:
            text = update_text(text, "workbook", block)
        else:
            first = re.search(r"<!-- excel-catalog:begin context-", text)
            if first:
                text = text[: first.start()] + block + newline * 2 + text[first.start() :]
            else:
                text = update_text(text, "workbook", block)
        state = {
            "buildKey": key,
            "blockSha256": digest(block.replace("\r\n", "\n").encode()),
            "generatedAt": utc_now(),
        }
        result["status"] = "written"
    if dry_run:
        return
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        fingerprint = content_fingerprint(archive)
    text = patch_frontmatter(
        text,
        {
            "contextStatus": "unverified" if retained else "partial" if omitted else "current",
            "contextSourceFingerprint": fingerprint,
            "contextAnalyzedSheets": names,
            "contextOmittedSheets": omitted,
            "contextGeneratedAt": state["generatedAt"],
        },
    )
    encoded = (b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b"") + text.encode(
        "utf-8"
    )
    if read_shared(workbook) != data:
        raise ContextError("Workbook changed during generation; run pull again")
    if note.read_bytes() != original:
        raise ContextError(
            "Proxy note changed during generation; workbook overview was not applied"
        )
    if encoded != original:
        text = patch_frontmatter(text, {"updated": utc_now()})
        encoded = (b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b"") + text.encode(
            "utf-8"
        )
        if result["status"] == "cached":
            result["status"] = "refreshed"
        temporary = note.with_name(f".{note.name}.{uuid.uuid4().hex}.tmp")
        try:
            temporary.write_bytes(encoded)
            if note.read_bytes() != original:
                raise ContextError("Proxy note changed before publication")
            temporary.replace(note)
        finally:
            temporary.unlink(missing_ok=True)
    if result["status"] == "written":
        atomic_json(state_path, state)


def run_context(
    source: SourceConfig,
    config: ContextConfig,
    *,
    workbook_selector: str,
    sheet_names: list[str],
    all_sheets: bool,
    list_only: bool,
    dry_run: bool,
    force: bool,
    logger: logging.Logger,
    note_path: Path | None = None,
    preview_text: str | None = None,
    include_overview: bool = False,
) -> dict[str, Any]:
    started = time.monotonic()
    workbook = resolve_workbook(source, workbook_selector, config)
    data = read_shared(workbook)
    sheets = sheet_list(data)
    if list_only:
        return {"command": "workbook list-sheets", "status": "success", "sheets": sheets}
    names = list(dict.fromkeys(sheet_names))
    unknown = set(names) - {sheet["name"] for sheet in sheets}
    if unknown:
        raise ContextError(f"Unknown sheet name(s): {', '.join(sorted(unknown))}")
    selected = [
        sheet
        for sheet in sheets
        if sheet["name"] in names or (all_sheets and sheet["state"] == "visible")
    ]
    if not selected:
        raise ContextError("No worksheet selected; use --sheet to include a hidden worksheet")
    result = build_context(
        source,
        workbook,
        data,
        selected,
        config,
        logger,
        dry_run=dry_run,
        force=force,
        note_path=note_path,
        preview_text=preview_text,
        include_overview=include_overview,
    )
    result["durationSeconds"] = round(time.monotonic() - started, 3)
    logger.info(
        "AI pull: result=%s | duration=%.1fs | AI calls=%d | input=%s tokens | output=%s tokens",
        result["status"],
        result["durationSeconds"],
        result["usage"]["calls"],
        result["usage"]["inputTokens"] if result["usage"]["inputTokens"] is not None else "unknown",
        result["usage"]["outputTokens"]
        if result["usage"]["outputTokens"] is not None
        else "unknown",
    )
    return result

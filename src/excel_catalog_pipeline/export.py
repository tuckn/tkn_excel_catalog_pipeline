"""One-shot workbook conversion without catalog notes or synchronization state."""

from __future__ import annotations

import json
import logging
import os
import tempfile
import uuid
from pathlib import Path
from typing import Any
from urllib.parse import quote

from .context import _validate_markdown, digest
from .context_profiles import load_prompt, profile_name
from .context_provider import (
    atomic_json,
    generate_markdown,
    generation_plan,
    resolve_profile,
    utc_now,
)
from .context_render import render_sheet
from .context_source import ContextError, extract_sheet, rendering_snapshot, sheet_list
from .models import ContextConfig
from .paths import state_root
from .shared_read import read_shared


def export_workbook(
    workbook: Path,
    output: Path,
    config: ContextConfig,
    logger: logging.Logger,
    *,
    sheet_names: list[str],
    dry_run: bool,
    force: bool,
) -> dict[str, Any]:
    workbook = workbook.expanduser().resolve()
    raw_output = output.expanduser().absolute()
    if raw_output.is_symlink():
        raise ContextError("Output must not be a symbolic link")
    output = raw_output.resolve()
    if workbook.suffix.lower() not in {".xlsx", ".xlsm"} or not workbook.is_file():
        raise ContextError("Select an existing .xlsx or .xlsm workbook")
    if output.suffix.lower() != ".md":
        raise ContextError("Output must have a .md extension")
    if output.exists() and (not output.is_file() or output.samefile(workbook)):
        raise ContextError("Output must be a Markdown file distinct from the workbook")
    original = output.read_bytes() if output.exists() else None
    if original is not None and not force:
        raise ContextError("Output already exists; use --force to replace it intentionally")
    if workbook.stat().st_size > config.max_workbook_mb * 1024 * 1024:
        raise ContextError("Workbook exceeds generation.max_workbook_mb")
    load_prompt(config)
    load_prompt(config, stage="workbook")
    data = read_shared(workbook)
    if len(data) > config.max_workbook_mb * 1024 * 1024:
        raise ContextError("Workbook exceeds generation.max_workbook_mb")
    sheets = sheet_list(data)
    unknown = set(sheet_names) - {sheet["name"] for sheet in sheets}
    if unknown:
        raise ContextError(f"Unknown sheet name(s): {', '.join(sorted(unknown))}")
    selected = [
        sheet
        for sheet in sheets
        if sheet["name"] in sheet_names or (not sheet_names and sheet["state"] == "visible")
    ]
    if not selected:
        raise ContextError("No worksheets selected")
    prepared = [
        extract_sheet(data, sheet, max_cells=config.max_cells, max_objects=config.max_objects)
        for sheet in selected
    ]
    for evidence in prepared:
        if len(json.dumps(evidence, ensure_ascii=False)) > config.max_input_chars:
            raise ContextError("Evidence exceeds generation.max_input_chars; no AI was called")
    snapshot_data = rendering_snapshot(data)
    connection = resolve_profile(config)
    plan = generation_plan(connection)
    result: dict[str, Any] = {
        "command": "export",
        "status": "planned" if dry_run else "success",
        "outputPath": str(output),
        "profile": profile_name(config),
        "sheets": [sheet["name"] for sheet in selected],
        "omittedSheets": [sheet["name"] for sheet in sheets if sheet not in selected],
        "snapshotSha256": digest(data),
        "bridgePlan": plan,
        "usage": {"calls": 0},
    }
    if dry_run:
        return result
    generation_plan(connection, check_executable=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    lock = output.with_name(output.name + ".export.lock")
    try:
        handle = lock.open("x", encoding="utf-8")
    except FileExistsError as exc:
        raise ContextError(
            "Export lock exists; inspect the active process before removing a stale lock"
        ) from exc
    records: list[dict[str, Any]] = []
    run_id = uuid.uuid4().hex
    usage_dir = state_root() / "export" / "usage" / run_id
    assets_name = f"excel-assets-{run_id}"
    destination = output.parent / assets_name
    try:
        with handle:
            handle.write(utc_now())
        with tempfile.TemporaryDirectory(prefix=".excel-export-", dir=output.parent) as temporary:
            working = Path(temporary)
            assets = working / assets_name
            assets.mkdir()
            snapshot = working / ("snapshot" + workbook.suffix)
            snapshot.write_bytes(snapshot_data)
            sections = []
            sheet_notes = []
            for sheet, evidence in zip(selected, prepared, strict=True):
                logger.info("Rendering and describing sheet %s.", sheet["name"])
                image_dir = assets / f"sheet-{sheet['id']}"
                images = render_sheet(snapshot, evidence, image_dir, config)
                for image in images:
                    image["relativePath"] = f"{assets_name}/{image_dir.name}/{image['image']}"
                evidence["images"] = images
                markdown = generate_markdown(
                    config,
                    evidence,
                    [image_dir / image["image"] for image in images],
                    usage_dir / f"sheet-{sheet['id']}.json",
                    logger,
                    profile=connection,
                    on_usage=records.append,
                )
                _validate_markdown(markdown, {image["relativePath"] for image in images})
                atomic_json(image_dir / "evidence.json", evidence)
                for pdf in image_dir.glob("*.pdf"):
                    pdf.unlink()
                heading = sheet["name"].replace("\n", " ").replace("\r", " ")
                links = "\n".join(
                    f"- [{image['range']}]({image['relativePath']})" for image in images
                )
                label = "根拠画像" if profile_name(config) == "default-ja" else "Source images"
                sections.append(f"## {heading}\n\n{markdown}\n\n### {label}\n\n{links}")
                sheet_notes.append({"sheet": sheet["name"], "markdown": markdown})
            overview_evidence = {
                "sheet": "Workbook overview",
                "notes": sheet_notes,
                "omittedSheets": result["omittedSheets"],
            }
            overview = generate_markdown(
                config,
                overview_evidence,
                [],
                usage_dir / "workbook.json",
                logger,
                profile=connection,
                stage="workbook",
                on_usage=records.append,
            )
            _validate_markdown(overview, set())
            metadata = {
                "sourceFileName": workbook.name,
                "snapshotSha256": digest(data),
                "generatedAt": utc_now(),
                "profile": profile_name(config),
                "sheets": result["sheets"],
                "omittedSheets": result["omittedSheets"],
                "promptSha256": {
                    stage: digest(load_prompt(config, stage=stage).encode())
                    for stage in ("sheet", "workbook")
                },
            }
            atomic_json(assets / "manifest.json", metadata)
            language = "ja" if profile_name(config) == "default-ja" else "en"
            scope_label = "解析対象シート" if language == "ja" else "Analyzed sheets"
            omitted_label = "対象外シート" if language == "ja" else "Omitted sheets"
            scope = f"{scope_label}: {json.dumps(result['sheets'], ensure_ascii=False)}\n\n{omitted_label}: {json.dumps(result['omittedSheets'], ensure_ascii=False)}"
            title = workbook.name.replace("\n", " ").replace("\r", " ")
            text = (
                f"# {title}\n\n{scope}\n\n{overview}\n\n"
                + "\n\n".join(sections)
                + f"\n\n[Source manifest]({quote(assets_name)}/manifest.json)\n"
            )
            staged = working / "result.md"
            staged.write_text(text, encoding="utf-8")
            current = output.read_bytes() if output.exists() else None
            if output.is_symlink() or current != original:
                raise ContextError("Output changed during generation; export was not published")
            assets.rename(destination)
            # Link creation refuses a concurrently-created destination; replacement is explicit.
            if original is None:
                os.link(staged, output)
            else:
                staged.replace(output)
            result["assetsPath"] = str(destination)
    finally:
        lock.unlink(missing_ok=True)
        result["usage"] = {"calls": len(records)}
        for field in ("inputTokens", "outputTokens", "cachedInputTokens", "reasoningTokens"):
            values = [record.get(field) for record in records]
            result["usage"][field] = (
                sum(value for value in values if isinstance(value, int))
                if all(type(value) is int for value in values)
                else None
            )
    logger.info("Export written: %s", output)
    return result

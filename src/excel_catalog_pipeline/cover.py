"""Plan embedded or rendered covers, preserving manual values and generation history."""

from __future__ import annotations

import hashlib
import logging
from dataclasses import asdict
from typing import Any

from .context_source import ContextError, sheet_list
from .cover_render import RENDER_VERSION, render_cover
from .cover_settings import validate_cover
from .models import CoverConfig, ProxyNote, SourceConfig, WorkbookInfo
from .shared_read import read_shared
from .thumbnail import CoverPlan, attachment_plan
from .thumbnail import plan_cover as embedded_cover


class CoverError(ValueError):
    """A cover selection cannot be rendered; metadata pull can still continue."""


def select_sheet(data: bytes, requested: str | None) -> dict[str, str]:
    sheets = sheet_list(data)
    if requested is not None:
        selected = next((sheet for sheet in sheets if sheet["name"] == requested), None)
        if selected is None:
            raise CoverError("Requested cover worksheet does not exist")
        if selected["state"] != "visible":
            raise CoverError("Cover worksheet must be visible")
        if not selected["part"].startswith("xl/worksheets/"):
            raise CoverError("Cover requires a worksheet, not a chart sheet")
        return selected
    selected = next(
        (
            sheet
            for sheet in sheets
            if sheet["state"] == "visible" and sheet["part"].startswith("xl/worksheets/")
        ),
        None,
    )
    if selected is None:
        raise CoverError("Workbook has no visible worksheet for its cover")
    return selected


def plan_cover(
    workbook: WorkbookInfo,
    source: SourceConfig,
    note: ProxyNote | None,
    entry: dict[str, Any] | None,
    *,
    options: CoverConfig,
    render: bool,
    max_workbook_mb: int = 100,
) -> CoverPlan:
    current = note.frontmatter.get("cover", "") if note else ""
    managed = str((entry or {}).get("managedCover", ""))
    previous = (entry or {}).get("coverGeneration", {})
    previous = previous if isinstance(previous, dict) else {}
    if current and current != managed:
        return CoverPlan(value=current, generation=previous or None, details={"status": "manual"})
    mode = options.mode if options.mode != "auto" else previous.get("mode", "embedded")
    if mode == "embedded":
        result = embedded_cover(workbook, source, note, entry)
        if not result.warning and (previous or options.mode == "embedded"):
            result.generation = {"mode": "embedded"}
        result.details = {"mode": "embedded", "status": "warning" if result.warning else "ready"}
        if options.mode == "auto" and any(
            value is not None for value in (options.sheet, options.range, options.width)
        ):
            result.warning = "Cover sheet options require --cover sheet or an existing sheet cover"
            return CoverPlan(
                current,
                managed,
                warning=result.warning,
                generation=previous or None,
                details={"mode": "embedded", "status": "warning"},
            )
        return result
    details: dict[str, Any] = {"mode": "sheet"}
    try:
        if mode != "sheet":
            raise CoverError("Unknown saved cover mode; select --cover embedded or --cover sheet")
        # Explicit non-null options override saved per-workbook choices independently.
        inherited = previous.get("settings", {})
        inherited = inherited if isinstance(inherited, dict) else {}
        merged = {"sheet": None, "range": "A1:Q50", "width": 2400, **inherited, "mode": "sheet"}
        merged.update(
            {
                key: value
                for key, value in asdict(options).items()
                if key != "mode" and value is not None
            }
        )
        settings = validate_cover(merged)
        address, width = settings.range or "A1:Q50", settings.width or 2400
        if workbook.path.stat().st_size > max_workbook_mb * 1024 * 1024:
            raise CoverError("Workbook exceeds generation.max_workbook_mb")
        data = read_shared(workbook.path)
        sheet = select_sheet(data, settings.sheet)
        details.update(sheet=sheet["name"], range=address, width=width)
        identity = {
            "mode": "sheet",
            "rendererVersion": RENDER_VERSION,
            "settings": {"sheet": settings.sheet, "range": address, "width": width},
            "sheetId": sheet["id"],
            "sheetName": sheet["name"],
            "workbookHash": hashlib.sha256(data).hexdigest(),
        }
        cached_hash = previous.get("imageHash", "")
        if (
            all(previous.get(key) == value for key, value in identity.items())
            and isinstance(cached_hash, str)
            and len(cached_hash) == 64
            and all(char in "0123456789abcdef" for char in cached_hash)
        ):
            cached = source.note_root / "img" / f"excel-cover-{cached_hash}.png"
            if cached.is_file():
                png = cached.read_bytes()
                if hashlib.sha256(png).hexdigest() == cached_hash:
                    result = attachment_plan(png, source)
                    result.generation = previous
                    result.details = {**details, "status": "cached"}
                    return result
        if not render:
            return CoverPlan(
                current,
                managed,
                generation=previous or None,
                planned=True,
                details={**details, "status": "planned"},
            )
        logging.getLogger(__name__).info(
            "Rendering cover: sheet=%s range=%s width=%d.", sheet["name"], address, width
        )
        png = render_cover(data, workbook.extension, sheet["name"], address, width)
        result = attachment_plan(png, source)
        result.generation = {**identity, "imageHash": hashlib.sha256(png).hexdigest()}
        result.details = {**details, "status": "rendered"}
        return result
    except Exception as exc:
        # COM and decoder errors can contain private paths or document contents.
        reason = str(exc) if isinstance(exc, (CoverError, ContextError)) else type(exc).__name__
        return CoverPlan(
            current,
            managed,
            warning=f"Sheet cover unavailable: {reason}",
            generation=previous or None,
            details={**details, "status": "warning"},
        )

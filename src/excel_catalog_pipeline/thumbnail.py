"""Read package thumbnails and plan local Obsidian cover attachments."""

from __future__ import annotations

import hashlib
import io
import posixpath
import struct
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET

from .adapters.markdown import find_obsidian_vault_root
from .models import ProxyNote, SourceConfig, WorkbookInfo
from .shared_read import read_shared

THUMBNAIL_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
MAX_THUMBNAIL_BYTES = 16 * 1024 * 1024
MAX_PIXELS = 16_000_000


class ThumbnailError(ValueError):
    """An optional thumbnail cannot safely be used."""


def read_thumbnail(archive: zipfile.ZipFile) -> bytes | None:
    """Follow only the package thumbnail relationship; never select worksheet art."""
    names = archive.namelist()
    target = ""
    if "_rels/.rels" in names:
        root = ET.fromstring(archive.read("_rels/.rels"))
        for rel in root:
            if rel.get("Type") != THUMBNAIL_REL:
                continue
            raw = rel.get("Target", "")
            uri = urlsplit(raw)
            if rel.get("TargetMode", "Internal") != "Internal" or uri.scheme or uri.netloc:
                raise ThumbnailError("External thumbnail relationships are unsupported")
            target = posixpath.normpath(unquote(uri.path).lstrip("/"))
            if target in ("", ".", "..") or target.startswith("../") or "\\" in target:
                raise ThumbnailError("Invalid thumbnail relationship target")
            break
    if not target:
        target = next(
            (
                name
                for name in sorted(names)
                if name.casefold()
                in {f"docprops/thumbnail.{ext}" for ext in ("png", "jpeg", "jpg", "emf", "wmf")}
            ),
            "",
        )
    if not target:
        return None
    if target not in names:
        raise ThumbnailError("Thumbnail relationship points to a missing package entry")
    if archive.getinfo(target).file_size > MAX_THUMBNAIL_BYTES:
        raise ThumbnailError("Thumbnail exceeds the 16 MiB limit")
    return archive.read(target)


def _placeable_wmf(data: bytes) -> bytes:
    """Add a placeable header for Excel's standard WMF using its initial window."""
    if not data.startswith((b"\x01\x00\x09\x00", b"\x02\x00\x09\x00")):
        return data
    offset = 18
    x = y = 0
    while offset + 6 <= len(data):
        words, function = struct.unpack_from("<IH", data, offset)
        end = offset + words * 2
        if words < 3 or end > len(data):
            raise ThumbnailError("Malformed WMF record")
        if function in (0x020B, 0x020C):
            if words < 5:
                raise ThumbnailError("Malformed WMF window record")
            vertical, horizontal = struct.unpack_from("<hh", data, offset + 6)
            if function == 0x020B:
                x, y = horizontal, vertical
            else:
                if horizontal <= 0 or vertical <= 0:
                    raise ThumbnailError("Unsupported WMF window dimensions")
                header = struct.pack(
                    "<IHhhhhHI", 0x9AC6CDD7, 0, x, y, x + horizontal, y + vertical, 72, 0
                )
                checksum = 0
                for word in struct.unpack("<10H", header):
                    checksum ^= word
                # Pillow's placeable WMF reader expects a memory metafile header.
                return header + struct.pack("<H", checksum) + b"\x01\x00" + data[2:]
        offset = end
    raise ThumbnailError("WMF has no initial window dimensions")


def thumbnail_png(data: bytes) -> bytes:
    """Convert locally using Pillow (Windows also supports EMF/WMF)."""
    from PIL import Image
    from PIL.WmfImagePlugin import WmfStubImageFile

    try:
        with Image.open(io.BytesIO(_placeable_wmf(data))) as picture:
            width, height = picture.size
            if width <= 0 or height <= 0:
                raise ThumbnailError("Thumbnail has invalid dimensions")
            if isinstance(picture, WmfStubImageFile):
                # Rasterize vector previews at the output size, avoiding large allocations.
                scale = min(1.0, 1200 / max(width, height))
                dpi = picture.info.get("dpi", 72)
                xdpi, ydpi = dpi if isinstance(dpi, tuple) else (dpi, dpi)
                picture.load(dpi=(xdpi * scale, ydpi * scale))
            else:
                if width * height > MAX_PIXELS:
                    raise ThumbnailError("Thumbnail exceeds the pixel limit")
                picture.load()
            picture.thumbnail((1200, 1200))
            output = io.BytesIO()
            picture.convert("RGBA").save(output, format="PNG")
            return output.getvalue()
    except (
        OSError,
        ValueError,
        SyntaxError,
        struct.error,
        ZeroDivisionError,
        Image.DecompressionBombError,
    ) as exc:
        raise ThumbnailError(f"Thumbnail conversion failed ({type(exc).__name__})") from exc


@dataclass
class CoverPlan:
    value: Any = ""
    managed: str = ""
    path: Path | None = None
    png: bytes = b""
    asset_changed: bool = False
    warning: str = ""

    def write(self) -> None:
        """Commit the attachment before the note; never delete old attachments."""
        if self.path is None or not self.asset_changed:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_name(f".{self.path.name}.{uuid.uuid4().hex[:8]}.tmp")
        try:
            temporary.write_bytes(self.png)
            temporary.replace(self.path)
        finally:
            temporary.unlink(missing_ok=True)


def plan_cover(
    workbook: WorkbookInfo,
    source: SourceConfig,
    note: ProxyNote | None,
    entry: dict[str, Any] | None,
) -> CoverPlan:
    current = note.frontmatter.get("cover", "") if note else ""
    managed = str((entry or {}).get("managedCover", ""))
    if current and current != managed:
        return CoverPlan(value=current)
    try:
        try:
            with zipfile.ZipFile(workbook.path) as archive:
                data = read_thumbnail(archive)
        except PermissionError:
            with zipfile.ZipFile(io.BytesIO(read_shared(workbook.path))) as archive:
                data = read_thumbnail(archive)
        if data is None:
            return CoverPlan()
        png = thumbnail_png(data)
        digest = hashlib.sha256(png).hexdigest()
        target = source.note_root / "img" / f"excel-cover-{digest}.png"
        vault = find_obsidian_vault_root(source.note_root)
        relative = target.resolve().relative_to(vault).as_posix()
        if any(char in relative for char in "[]|#^\r\n"):
            raise ThumbnailError("Attachment path contains unsupported Obsidian link characters")
        link = f"[[{relative}]]"
        changed = not target.is_file() or target.read_bytes() != png
        return CoverPlan(link, link, target, png, changed)
    except (
        OSError,
        ValueError,
        RuntimeError,
        ImportError,
        zipfile.BadZipFile,
        ET.ParseError,
    ) as exc:
        # Do not expose image contents or decoder diagnostics in reports.
        return CoverPlan(current, managed, warning=f"Thumbnail unavailable ({type(exc).__name__})")

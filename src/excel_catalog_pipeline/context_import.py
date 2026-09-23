"""Import an existing sheet Markdown and its relative image links without AI."""

from __future__ import annotations

import json
import logging
import re
import uuid
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit

from .adapters.markdown import read_note
from .context import (
    _assets_intact,
    _read_state,
    digest,
    existing_block,
    find_proxy,
    note_lock,
    resolve_workbook,
    update_text,
)
from .context_provider import atomic_json, utc_now
from .context_source import ContextError, sheet_list
from .models import ContextConfig, SourceConfig
from .note_layout import migrate_layout
from .paths import state_root
from .shared_read import read_shared

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
LINK = re.compile(
    r"(?P<prefix>!?\[[^\]\n]*\]\()(?P<target><[^>\n]+>|[^\s)]+)(?P<tail>(?:\s+['\"][^\n]*?['\"])?\))"
)
REFERENCE = re.compile(r"^(?P<prefix> {0,3}\[[^\]\n]+\]:\s*)(?P<target><[^>\n]+>|\S+)(?P<tail>.*)$")


def convert_body(
    body: str, sheet: str, sheet_id: str, input_dir: Path, relative_dir: Path
) -> tuple[str, dict[str, bytes]]:
    """Demote ATX headings and copy only explicitly referenced local image bytes."""
    if "<!-- excel-catalog:" in body:
        raise ContextError("Input contains reserved management markers")
    assets: dict[str, bytes] = {}

    def rewrite(match: re.Match[str]) -> str:
        target = match.group("target").strip("<>")
        parsed = urlsplit(target)
        if target.startswith("#") or parsed.scheme in {"https", "http", "mailto"}:
            return match.group(0)
        if parsed.scheme or parsed.netloc or target.startswith(("/", "\\")) or "\\" in target:
            raise ContextError(f"Only relative local image paths are imported: {target}")
        asset = (input_dir / unquote(parsed.path)).resolve()
        if (
            not asset.is_relative_to(input_dir.resolve())
            or asset.suffix.lower() not in IMAGE_EXTENSIONS
        ):
            raise ContextError(
                f"Local link must name an image inside the Markdown directory: {target}"
            )
        if parsed.query or not asset.is_file():
            raise ContextError(f"Image not found or unsupported query: {target}")
        content = asset.read_bytes()
        relative = (relative_dir / (digest(content)[:20] + asset.suffix.lower())).as_posix()
        assets[relative] = content
        fragment = "#" + parsed.fragment if parsed.fragment else ""
        return match.group("prefix") + relative + fragment + match.group("tail")

    output = []
    fence = ""
    titles = 0
    for line in body.replace("\r\n", "\n").splitlines(keepends=True):
        marker = re.match(r"^ {0,3}(`{3,}|~{3,})", line)
        if marker:
            run = marker.group(1)
            if not fence:
                fence = run
            elif run[0] == fence[0] and len(run) >= len(fence):
                fence = ""
            output.append(line)
            continue
        if fence:
            output.append(line)
            continue
        heading = re.match(r"^(#{1,6})[ \t]+(.+?)(?:\n)?$", line)
        if heading:
            level = len(heading.group(1))
            if level == 6:
                raise ContextError(
                    "Input H6 cannot be demoted; simplify the heading hierarchy first"
                )
            if level == 1:
                titles += 1
                line = f"## {sheet} (sheetId: {sheet_id})\n"
            else:
                line = "#" + line
        # Inline code is source text, not a link to an attachment.
        pieces = re.split(r"(`+[^`]*`+)", line)
        line = "".join(
            piece if index % 2 else LINK.sub(rewrite, piece) for index, piece in enumerate(pieces)
        )
        if REFERENCE.match(line):
            line = REFERENCE.sub(rewrite, line.rstrip("\n")) + "\n"
        output.append(line)
    if fence:
        raise ContextError("Input has an unclosed code fence")
    if titles != 1:
        raise ContextError("Input Markdown must have exactly one H1 title (ATX headings)")
    return "".join(output).strip(), assets


def _replace_note(path: Path, content: bytes, expected: bytes) -> None:
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_bytes(content)
        if path.read_bytes() != expected:
            raise ContextError("Proxy note changed during import; it was not overwritten")
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def import_context(
    source: SourceConfig,
    config: ContextConfig,
    *,
    workbook_selector: str,
    sheet_name: str,
    markdown_path: Path,
    dry_run: bool,
    force: bool,
    description: str | None,
    logger: logging.Logger,
) -> dict[str, Any]:
    workbook = resolve_workbook(source, workbook_selector, config)
    data = read_shared(workbook)
    sheets = [sheet for sheet in sheet_list(data) if sheet["name"] == sheet_name]
    if len(sheets) != 1:
        raise ContextError(f"Unknown or ambiguous sheet: {sheet_name}")
    sheet = sheets[0]
    note, book_key = find_proxy(source, workbook, data)
    input_path = markdown_path.expanduser().resolve()
    if input_path == note.resolve():
        raise ContextError("Input Markdown must be separate from the proxy note")
    input_bytes = input_path.read_bytes()
    imported = read_note(input_path)
    if input_path.read_bytes() != input_bytes:
        raise ContextError("Input Markdown changed while reading; retry the import")
    proxy = read_note(note)
    for field, expected in {
        "sourceId": proxy.source_id,
        "sourceSheet": sheet_name,
        "sourceSheetId": sheet["id"],
    }.items():
        if field in imported.frontmatter and str(imported.frontmatter[field]) != expected:
            raise ContextError(f"Input {field} does not match the selected workbook/sheet")
    if (
        "sourceFileName" in imported.frontmatter
        and str(imported.frontmatter["sourceFileName"]).replace("\\", "/")
        != workbook.relative_to(source.path.resolve()).as_posix()
    ):
        raise ContextError("Input sourceFileName does not match the selected workbook")
    relative_dir = (
        Path("img") / book_key / f"sheet-{sheet['id']}" / f"import-{digest(input_bytes)[:16]}"
    )
    content, image_bytes = convert_body(
        imported.body, sheet_name, sheet["id"], input_path.parent, relative_dir
    )
    asset_hashes = {relative: digest(value) for relative, value in image_bytes.items()}
    import_key = digest(
        json.dumps(
            {"markdownSha256": digest(input_bytes), "assets": asset_hashes}, sort_keys=True
        ).encode()
    )
    state_dir = state_root() / "context" / digest(str(note.resolve()).encode())[:24]
    state_path = state_dir / f"sheet-{sheet['id']}.json"
    snapshot_hash = str(imported.frontmatter.get("sourceSha256", ""))
    if snapshot_hash and snapshot_hash.lower() != digest(data):
        logger.warning(
            "Imported content describes an older/different saved snapshot; its provenance is retained."
        )
    # Validate all output locations before creating any artifact.
    for relative in [*image_bytes, (relative_dir / "provenance.json").as_posix()]:
        if not (note.parent / relative).resolve().is_relative_to(note.parent.resolve()):
            raise ContextError("Image destination escapes the proxy directory")
    for relative, content_bytes in image_bytes.items():
        destination = note.parent / relative
        if destination.exists() and destination.read_bytes() != content_bytes and not force:
            raise ContextError(
                "An imported image was modified; use --force only to restore it intentionally"
            )
    result: dict[str, Any] = {
        "command": "context import",
        "status": "success",
        "dryRun": dry_run,
        "notePath": str(note),
        "sheet": sheet_name,
        "imageCount": len(image_bytes),
        "usage": {"calls": 0, "inputTokens": 0, "outputTokens": 0},
    }

    def run() -> None:
        original = note.read_bytes()
        text = original.decode("utf-8-sig")
        state = _read_state(state_path)
        old = existing_block(text, sheet["id"])
        if (
            old
            and digest(old.replace("\r\n", "\n").encode()) != state.get("blockSha256")
            and not force
        ):
            raise ContextError(
                "Existing context was edited or has no matching state; use --force only to replace it intentionally"
            )
        newline = "\r\n" if b"\r\n" in original else "\n"
        same_import = (
            not force
            and old is not None
            and state.get("origin") == "import"
            and state.get("importKey") == import_key
        )
        if same_import:
            block = old or ""
        else:
            block = (
                f"<!-- excel-catalog:begin context-{sheet['id']} -->\n{content}\n\n"
                f"<!-- Imported Markdown SHA256: {digest(input_bytes)}; no AI generation -->\n"
                f"<!-- excel-catalog:end context-{sheet['id']} -->"
            ).replace("\n", newline)
        migrated = migrate_layout(text, str(workbook), description=description)
        updated = update_text(migrated, sheet["id"], block)
        encoded = (
            b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b""
        ) + updated.encode("utf-8")
        if same_import and encoded == original and _assets_intact(note, state):
            result["result"] = "unchanged"
            return
        result["result"] = "planned" if dry_run else "imported"
        if dry_run:
            return
        if note.read_bytes() != original:
            raise ContextError("Proxy note changed during import")
        backup = state_dir / "backups" / f"{uuid.uuid4().hex}.md"
        backup.parent.mkdir(parents=True, exist_ok=True)
        backup.write_bytes(original)
        for relative, image_content in image_bytes.items():
            destination = note.parent / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            if destination.exists():
                if destination.read_bytes() == image_content:
                    continue
                if not force:
                    raise ContextError(
                        "An imported image was modified; use --force only to restore it intentionally"
                    )
            temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex}.tmp")
            try:
                temporary.write_bytes(image_content)
                temporary.replace(destination)
            finally:
                temporary.unlink(missing_ok=True)
        provenance = {
            "origin": "import",
            "markdownSha256": digest(input_bytes),
            "sourceMetadata": imported.frontmatter,
            "assets": asset_hashes,
            "aiCalls": 0,
        }
        atomic_json(note.parent / relative_dir / "provenance.json", provenance)
        new_state = {
            **provenance,
            "importKey": import_key,
            "blockSha256": digest(block.replace("\r\n", "\n").encode()),
            "snapshotSha256": snapshot_hash,
            "sheet": sheet_name,
            "importedAt": utc_now(),
            "backupPath": str(backup),
        }
        _replace_note(note, encoded, original)
        try:
            atomic_json(state_path, new_state)
        except Exception:
            # Roll back only our own publication; never overwrite a concurrent editor.
            if note.read_bytes() == encoded:
                _replace_note(note, original, encoded)
            raise
        result["backupPath"] = str(backup)
        result["statePath"] = str(state_path)

    if dry_run:
        run()
    else:
        with note_lock(note):
            run()
    logger.info(
        "Context import: %s | sheet=%s | images=%d | AI calls=0 | input=0 tokens | output=0 tokens",
        result["result"],
        sheet_name,
        len(image_bytes),
    )
    return result

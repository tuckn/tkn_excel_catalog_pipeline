"""Load and validate application-owned proxy-note resources."""

from __future__ import annotations

import re
from dataclasses import dataclass
from functools import cache
from importlib.resources import files

PROFILE_PATTERN = re.compile(r"^[a-z0-9][a-z0-9._-]*$")
PLACEHOLDER_PATTERN = re.compile(
    r"{{\s*(?P<name>[a-z][a-z0-9_]*)"
    r"(?:\s*\|\s*(?P<default>[^{}]*?))?\s*}}"
)
MANAGED_BLOCK_PATTERN = re.compile(
    r"<!-- excel-catalog:begin (?P<name>[a-z0-9-]+) -->\n"
    r"(?P<content>.*?)"
    r"<!-- excel-catalog:end (?P=name) -->",
    re.DOTALL,
)
HEADING_PATTERN = re.compile(r"\A##\s+(?P<heading>[^\r\n]+)\r?\n")


class NoteResourceError(ValueError):
    """An application-owned note resource is missing or malformed."""


@dataclass(frozen=True)
class ManagedBlock:
    """One rendered section controlled by excel-catalog markers."""

    name: str
    heading: str
    text: str


@dataclass(frozen=True)
class NoteTemplate:
    """Validated proxy-note Markdown template for one profile."""

    profile: str
    text: str
    placeholders: frozenset[str]
    managed_names: tuple[str, ...]

    def render(self, values: dict[str, str]) -> str:
        missing = sorted(self.placeholders - values.keys())
        if missing:
            raise NoteResourceError(
                f"Note profile {self.profile!r} is missing renderer values: {', '.join(missing)}"
            )
        def replace(match: re.Match[str]) -> str:
            value = values[match.group("name")]
            default = match.group("default")
            return value if value or default is None else default.strip()

        return PLACEHOLDER_PATTERN.sub(replace, self.text)


def _validate_template(profile: str, text: str) -> NoteTemplate:
    if not text.strip():
        raise NoteResourceError(f"Note profile {profile!r} has an empty template.md")

    managed_names: list[str] = []
    for match in MANAGED_BLOCK_PATTERN.finditer(text):
        name = match.group("name")
        if name in managed_names:
            raise NoteResourceError(
                f"Note profile {profile!r} repeats managed section {name!r}"
            )
        if HEADING_PATTERN.match(match.group("content")) is None:
            raise NoteResourceError(
                f"Managed section {name!r} in note profile {profile!r} must start with an H2"
            )
        managed_names.append(name)

    if not managed_names:
        raise NoteResourceError(f"Note profile {profile!r} defines no managed sections")

    return NoteTemplate(
        profile=profile,
        text=text.replace("\r\n", "\n"),
        placeholders=frozenset(
            match.group("name") for match in PLACEHOLDER_PATTERN.finditer(text)
        ),
        managed_names=tuple(managed_names),
    )


@cache
def load_note_template(profile: str) -> NoteTemplate:
    """Load a profile template from the installed Python package."""

    if PROFILE_PATTERN.fullmatch(profile) is None:
        raise NoteResourceError(f"Invalid note profile name: {profile!r}")
    resource = files("excel_catalog_pipeline").joinpath(
        "note_profiles", profile, "template.md"
    )
    try:
        text = resource.read_text(encoding="utf-8")
    except (FileNotFoundError, ModuleNotFoundError) as exc:
        raise NoteResourceError(f"Unknown note profile: {profile!r}") from exc
    return _validate_template(profile, text)


def managed_blocks(rendered: str, template: NoteTemplate) -> tuple[ManagedBlock, ...]:
    """Parse rendered managed blocks and verify their template-defined order."""

    blocks: list[ManagedBlock] = []
    for match in MANAGED_BLOCK_PATTERN.finditer(rendered):
        heading_match = HEADING_PATTERN.match(match.group("content"))
        if heading_match is None:  # Defensive: the unrendered template was already validated.
            raise NoteResourceError(
                f"Managed section {match.group('name')!r} no longer starts with an H2"
            )
        blocks.append(
            ManagedBlock(
                name=match.group("name"),
                heading=heading_match.group("heading").strip(),
                text=match.group(0),
            )
        )
    names = tuple(block.name for block in blocks)
    if names != template.managed_names:
        raise NoteResourceError(
            f"Rendered note profile {template.profile!r} changed its managed-section contract"
        )
    return tuple(blocks)

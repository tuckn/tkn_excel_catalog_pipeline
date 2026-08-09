"""Field-level three-way comparison and resolution."""

from __future__ import annotations

from .models import METADATA_FIELDS, Direction, FieldDecision


def compare_field(field: str, base: str, source: str, note: str) -> FieldDecision:
    direction: Direction
    if source == base and note == base:
        direction = "unchanged"
    elif source != base and note == base:
        direction = "pull"
    elif source == base and note != base:
        direction = "push"
    elif source == note:
        direction = "converged"
    else:
        direction = "conflict"
    return FieldDecision(field=field, base=base, source=source, note=note, direction=direction)


def compare_metadata(
    base: dict[str, str],
    source: dict[str, str],
    note: dict[str, str],
) -> list[FieldDecision]:
    return [
        compare_field(field, base.get(field, ""), source.get(field, ""), note.get(field, ""))
        for field in METADATA_FIELDS
    ]


def resolve_for_pull(
    decisions: list[FieldDecision],
    *,
    preference: str | None = None,
) -> dict[str, str]:
    values: dict[str, str] = {}
    for decision in decisions:
        if preference == "source":
            values[decision.field] = decision.source
        elif decision.direction == "conflict":
            if preference == "note":
                values[decision.field] = decision.note
            else:
                values[decision.field] = decision.note
        elif decision.direction == "pull":
            values[decision.field] = decision.source
        else:
            values[decision.field] = decision.note
    return values


def core_updates_for_push(
    decisions: list[FieldDecision],
    *,
    preference: str | None = None,
) -> dict[str, str]:
    updates: dict[str, str] = {}
    for decision in decisions:
        if decision.field == "sourceFileName":
            continue
        if decision.direction == "push" or (
            decision.direction == "conflict" and preference == "note"
        ):
            updates[decision.field] = decision.note
    return updates


def direction_fields(decisions: list[FieldDecision], direction: str) -> list[str]:
    return [decision.field for decision in decisions if decision.direction == direction]

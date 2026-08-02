"""Run reports and compact machine-readable summaries."""

from __future__ import annotations

import csv
import json
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

from .models import Action
from .paths import runs_root


def new_run_id(command: str) -> str:
    stamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    return f"{stamp}-{command}"


def write_report(
    command: str,
    actions: list[Action],
    *,
    write_enabled: bool,
    report_root: Path | None = None,
    extra: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], Path]:
    run_id = new_run_id(command)
    run_dir = (report_root or runs_root()) / run_id
    suffix = 1
    while run_dir.exists():
        run_dir = (report_root or runs_root()) / f"{run_id}-{suffix}"
        suffix += 1
    run_dir.mkdir(parents=True, exist_ok=False)
    rows = [action.as_dict() for action in actions]
    counts = dict(Counter(action.status for action in actions))
    conflicts = sum(bool(action.conflict_fields) for action in actions)
    errors = sum(action.status.endswith("error") for action in actions)
    changed_statuses = {
        "would-create",
        "created",
        "would-update",
        "updated",
        "would-write",
        "written",
        "would-adopt",
        "adopted",
        "would-rename",
        "renamed",
    }
    changed = sum(action.status in changed_statuses for action in actions)
    summary: dict[str, Any] = {
        "status": "error" if errors else "conflict" if conflicts else "success",
        "command": command,
        "changed": changed,
        "conflicts": conflicts,
        "errors": errors,
        "writeEnabled": write_enabled,
        "statusCounts": counts,
        "reportPath": str(run_dir),
    }
    if extra:
        summary.update(extra)
    (run_dir / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    (run_dir / "details.json").write_text(
        json.dumps(rows, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    with (run_dir / "actions.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "status",
                "sourceRoot",
                "sourcePath",
                "notePath",
                "workbookId",
                "changedFields",
                "conflictFields",
                "message",
                "details",
            ],
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    **row,
                    "changedFields": ";".join(row["changedFields"]),
                    "conflictFields": ";".join(row["conflictFields"]),
                    "details": json.dumps(row["details"], ensure_ascii=False, sort_keys=True),
                }
            )
    return summary, run_dir

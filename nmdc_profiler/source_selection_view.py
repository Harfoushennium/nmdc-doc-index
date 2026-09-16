from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, Sequence

from .excel_bridge import load_approved_state, load_latest_stage
from .source_selection import read_source_exclusions


FIELDS: Sequence[str] = (
    "Owner Choice",
    "Source File",
    "Source Family",
    "Current Status",
    "Selection Reason",
    "Last Processed Run",
)


def _write_csv(path: Path, rows: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(FIELDS), lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in FIELDS})


def export_source_selection(state_dir: Path, exchange_dir: Path, config_dir: Path) -> Path:
    """Write a simple owner-facing source list from the latest staged/approved manifest."""
    staged = load_latest_stage(state_dir)
    approved = load_approved_state(state_dir)
    manifest = staged.get("manifest", {}) if staged.get("run_id") else approved.get("manifest", {})
    exclusions = read_source_exclusions(Path(config_dir) / "source_exclusions.csv")

    rows = []
    for item in manifest.get("files", []) if isinstance(manifest, dict) else []:
        rel = str(item.get("relative_path", "") or "").replace("\\", "/")
        if not rel:
            continue
        normalized = rel.casefold().lstrip("/")
        status = str(item.get("selection_status", "") or "")
        excluded = normalized in exclusions
        family = rel.split("/", 1)[0] if "/" in rel else ""
        reason = (
            exclusions[normalized].get("Reason", "Owner excluded source")
            if excluded
            else str(item.get("selection_exclusion_reason", "") or "")
        )
        rows.append(
            {
                "Owner Choice": "EXCLUDE" if excluded else "INCLUDE",
                "Source File": rel,
                "Source Family": family,
                "Current Status": status,
                "Selection Reason": reason,
                "Last Processed Run": str(item.get("last_processed_run", "") or ""),
            }
        )

    rows.sort(key=lambda row: str(row["Source File"]).casefold())
    target = Path(exchange_dir) / "source_selection.csv"
    _write_csv(target, rows)
    return target

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, Sequence, Set

from .excel_bridge import load_approved_state, load_latest_stage
from .source_selection import read_source_exclusions


FIELDS: Sequence[str] = (
    "Include in Index?",
    "Project No.",
    "Source File",
    "Source Family",
    "Current Status",
    "Owner Note",
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


def _normalize_source(value: Any) -> str:
    return str(value or "").replace("\\", "/").lstrip("/").casefold()


def _project_numbers_by_source(*record_sets: Sequence[Mapping[str, Any]]) -> Dict[str, Set[str]]:
    projects: Dict[str, Set[str]] = {}
    for records in record_sets:
        for row in records:
            source = _normalize_source(row.get("Source File", ""))
            project = str(row.get("Project No.", "") or "").strip()
            if source and project:
                projects.setdefault(source, set()).add(project)
    return projects


def _project_sort_key(value: str) -> tuple[int, str]:
    text = str(value or "").strip()
    try:
        return (0, f"{int(text):012d}")
    except ValueError:
        return (1, text.casefold())


def export_source_selection(state_dir: Path, exchange_dir: Path, config_dir: Path) -> Path:
    """Write the source-scope list consumed by the Pending Update checkbox panel.

    Project No. is included so the owner can identify the project directly beside
    each source workbook.  The project mapping is built from both staged and
    approved records so an intentionally excluded source remains identifiable
    even after its staged records have been removed from the proposal.
    """
    staged = load_latest_stage(state_dir)
    approved = load_approved_state(state_dir)
    manifest = staged.get("manifest", {}) if staged.get("run_id") else approved.get("manifest", {})
    exclusions = read_source_exclusions(Path(config_dir) / "source_exclusions.csv")

    projects = _project_numbers_by_source(
        list(staged.get("records", []) or []),
        list(approved.get("records", []) or []),
    )

    rows = []
    for item in manifest.get("files", []) if isinstance(manifest, dict) else []:
        rel = str(item.get("relative_path", "") or "").replace("\\", "/")
        if not rel:
            continue
        normalized = _normalize_source(rel)
        status = str(item.get("selection_status", "") or "")
        excluded = normalized in exclusions
        family = rel.split("/", 1)[0] if "/" in rel else ""
        owner_note = exclusions[normalized].get("Reason", "") if excluded else ""
        automatic_reason = str(item.get("selection_exclusion_reason", "") or "")
        project_text = "; ".join(sorted(projects.get(normalized, set()), key=_project_sort_key))
        rows.append(
            {
                "Include in Index?": "FALSE" if excluded else "TRUE",
                "Project No.": project_text,
                "Source File": rel,
                "Source Family": family,
                "Current Status": status,
                "Owner Note": owner_note,
                "Selection Reason": automatic_reason,
                "Last Processed Run": str(item.get("last_processed_run", "") or ""),
            }
        )

    rows.sort(key=lambda row: (_project_sort_key(str(row["Project No."]).split(";", 1)[0]), str(row["Source File"]).casefold()))
    target = Path(exchange_dir) / "source_selection.csv"
    _write_csv(target, rows)
    return target

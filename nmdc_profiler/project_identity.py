from __future__ import annotations

import csv
from dataclasses import replace
from pathlib import Path
from typing import Dict, Mapping, Sequence, Tuple

from .full_extractor import FullWorkItem


def load_project_identity_overrides(path: Path) -> Dict[str, Dict[str, str]]:
    """Load owner-approved project identities keyed by exact source path."""
    if not path.exists():
        return {}
    with path.open(newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out: Dict[str, Dict[str, str]] = {}
    for row in rows:
        source = (row.get("Source File") or "").strip()
        project = (row.get("Approved Project No.") or "").strip()
        status = (row.get("Decision Status") or "").strip().upper()
        if not source or not project or status != "APPROVED":
            continue
        out[source] = {k: (v or "") for k, v in row.items()}
    return out


def apply_project_identity_overrides(
    work_items: Sequence[FullWorkItem],
    overrides: Mapping[str, Mapping[str, str]],
) -> Tuple[list[FullWorkItem], list[dict[str, str]]]:
    """Apply explicit owner decisions without changing source files.

    Returns adjusted work items plus an audit list suitable for logging/reporting.
    """
    adjusted: list[FullWorkItem] = []
    audit: list[dict[str, str]] = []
    for item in work_items:
        override = overrides.get(item.source_file)
        if not override:
            adjusted.append(item)
            continue
        approved = str(override.get("Approved Project No.", "")).strip()
        if not approved:
            adjusted.append(item)
            continue
        adjusted.append(replace(item, project_number=approved))
        audit.append(
            {
                "Source File": item.source_file,
                "Previous Project No.": item.project_number,
                "Approved Project No.": approved,
                "Decision": str(override.get("Owner Decision", "")),
                "Owner Note": str(override.get("Owner Note", "")),
                "Approved Date": str(override.get("Approved Date", "")),
                "Flag": "OWNER_PROJECT_OVERRIDE_APPLIED",
            }
        )
    return adjusted, audit

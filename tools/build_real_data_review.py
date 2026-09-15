"""Build the deterministic CSV exchange used by the owner review workbook.

This is intentionally a review-package builder, not an approval operation.  It
reads the already-produced Cycle 3 outputs, derives the user-facing views with
the same Excel bridge used by the runtime, and writes a self-contained set of
CSV files.  It never writes to DATA/ and never creates or changes approved
runtime state.
"""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from nmdc_profiler.excel_bridge import (  # noqa: E402
    DASHBOARD_FIELDS,
    DOCUMENT_FIELDS,
    ERROR_FIELDS,
    EVENT_FIELDS,
    FLAG_FIELDS,
    HISTORY_FIELDS,
    PENDING_FIELDS,
    REVISION_FIELDS,
    build_document_rows,
    build_event_rows,
    build_revision_rows,
)


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return [{key: (value or "") for key, value in row.items()} for row in csv.DictReader(handle)]


def _write_csv(path: Path, fields: Sequence[str], rows: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(fields), extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fields})


def _plain_review_reason(reason: str, warnings: str) -> tuple[str, str]:
    reason = (reason or "").strip()
    warnings = (warnings or "").strip()
    if "LAYOUT_DOCUMENT_HEADER_NOT_FOUND" in warnings:
        return (
            "The worksheet does not have a document-number header that the safe extractor can identify.",
            "Review the worksheet layout and map it before including it in the approved index.",
        )
    if "LAYOUT_FIRST_DATA_ROW_NOT_FOUND" in warnings:
        return (
            "The worksheet has headings, but no safe first document row could be established.",
            "Review the worksheet layout; do not guess which rows are documents.",
        )
    return (
        reason or "The worksheet needs review before it can be safely included.",
        "Review the source worksheet and decide whether a mapping or parser change is required.",
    )


def build_review_exchange(root: Path = ROOT, output_dir: Path | None = None) -> dict[str, int | str]:
    output_dir = output_dir or root.parent / "outputs" / "nmdc_doc_index_real_review" / "exchange"
    records = _read_csv(root / "outputs" / "cycle3" / "full_records.csv")
    review_queue = _read_csv(root / "outputs" / "cycle3" / "review_queue.csv")
    reconciliation = json.loads((root / "outputs" / "cycle3" / "full_reconciliation.json").read_text(encoding="utf-8"))
    inventory = _read_csv(root / "outputs" / "cycle1" / "source_inventory.csv")
    rules = _read_csv(root / "config" / "classification_rules.csv")

    documents = build_document_rows(records)
    revisions = build_revision_rows(records)
    events = build_event_rows(records)

    flags: list[dict[str, Any]] = []
    for item in review_queue:
        message, action = _plain_review_reason(item.get("Reason", ""), item.get("Warnings", ""))
        flags.append(
            {
                "Flag Level": "REVIEW",
                "Flag Code": "UNRECOGNIZED_LAYOUT",
                "Plain-English Problem": message,
                "Recommended User Action": action,
                "Project No.": item.get("Project No.", ""),
                "Document No.": "",
                "Revision": "",
                "Source File": item.get("Source File", ""),
                "Source Sheet": item.get("Worksheet", ""),
                "Source Row": "",
                "Source Cell": "",
                "User Decision": "",
                "User Comment": "",
                "Resolution Status": "OPEN",
                "Event Key": "",
            }
        )

    summary = reconciliation.get("summary", {})
    baseline_run = "CYCLE3-REAL-DATA-BASELINE"
    dashboard = {
        "Approved Status": "REAL-DATA BASELINE — REVIEW ONLY",
        "Approved Run ID": baseline_run,
        "Last Successful Update": "",
        "Approved Documents": len(documents),
        "Approved Revisions": len(revisions),
        "Approved Transactions": len(events),
        "Current Data Folder": "DATA",
        "Pending Run ID": "",
        "Pending Status": "REVIEW ONLY",
        "Pending Added": 0,
        "Pending Modified": 0,
        "Pending Removed": 0,
        "Pending Unchanged": 0,
        "Review Flags": len(flags),
        "Conflict Flags": 0,
    }
    history = [
        {
            "Date/Time": "",
            "Event": "BASELINE_IMPORTED",
            "Run ID": baseline_run,
            "Mode": "FULL_RESCAN",
            "Decision": "REVIEW_ONLY",
            "Status": "REVIEW",
            "New Sources": "",
            "Changed Sources": "",
            "Removed Sources": "",
            "Added Records": "",
            "Modified Records": "",
            "Removed Records": "",
            "Review Flags": len(flags),
            "Conflict Flags": 0,
            "Note": "Cycle 3 real-data extraction loaded for owner review. No approved master was changed.",
        }
    ]

    _write_csv(output_dir / "dashboard.csv", DASHBOARD_FIELDS, [dashboard])
    _write_csv(output_dir / "master_documents.csv", DOCUMENT_FIELDS, documents)
    _write_csv(output_dir / "revisions.csv", REVISION_FIELDS, revisions)
    _write_csv(output_dir / "events.csv", EVENT_FIELDS, events)
    _write_csv(output_dir / "pending_update.csv", PENDING_FIELDS, [])
    _write_csv(output_dir / "flags.csv", FLAG_FIELDS, flags)
    _write_csv(output_dir / "history.csv", HISTORY_FIELDS, history)
    _write_csv(output_dir / "errors.csv", ERROR_FIELDS, [])

    # Supporting exchange files keep the workbook auditable without adding
    # parser-specific logic to the workbook itself.
    inventory_fields = list(inventory[0].keys()) if inventory else []
    if inventory_fields:
        _write_csv(output_dir / "source_inventory.csv", inventory_fields, inventory)
    rule_fields = list(rules[0].keys()) if rules else []
    if rule_fields:
        _write_csv(output_dir / "classification_rules.csv", rule_fields, rules)

    (output_dir / "baseline_summary.json").write_text(
        json.dumps(
            {
                "run_id": baseline_run,
                "source": "outputs/cycle3/full_records.csv",
                "review_only": True,
                "data_read_only": True,
                "counts": {
                    "selected_workbooks": summary.get("selected_workbooks", len(inventory)),
                    "processed_workbooks": summary.get("processed_workbooks", ""),
                    "worksheets_extracted": summary.get("included_worksheets_extracted", ""),
                    "worksheets_review_required": summary.get("worksheets_review_required", len(review_queue)),
                    "documents": len(documents),
                    "revisions": len(revisions),
                    "events": len(events),
                    "hyperlinks": summary.get("hyperlinks_preserved", ""),
                    "review_flags": len(flags),
                    "conflict_flags": 0,
                },
            },
            indent=2,
            ensure_ascii=False,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
        newline="\n",
    )
    return {
        "output_dir": str(output_dir),
        "documents": len(documents),
        "revisions": len(revisions),
        "events": len(events),
        "flags": len(flags),
    }


if __name__ == "__main__":
    result = build_review_exchange()
    for key, value in result.items():
        print(f"{key}: {value}")

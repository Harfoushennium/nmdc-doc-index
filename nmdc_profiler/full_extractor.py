from __future__ import annotations

import csv
import hashlib
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Sequence, Tuple
import xml.etree.ElementTree as ET

from .extractor import OUTPUT_FIELDS, SentinelCase, _stable_key, extract_model, read_sheet_model
from .ooxml import workbook_sheet_map
from .rules import Rule


@dataclass(frozen=True)
class FullWorkItem:
    case_id: str
    source_file: str
    worksheet: str
    project_number: str
    source_family: str


REVIEW_FIELDS = [
    "Scope",
    "Source File",
    "Worksheet",
    "Project No.",
    "Source Family",
    "Reason",
    "Warnings",
    "Rows With Identity",
    "Event Records",
]


def _canonical_sheet_name(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def resolve_sheet_name(source: Path, requested: str) -> str:
    """Resolve only harmless case/whitespace differences; never guess aliases."""
    with zipfile.ZipFile(source) as z:
        names = [sheet["name"] for sheet in workbook_sheet_map(z)]
    if requested in names:
        return requested
    wanted = _canonical_sheet_name(requested)
    matches = [name for name in names if _canonical_sheet_name(name) == wanted]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise KeyError(f"Worksheet not found: {requested}")
    raise KeyError(f"Worksheet name is ambiguous after whitespace normalization: {requested}")


def _read_csv(path: Path) -> List[Dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        return [{k: (v or "") for k, v in row.items()} for row in csv.DictReader(f)]


def _case_id(source_file: str, worksheet: str) -> str:
    payload = f"{source_file}|{_canonical_sheet_name(worksheet)}".encode("utf-8")
    return "FULL-" + hashlib.sha256(payload).hexdigest()[:16]


def build_full_worklist(
    inventory_rows: Sequence[Mapping[str, str]],
    discovery_rows: Sequence[Mapping[str, str]],
) -> Tuple[List[FullWorkItem], Dict[str, Dict[str, str]], List[Dict[str, str]]]:
    selected: Dict[str, Dict[str, str]] = {
        str(row.get("relative_path", "")): dict(row)
        for row in inventory_rows
        if str(row.get("selected_excluded_status", "")).strip().upper() == "SELECTED"
    }

    work_items: List[FullWorkItem] = []
    seen: set[Tuple[str, str]] = set()
    for row in discovery_rows:
        source_file = str(row.get("workbook_path", "")).strip()
        if source_file not in selected:
            continue
        if str(row.get("proposed_action", "")).strip().upper() != "INCLUDE":
            continue
        worksheet = str(row.get("worksheet_name", "")).strip()
        if not worksheet or worksheet == "[UNREADABLE]":
            continue
        canonical = _canonical_sheet_name(worksheet)
        key = (source_file, canonical)
        if key in seen:
            continue
        seen.add(key)
        inv = selected[source_file]
        project = str(row.get("project_number", "")).strip() or str(inv.get("inferred_project_numbers", "")).strip()
        family = str(row.get("source_family", "")).strip() or str(inv.get("source_family", "")).strip()
        work_items.append(
            FullWorkItem(
                _case_id(source_file, worksheet),
                source_file,
                worksheet,
                project,
                family,
            )
        )

    work_items.sort(key=lambda item: (item.source_file.casefold(), _canonical_sheet_name(item.worksheet)))

    attempted_workbooks = {item.source_file for item in work_items}
    workbook_review: List[Dict[str, str]] = []
    for source_file, inv in sorted(selected.items(), key=lambda kv: kv[0].casefold()):
        if source_file in attempted_workbooks:
            continue
        workbook_review.append(
            {
                "Scope": "WORKBOOK",
                "Source File": source_file,
                "Worksheet": "",
                "Project No.": str(inv.get("inferred_project_numbers", "")),
                "Source Family": str(inv.get("source_family", "")),
                "Reason": "NO_INCLUDED_WORKSHEET_DISCOVERED",
                "Warnings": "",
                "Rows With Identity": "0",
                "Event Records": "0",
            }
        )
    return work_items, selected, workbook_review


def _review_row(item: FullWorkItem, reason: str, recon: Mapping[str, object] | None = None) -> Dict[str, str]:
    recon = recon or {}
    warnings = recon.get("warnings", [])
    if isinstance(warnings, (list, tuple)):
        warning_text = ";".join(str(v) for v in warnings if v)
    else:
        warning_text = str(warnings or "")
    return {
        "Scope": "WORKSHEET",
        "Source File": item.source_file,
        "Worksheet": item.worksheet,
        "Project No.": item.project_number,
        "Source Family": item.source_family,
        "Reason": reason,
        "Warnings": warning_text,
        "Rows With Identity": str(recon.get("rows_with_identity", 0)),
        "Event Records": str(recon.get("event_records", 0)),
    }


def run_full_extraction(
    root: Path,
    work_items: Sequence[FullWorkItem],
    selected_inventory: Mapping[str, Mapping[str, str]],
    initial_review: Sequence[Mapping[str, str]],
    rules: Sequence[Rule],
) -> Tuple[List[Dict[str, object]], Dict[str, object], List[Dict[str, str]]]:
    records: List[Dict[str, object]] = []
    worksheet_reconciliation: List[Dict[str, object]] = []
    review_queue: List[Dict[str, str]] = [dict(row) for row in initial_review]
    processed_workbooks: set[str] = set()

    for item in work_items:
        source = root / item.source_file
        recon: Dict[str, object]
        case_records: List[Dict[str, object]]
        if not source.exists():
            case_records = []
            recon = {
                "case_id": item.case_id,
                "source_file": item.source_file,
                "worksheet": item.worksheet,
                "status": "REVIEW_REQUIRED",
                "reason": "SOURCE_FILE_NOT_FOUND",
                "warnings": ["SOURCE_FILE_NOT_FOUND"],
                "rows_with_identity": 0,
                "event_records": 0,
                "distinct_documents": 0,
                "distinct_revisions": 0,
                "hyperlink_targets_preserved": 0,
            }
        else:
            try:
                actual_sheet = resolve_sheet_name(source, item.worksheet)
                model = read_sheet_model(source, root, actual_sheet)
                case = SentinelCase(item.case_id, item.source_file, actual_sheet, "INCLUDE")
                case_records, recon = extract_model(case, model, rules)
            except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
                case_records = []
                recon = {
                    "case_id": item.case_id,
                    "source_file": item.source_file,
                    "worksheet": item.worksheet,
                    "status": "REVIEW_REQUIRED",
                    "reason": f"EXTRACTION_FAILED:{exc.__class__.__name__}:{exc}",
                    "warnings": [f"EXTRACTION_FAILED:{exc.__class__.__name__}"],
                    "rows_with_identity": 0,
                    "event_records": 0,
                    "distinct_documents": 0,
                    "distinct_revisions": 0,
                    "hyperlink_targets_preserved": 0,
                }

        processed_workbooks.add(item.source_file)
        recon = dict(recon)
        recon["project_number"] = item.project_number
        recon["source_family"] = item.source_family
        recon["requested_worksheet"] = item.worksheet
        worksheet_reconciliation.append(recon)

        if str(recon.get("status", "")) != "INCLUDE" or int(recon.get("event_records", 0) or 0) <= 0:
            review_queue.append(_review_row(item, str(recon.get("reason", "")) or "NO_EVENT_RECORDS", recon))
            continue

        row_review_count = 0
        for rec in case_records:
            # Use the project identity already validated by source selection/discovery.
            project = item.project_number or str(rec.get("Project No.", ""))
            rec["Project No."] = project
            rec["Global_Document_Key"] = _stable_key(
                "GDOC",
                item.source_family or str(rec.get("Source Family", "")),
                project,
                rec.get("Document No.", ""),
            )
            if str(rec.get("Parsing Status", "")) != "INCLUDE":
                row_review_count += 1
            records.append(rec)
        if row_review_count:
            review_queue.append(_review_row(item, f"ROW_CLASSIFICATION_REVIEW:{row_review_count}", recon))

    records.sort(
        key=lambda r: (
            str(r.get("Project No.", "")),
            str(r.get("Source Family", "")),
            str(r.get("Source File", "")),
            str(r.get("Source Sheet", "")),
            int(r.get("Source Row", 0) or 0),
            str(r.get("Event_Key", "")),
        )
    )
    review_queue.sort(
        key=lambda r: (
            str(r.get("Source File", "")).casefold(),
            str(r.get("Worksheet", "")).casefold(),
            str(r.get("Reason", "")),
        )
    )

    source_doc_keys = {str(r.get("Source_Document_Key", "")) for r in records if r.get("Source_Document_Key")}
    global_doc_keys = {str(r.get("Global_Document_Key", "")) for r in records if r.get("Global_Document_Key")}
    revision_keys = {str(r.get("Revision_Key", "")) for r in records if r.get("Revision_Key")}
    event_keys = [str(r.get("Event_Key", "")) for r in records if r.get("Event_Key")]
    selected_paths = set(selected_inventory)
    record_paths = {str(r.get("Source File", "")) for r in records}
    leaked_paths = sorted(record_paths - selected_paths)
    unaccounted_workbooks = sorted(selected_paths - processed_workbooks - {str(r.get("Source File", "")) for r in initial_review})

    summary: Dict[str, object] = {
        "selected_workbooks": len(selected_paths),
        "processed_workbooks": len(processed_workbooks),
        "included_worksheets_attempted": len(work_items),
        "included_worksheets_extracted": sum(1 for r in worksheet_reconciliation if r.get("status") == "INCLUDE" and int(r.get("event_records", 0) or 0) > 0),
        "worksheets_review_required": sum(1 for r in worksheet_reconciliation if r.get("status") != "INCLUDE" or int(r.get("event_records", 0) or 0) <= 0),
        "workbook_review_items": sum(1 for r in review_queue if r.get("Scope") == "WORKBOOK"),
        "review_queue_items": len(review_queue),
        "event_records": len(records),
        "distinct_source_documents": len(source_doc_keys),
        "distinct_global_documents": len(global_doc_keys),
        "distinct_revisions": len(revision_keys),
        "hyperlinks_preserved": sum(1 for r in records if r.get("Document Link")),
        "row_level_review_records": sum(1 for r in records if str(r.get("Parsing Status", "")) != "INCLUDE"),
        "event_key_duplicates": len(event_keys) - len(set(event_keys)),
        "document_flag_sum": sum(int(r.get("Document_Row_Flag", 0) or 0) for r in records),
        "revision_flag_sum": sum(int(r.get("Revision_Row_Flag", 0) or 0) for r in records),
        "leaked_nonselected_sources": leaked_paths,
        "unaccounted_selected_workbooks": unaccounted_workbooks,
    }
    reconciliation = {"summary": summary, "worksheets": worksheet_reconciliation}
    return records, reconciliation, review_queue


def load_cycle3_inputs(root: Path) -> Tuple[List[FullWorkItem], Dict[str, Dict[str, str]], List[Dict[str, str]]]:
    inventory = _read_csv(root / "outputs" / "cycle1" / "source_inventory.csv")
    discovery = _read_csv(root / "outputs" / "cycle1" / "classification_discovery.csv")
    return build_full_worklist(inventory, discovery)


def write_full_outputs(
    records: Sequence[Mapping[str, object]],
    reconciliation: Mapping[str, object],
    review_queue: Sequence[Mapping[str, str]],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    with (output_dir / "full_records.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDS, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in records:
            writer.writerow({key: row.get(key, "") for key in OUTPUT_FIELDS})

    with (output_dir / "review_queue.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=REVIEW_FIELDS, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in review_queue:
            writer.writerow({key: row.get(key, "") for key in REVIEW_FIELDS})

    text = json.dumps(reconciliation, indent=2, ensure_ascii=False, sort_keys=True) + "\n"
    with (output_dir / "full_reconciliation.json").open("w", encoding="utf-8", newline="\n") as f:
        f.write(text)

    summary = dict(reconciliation.get("summary", {}))
    lines = [
        "# Cycle 3 Full Extraction Report",
        "",
        "## Simple status",
        "",
        f"- Selected workbooks: **{summary.get('selected_workbooks', 0)}**",
        f"- Workbooks processed: **{summary.get('processed_workbooks', 0)}**",
        f"- Included worksheets attempted: **{summary.get('included_worksheets_attempted', 0)}**",
        f"- Worksheets extracted: **{summary.get('included_worksheets_extracted', 0)}**",
        f"- Worksheets requiring review: **{summary.get('worksheets_review_required', 0)}**",
        f"- Review queue items: **{summary.get('review_queue_items', 0)}**",
        f"- Event records: **{summary.get('event_records', 0)}**",
        f"- Source documents: **{summary.get('distinct_source_documents', 0)}**",
        f"- Global documents: **{summary.get('distinct_global_documents', 0)}**",
        f"- Revisions: **{summary.get('distinct_revisions', 0)}**",
        f"- Hyperlinks preserved: **{summary.get('hyperlinks_preserved', 0)}**",
        f"- Row-level review records: **{summary.get('row_level_review_records', 0)}**",
        "",
        "## Safety",
        "",
        "Only Cycle-1 SELECTED workbooks are eligible. `DATA/` is read-only. No final Excel index is produced in Cycle 3.",
        "",
        "## Review queue",
        "",
    ]
    if review_queue:
        for item in review_queue:
            label = item.get("Worksheet", "") or "[workbook]"
            lines.append(f"- `{item.get('Source File', '')}` / `{label}` — {item.get('Reason', '')}")
    else:
        lines.append("- None")
    lines.append("")
    with (output_dir / "full_report.md").open("w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))

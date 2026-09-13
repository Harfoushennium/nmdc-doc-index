from __future__ import annotations

import csv
import hashlib
import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence, Tuple

Record = Dict[str, Any]

DASHBOARD_FIELDS = [
    "Approved Status",
    "Approved Run ID",
    "Last Successful Update",
    "Approved Documents",
    "Approved Revisions",
    "Approved Transactions",
    "Current Data Folder",
    "Pending Run ID",
    "Pending Status",
    "Pending Added",
    "Pending Modified",
    "Pending Removed",
    "Pending Unchanged",
    "Review Flags",
    "Conflict Flags",
]

DOCUMENT_FIELDS = [
    "Flag Level",
    "Project No.",
    "Source Family",
    "Discipline",
    "Category",
    "Subcategory",
    "Document No.",
    "Document Title",
    "Company Document No.",
    "Latest Revision",
    "Latest Event Date",
    "Latest Event Status",
    "Document Link",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Cell",
    "Global Document Key",
]

REVISION_FIELDS = [
    "Flag Level",
    "Project No.",
    "Document No.",
    "Document Title",
    "Revision",
    "Is Latest Revision",
    "Revision Key",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Cell",
]

EVENT_FIELDS = [
    "Flag Level",
    "Project No.",
    "Document No.",
    "Document Title",
    "Revision",
    "Event Type",
    "Event Date",
    "Event Reference",
    "Event Status",
    "Event Values JSON",
    "Is Latest Event",
    "Document Link",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Cell",
    "Event Key",
]

PENDING_FIELDS = [
    "Change Type",
    "Record Identity",
    "Project No.",
    "Document No.",
    "Revision",
    "Event Type",
    "Source File",
    "Plain-English Summary",
    "Review Required",
]

FLAG_FIELDS = [
    "Flag Level",
    "Flag Code",
    "Plain-English Problem",
    "Recommended User Action",
    "Project No.",
    "Document No.",
    "Revision",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Source Cell",
    "User Decision",
    "User Comment",
    "Resolution Status",
    "Event Key",
]

HISTORY_FIELDS = [
    "Date/Time",
    "Event",
    "Run ID",
    "Mode",
    "Decision",
    "Status",
    "New Sources",
    "Changed Sources",
    "Removed Sources",
    "Added Records",
    "Modified Records",
    "Removed Records",
    "Review Flags",
    "Conflict Flags",
    "Note",
]

ERROR_FIELDS = [
    "Date/Time",
    "Severity",
    "Action",
    "Plain-English Error",
    "Recommended Action",
    "Technical Detail",
    "Run ID",
    "Source File",
    "Worksheet",
    "Source Row/Cell",
]


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def _read_jsonl(path: Path) -> List[Record]:
    if not path.exists():
        return []
    rows: List[Record] = []
    with path.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                rows.append(dict(json.loads(line)))
    return rows


def _write_csv(path: Path, fields: Sequence[str], rows: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.tmp-{uuid.uuid4().hex[:8]}")
    try:
        with temp.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=list(fields), extrasaction="ignore", lineterminator="\n")
            writer.writeheader()
            for row in rows:
                writer.writerow({field: row.get(field, "") for field in fields})
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def _as_int(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _record_identity(record: Mapping[str, Any]) -> str:
    event_key = str(record.get("Event_Key", "")).strip()
    if event_key:
        return "EVENT:" + event_key
    canonical = json.dumps(dict(record), ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return "ROW:" + hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def _index_records(rows: Sequence[Mapping[str, Any]]) -> Dict[str, Record]:
    index: Dict[str, Record] = {}
    counts: Dict[str, int] = {}
    for row in rows:
        base = _record_identity(row)
        counts[base] = counts.get(base, 0) + 1
        key = base if counts[base] == 1 else f"{base}#DUP{counts[base]}"
        index[key] = dict(row)
    return index


def _approved_version_dir(state_dir: Path) -> Tuple[str, Path | None]:
    approved_dir = Path(state_dir) / "approved"
    pointer = _read_json(approved_dir / "current.json", {})
    run_id = str(pointer.get("run_id", ""))
    if run_id:
        version_dir = approved_dir / "versions" / run_id
        if version_dir.exists():
            return run_id, version_dir
    if (approved_dir / "records.jsonl").exists():
        return str(_read_json(approved_dir / "manifest.json", {}).get("run_id", "")), approved_dir
    return "", None


def load_approved_state(state_dir: Path) -> Dict[str, Any]:
    run_id, version_dir = _approved_version_dir(Path(state_dir))
    if version_dir is None:
        return {"run_id": "", "manifest": {}, "records": [], "approval": {}}
    return {
        "run_id": run_id,
        "manifest": _read_json(version_dir / "manifest.json", {}),
        "records": _read_jsonl(version_dir / "records.jsonl"),
        "approval": _read_json(version_dir / "approval.json", {}),
    }


def load_latest_stage(state_dir: Path) -> Dict[str, Any]:
    state_dir = Path(state_dir)
    latest = _read_json(state_dir / "staging" / "latest.json", {})
    run_id = str(latest.get("run_id", ""))
    if not run_id:
        return {"run_id": "", "stage_dir": None, "manifest": {}, "summary": {}, "flags": [], "changes": {}, "records": []}
    stage_dir = state_dir / "staging" / run_id
    if not stage_dir.exists():
        return {"run_id": run_id, "stage_dir": None, "manifest": {}, "summary": {}, "flags": [], "changes": {}, "records": []}
    return {
        "run_id": run_id,
        "stage_dir": stage_dir,
        "manifest": _read_json(stage_dir / "manifest.json", {}),
        "summary": _read_json(stage_dir / "summary.json", {}),
        "flags": _read_json(stage_dir / "flags.json", []),
        "changes": _read_json(stage_dir / "record_changes.json", {}),
        "records": _read_jsonl(stage_dir / "records.jsonl"),
        "decision": _read_json(stage_dir / "decision.json", {}),
    }


def _latest_by_key(records: Sequence[Mapping[str, Any]], key_name: str, flag_name: str) -> Dict[str, Record]:
    selected: Dict[str, Record] = {}
    for row in records:
        key = str(row.get(key_name, "")).strip()
        if not key:
            continue
        if _as_int(row.get(flag_name)) == 1:
            selected[key] = dict(row)
        elif key not in selected:
            selected[key] = dict(row)
    return selected


def build_document_rows(records: Sequence[Mapping[str, Any]]) -> List[Record]:
    by_doc: Dict[str, List[Record]] = {}
    for raw in records:
        row = dict(raw)
        key = str(row.get("Global_Document_Key", "") or row.get("Source_Document_Key", "")).strip()
        if not key:
            key = "DOC:" + str(row.get("Project No.", "")) + ":" + str(row.get("Document No.", ""))
        by_doc.setdefault(key, []).append(row)

    out: List[Record] = []
    for key, group in sorted(by_doc.items(), key=lambda item: item[0]):
        doc_row = next((r for r in group if _as_int(r.get("Document_Row_Flag")) == 1), group[0])
        latest_rev = next((r for r in group if _as_int(r.get("Is_Latest_Revision")) == 1), doc_row)
        latest_revision_key = str(latest_rev.get("Revision_Key", "")).strip()
        if latest_revision_key:
            latest_revision_rows = [
                r for r in group if str(r.get("Revision_Key", "")).strip() == latest_revision_key
            ]
        else:
            # Backward-compatible fallback for any pre-key canonical records.
            latest_source_document_key = str(latest_rev.get("Source_Document_Key", "")).strip()
            latest_revision_value = str(latest_rev.get("Revision", ""))
            latest_revision_rows = [
                r
                for r in group
                if str(r.get("Source_Document_Key", "")).strip() == latest_source_document_key
                and str(r.get("Revision", "")) == latest_revision_value
            ]
        latest_evt = next(
            (r for r in latest_revision_rows if _as_int(r.get("Is_Latest_Event")) == 1),
            latest_revision_rows[-1] if latest_revision_rows else latest_rev,
        )
        flag_level = "REVIEW" if str(doc_row.get("Parsing Status", "INCLUDE")) != "INCLUDE" else "OK"
        out.append(
            {
                "Flag Level": flag_level,
                "Project No.": doc_row.get("Project No.", ""),
                "Source Family": doc_row.get("Source Family", ""),
                "Discipline": doc_row.get("Discipline", ""),
                "Category": doc_row.get("Category", ""),
                "Subcategory": doc_row.get("Subcategory", ""),
                "Document No.": doc_row.get("Document No.", ""),
                "Document Title": doc_row.get("Document Title", ""),
                "Company Document No.": doc_row.get("Company Document No.", ""),
                "Latest Revision": latest_rev.get("Revision", ""),
                "Latest Event Date": latest_evt.get("Event Date", ""),
                "Latest Event Status": latest_evt.get("Event Status", ""),
                "Document Link": doc_row.get("Document Link", ""),
                "Source File": doc_row.get("Source File", ""),
                "Source Sheet": doc_row.get("Source Sheet", ""),
                "Source Row": doc_row.get("Source Row", ""),
                "Source Cell": doc_row.get("Source Cell", ""),
                "Global Document Key": key,
            }
        )
    return out


def build_revision_rows(records: Sequence[Mapping[str, Any]]) -> List[Record]:
    chosen = _latest_by_key(records, "Revision_Key", "Revision_Row_Flag")
    out: List[Record] = []
    for key, row in sorted(chosen.items(), key=lambda item: item[0]):
        out.append(
            {
                "Flag Level": "REVIEW" if str(row.get("Parsing Status", "INCLUDE")) != "INCLUDE" else "OK",
                "Project No.": row.get("Project No.", ""),
                "Document No.": row.get("Document No.", ""),
                "Document Title": row.get("Document Title", ""),
                "Revision": row.get("Revision", ""),
                "Is Latest Revision": row.get("Is_Latest_Revision", ""),
                "Revision Key": key,
                "Source File": row.get("Source File", ""),
                "Source Sheet": row.get("Source Sheet", ""),
                "Source Row": row.get("Source Row", ""),
                "Source Cell": row.get("Source Cell", ""),
            }
        )
    return out


def build_event_rows(records: Sequence[Mapping[str, Any]]) -> List[Record]:
    out: List[Record] = []
    for row in records:
        out.append(
            {
                "Flag Level": "REVIEW" if str(row.get("Parsing Status", "INCLUDE")) != "INCLUDE" else "OK",
                "Project No.": row.get("Project No.", ""),
                "Document No.": row.get("Document No.", ""),
                "Document Title": row.get("Document Title", ""),
                "Revision": row.get("Revision", ""),
                "Event Type": row.get("Event Type", ""),
                "Event Date": row.get("Event Date", ""),
                "Event Reference": row.get("Event Reference", ""),
                "Event Status": row.get("Event Status", ""),
                "Event Values JSON": row.get("Event Values JSON", ""),
                "Is Latest Event": row.get("Is_Latest_Event", ""),
                "Document Link": row.get("Document Link", ""),
                "Source File": row.get("Source File", ""),
                "Source Sheet": row.get("Source Sheet", ""),
                "Source Row": row.get("Source Row", ""),
                "Source Cell": row.get("Source Cell", ""),
                "Event Key": row.get("Event_Key", ""),
            }
        )
    return out


def build_pending_rows(approved_records: Sequence[Mapping[str, Any]], stage: Mapping[str, Any]) -> List[Record]:
    changes = dict(stage.get("changes", {}))
    staged_records = list(stage.get("records", []))
    old_index = _index_records(approved_records)
    new_index = _index_records(staged_records)
    rows: List[Record] = []
    for change_type, key_name in (
        ("ADDED", "added"),
        ("MODIFIED", "modified"),
        ("REMOVED", "removed"),
        ("UNCHANGED", "unchanged"),
    ):
        for identity in changes.get(key_name, []) or []:
            row = new_index.get(identity) if change_type != "REMOVED" else old_index.get(identity)
            row = row or {}
            summary = {
                "ADDED": "New record will be added if this update is approved.",
                "MODIFIED": "Existing approved record changed in the staged update.",
                "REMOVED": "Approved record will be removed if this update is approved.",
                "UNCHANGED": "Approved record is unchanged and will be carried forward.",
            }[change_type]
            rows.append(
                {
                    "Change Type": change_type,
                    "Record Identity": identity,
                    "Project No.": row.get("Project No.", ""),
                    "Document No.": row.get("Document No.", ""),
                    "Revision": row.get("Revision", ""),
                    "Event Type": row.get("Event Type", ""),
                    "Source File": row.get("Source File", ""),
                    "Plain-English Summary": summary,
                    "Review Required": "NO" if change_type == "UNCHANGED" else "YES",
                }
            )
    return rows


def build_flag_rows(stage: Mapping[str, Any]) -> List[Record]:
    rows: List[Record] = []
    for flag in stage.get("flags", []) or []:
        rows.append(
            {
                "Flag Level": flag.get("level", ""),
                "Flag Code": flag.get("code", ""),
                "Plain-English Problem": flag.get("message", ""),
                "Recommended User Action": flag.get("recommended_action", ""),
                "Project No.": flag.get("project_no", ""),
                "Document No.": flag.get("document_no", ""),
                "Revision": flag.get("revision", ""),
                "Source File": flag.get("source", ""),
                "Source Sheet": flag.get("source_sheet", ""),
                "Source Row": flag.get("source_row", ""),
                "Source Cell": flag.get("source_cell", ""),
                "User Decision": flag.get("user_decision", ""),
                "User Comment": flag.get("user_comment", ""),
                "Resolution Status": flag.get("resolution_status", "OPEN"),
                "Event Key": flag.get("event_key", ""),
            }
        )
    return rows


def build_history_rows(state_dir: Path) -> List[Record]:
    rows: List[Record] = []
    for item in _read_jsonl(Path(state_dir) / "logs" / "history.jsonl"):
        source_counts = item.get("source_counts", {}) or {}
        record_counts = item.get("record_counts", {}) or {}
        rows.append(
            {
                "Date/Time": item.get("at", item.get("approved_at", item.get("held_at", item.get("rejected_at", "")))),
                "Event": item.get("event", ""),
                "Run ID": item.get("run_id", ""),
                "Mode": item.get("mode", ""),
                "Decision": item.get("decision", ""),
                "Status": item.get("status", ""),
                "New Sources": source_counts.get("NEW", ""),
                "Changed Sources": source_counts.get("CHANGED", ""),
                "Removed Sources": source_counts.get("REMOVED", ""),
                "Added Records": record_counts.get("added", ""),
                "Modified Records": record_counts.get("modified", ""),
                "Removed Records": record_counts.get("removed", ""),
                "Review Flags": item.get("review_flags", ""),
                "Conflict Flags": item.get("blocking_flags", ""),
                "Note": item.get("note", ""),
            }
        )
    return rows


def build_error_rows(stage: Mapping[str, Any]) -> List[Record]:
    run_id = str(stage.get("run_id", ""))
    rows: List[Record] = []
    for flag in stage.get("flags", []) or []:
        if str(flag.get("level", "")) != "CONFLICT":
            continue
        rows.append(
            {
                "Date/Time": "",
                "Severity": "CONFLICT",
                "Action": "Review staged update",
                "Plain-English Error": flag.get("message", ""),
                "Recommended Action": flag.get("recommended_action", ""),
                "Technical Detail": flag.get("code", ""),
                "Run ID": run_id,
                "Source File": flag.get("source", ""),
                "Worksheet": flag.get("source_sheet", ""),
                "Source Row/Cell": flag.get("source_cell", "") or flag.get("source_row", ""),
            }
        )
    return rows


def build_dashboard_row(approved: Mapping[str, Any], stage: Mapping[str, Any]) -> Record:
    approved_records = list(approved.get("records", []))
    documents = build_document_rows(approved_records)
    revisions = build_revision_rows(approved_records)
    manifest = approved.get("manifest", {}) or {}
    stage_manifest = stage.get("manifest", {}) or {}
    summary = stage.get("summary", {}) or {}
    decision = stage.get("decision", {}) or {}
    counts = summary.get("record_counts", {}) or {}
    flags = stage.get("flags", []) or []
    return {
        "Approved Status": "APPROVED" if approved.get("run_id") else "NO APPROVED INDEX",
        "Approved Run ID": approved.get("run_id", ""),
        "Last Successful Update": (approved.get("approval", {}) or {}).get("approved_at", ""),
        "Approved Documents": len(documents),
        "Approved Revisions": len(revisions),
        "Approved Transactions": len(approved_records),
        "Current Data Folder": stage_manifest.get("data_dir", manifest.get("data_dir", "")),
        "Pending Run ID": stage.get("run_id", ""),
        "Pending Status": decision.get("decision", summary.get("status", "NONE")) if stage.get("run_id") else "NONE",
        "Pending Added": counts.get("added", 0),
        "Pending Modified": counts.get("modified", 0),
        "Pending Removed": counts.get("removed", 0),
        "Pending Unchanged": counts.get("unchanged", 0),
        "Review Flags": sum(1 for f in flags if f.get("level") == "REVIEW"),
        "Conflict Flags": sum(1 for f in flags if f.get("level") == "CONFLICT"),
    }


def export_excel_exchange(state_dir: Path, output_dir: Path) -> Dict[str, Any]:
    state_dir = Path(state_dir)
    output_dir = Path(output_dir)
    approved = load_approved_state(state_dir)
    stage = load_latest_stage(state_dir)
    approved_records = list(approved.get("records", []))

    dashboard = build_dashboard_row(approved, stage)
    documents = build_document_rows(approved_records)
    revisions = build_revision_rows(approved_records)
    events = build_event_rows(approved_records)
    pending = build_pending_rows(approved_records, stage)
    flags = build_flag_rows(stage)
    history = build_history_rows(state_dir)
    errors = build_error_rows(stage)

    _write_csv(output_dir / "dashboard.csv", DASHBOARD_FIELDS, [dashboard])
    _write_csv(output_dir / "master_documents.csv", DOCUMENT_FIELDS, documents)
    _write_csv(output_dir / "revisions.csv", REVISION_FIELDS, revisions)
    _write_csv(output_dir / "events.csv", EVENT_FIELDS, events)
    _write_csv(output_dir / "pending_update.csv", PENDING_FIELDS, pending)
    _write_csv(output_dir / "flags.csv", FLAG_FIELDS, flags)
    _write_csv(output_dir / "history.csv", HISTORY_FIELDS, history)
    _write_csv(output_dir / "errors.csv", ERROR_FIELDS, errors)

    return {
        "approved_run_id": approved.get("run_id", ""),
        "pending_run_id": stage.get("run_id", ""),
        "documents": len(documents),
        "revisions": len(revisions),
        "events": len(events),
        "pending_rows": len(pending),
        "flags": len(flags),
        "errors": len(errors),
        "output_dir": str(output_dir),
    }


def create_support_request(
    state_dir: Path,
    *,
    message: str,
    source_file: str = "",
    worksheet: str = "",
    source_row: str = "",
    source_cell: str = "",
    project_no: str = "",
    document_no: str = "",
    revision: str = "",
    event_identity: str = "",
    flag_code: str = "",
    current_field: str = "",
    current_value: str = "",
    expected_value: str = "",
    user_name: str = "",
    parser_version: str = "",
    configuration_version: str = "",
) -> Path:
    state_dir = Path(state_dir)
    now = datetime.now(timezone.utc).replace(microsecond=0)
    request_id = now.strftime("SUPPORT-%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:8]
    support_dir = state_dir / "support"
    support_dir.mkdir(parents=True, exist_ok=True)
    path = support_dir / f"{request_id}.json"
    approved = load_approved_state(state_dir)
    stage = load_latest_stage(state_dir)
    active_manifest = stage.get("manifest", {}) or approved.get("manifest", {}) or {}
    effective_parser_version = parser_version or str(active_manifest.get("parser_version", ""))
    effective_configuration_version = configuration_version or str(
        active_manifest.get("config_version", active_manifest.get("config_fingerprint", ""))
    )
    payload = {
        "request_id": request_id,
        "created_at": now.isoformat(),
        "user": user_name,
        "message": message,
        "source_file": source_file,
        "worksheet": worksheet,
        "source_row": source_row,
        "source_cell": source_cell,
        "project_no": project_no,
        "document_no": document_no,
        "revision": revision,
        "event_identity": event_identity,
        "flag_code": flag_code,
        "current_field": current_field,
        "current_value": current_value,
        "expected_value": expected_value,
        "parser_version": effective_parser_version,
        "configuration_version": effective_configuration_version,
        "configuration_fingerprint": str(active_manifest.get("config_fingerprint", "")),
        "approved_run_id": approved.get("run_id", ""),
        "pending_run_id": stage.get("run_id", ""),
        "run_id": stage.get("run_id", "") or approved.get("run_id", ""),
    }
    with path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(payload, indent=2, ensure_ascii=False, sort_keys=True) + "\n")
    return path

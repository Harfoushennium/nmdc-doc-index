from __future__ import annotations

import csv
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Tuple

ALLOWED_DECISIONS = {
    "ACKNOWLEDGED",
    "NO ACTION REQUIRED",
    "NEEDS SOURCE CORRECTION",
    "NEEDS PARSER/MAPPING FIX",
    "HOLD FOR REVIEW",
}
ALLOWED_STATUSES = {"OPEN", "ACKNOWLEDGED", "RESOLVED", "DEFERRED"}


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def _append_history(state_dir: Path, payload: Mapping[str, Any]) -> None:
    path = Path(state_dir) / "logs" / "history.jsonl"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(dict(payload), ensure_ascii=False, sort_keys=True, default=str) + "\n")


def _key_from_flag(flag: Mapping[str, Any]) -> Tuple[str, ...]:
    return (
        str(flag.get("code", "")).strip(),
        str(flag.get("source", "")).strip(),
        str(flag.get("source_sheet", "")).strip(),
        str(flag.get("event_key", "")).strip(),
        str(flag.get("project_no", "")).strip(),
        str(flag.get("document_no", "")).strip(),
        str(flag.get("revision", "")).strip(),
    )


def _key_from_row(row: Mapping[str, str]) -> Tuple[str, ...]:
    return (
        str(row.get("Flag Code", "")).strip(),
        str(row.get("Source File", "")).strip(),
        str(row.get("Source Sheet", "")).strip(),
        str(row.get("Event Key", "")).strip(),
        str(row.get("Project No.", "")).strip(),
        str(row.get("Document No.", "")).strip(),
        str(row.get("Revision", "")).strip(),
    )


def apply_review_decisions(state_dir: Path, decisions_file: Path) -> Dict[str, Any]:
    state_dir = Path(state_dir)
    decisions_file = Path(decisions_file)
    if not decisions_file.exists():
        raise FileNotFoundError(f"Review decisions file not found: {decisions_file}")

    latest = _read_json(state_dir / "staging" / "latest.json", {})
    run_id = str(latest.get("run_id", ""))
    if not run_id:
        raise ValueError("There is no staged update available for review decisions.")
    stage_dir = state_dir / "staging" / run_id
    flags_path = stage_dir / "flags.json"
    flags = _read_json(flags_path, [])
    if not isinstance(flags, list):
        raise ValueError("The staged review flag file is invalid.")

    with decisions_file.open(newline="", encoding="utf-8-sig") as handle:
        rows = [{k: (v or "") for k, v in row.items()} for row in csv.DictReader(handle)]

    index: Dict[Tuple[str, ...], List[Dict[str, Any]]] = {}
    for raw in flags:
        if not isinstance(raw, dict):
            continue
        index.setdefault(_key_from_flag(raw), []).append(raw)

    updated = 0
    unmatched: List[Tuple[str, ...]] = []
    for row in rows:
        decision = str(row.get("User Decision", "")).strip().upper()
        comment = str(row.get("User Comment", "")).strip()
        status = str(row.get("Resolution Status", "")).strip().upper()
        if not decision and not comment and not status:
            continue
        if decision and decision not in ALLOWED_DECISIONS:
            raise ValueError(f"Unsupported review decision: {decision}")
        if status and status not in ALLOWED_STATUSES:
            raise ValueError(f"Unsupported resolution status: {status}")

        key = _key_from_row(row)
        matches = index.get(key, [])
        if not matches:
            unmatched.append(key)
            continue
        for flag in matches:
            if decision:
                flag["user_decision"] = decision
            if comment:
                flag["user_comment"] = comment
            if status:
                flag["resolution_status"] = status
            flag["reviewed_at"] = _utc_now()
            updated += 1

    _write_json(flags_path, flags)
    result = {
        "run_id": run_id,
        "updated_flags": updated,
        "unmatched_decisions": len(unmatched),
        "saved_at": _utc_now(),
    }
    _append_history(state_dir, {"event": "REVIEW_DECISIONS_SAVED", **result})
    return result

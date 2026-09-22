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


def _clean_md(value: Any) -> str:
    return str(value or "").replace("\r", " ").replace("\n", " ").replace("|", "\\|").strip()


def _write_parser_mapping_fix_request(
    state_dir: Path,
    run_id: str,
    flags: Iterable[Mapping[str, Any]],
    request_dir: Path | None = None,
) -> Path:
    support_dir = Path(request_dir) if request_dir is not None else Path(state_dir) / "support"
    support_dir.mkdir(parents=True, exist_ok=True)
    rows = [dict(flag) for flag in flags]
    payload = {
        "request_type": "PARSER_MAPPING_FIX",
        "run_id": run_id,
        "created_at": _utc_now(),
        "instructions": (
            "Upload this request together with the affected source workbook(s) to the NMDC Document Index "
            "maintainer / ChatGPT. The installed workbook records and retries extraction, but it does not "
            "rewrite its own packaged parser code automatically."
        ),
        "flags": rows,
    }
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    json_path = support_dir / f"PARSER_FIX_REQUEST_{stamp}.json"
    latest_json_path = support_dir / "PARSER_FIX_REQUEST_LATEST.json"
    _write_json(json_path, payload)
    _write_json(latest_json_path, payload)

    lines = [
        "# NMDC Document Index — Parser / Mapping Fix Request",
        "",
        f"- Staged run: `{run_id}`",
        f"- Created: `{payload['created_at']}`",
        f"- Affected review flags: **{len(rows)}**",
        "",
        "## What the owner is asking",
        "",
        "The owner selected **NEEDS PARSER/MAPPING FIX**. The current installed workbook will preserve the",
        "approved index and source DATA. It will not pretend to rewrite its own packaged parser automatically.",
        "Use this request together with the affected source workbook(s) to implement and test the correction.",
        "",
        "After a corrected parser/configuration is installed, use **Retry After Fix** in Review Flags (or Full Rescan)",
        "to re-extract the source. The flag should disappear only when the new extraction succeeds.",
        "",
        "## Affected items",
        "",
        "| Source File | Source Sheet | Project No. | Document No. | Revision | Flag Code | Owner Comment |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for flag in rows:
        lines.append(
            "| " + " | ".join(
                [
                    _clean_md(flag.get("source", "")),
                    _clean_md(flag.get("source_sheet", "")),
                    _clean_md(flag.get("project_no", "")),
                    _clean_md(flag.get("document_no", "")),
                    _clean_md(flag.get("revision", "")),
                    _clean_md(flag.get("code", "")),
                    _clean_md(flag.get("user_comment", "")),
                ]
            ) + " |"
        )
    lines.extend(
        [
            "",
            "## Maintainer acceptance",
            "",
            "- Reproduce against the exact source workbook/sheet.",
            "- Fix parser/configuration only; never modify source DATA.",
            "- Add regression coverage for the reproduced layout/mapping.",
            "- Re-run extraction and confirm the Review Flag no longer appears.",
        ]
    )
    md_text = "\n".join(lines) + "\n"
    md_path = support_dir / f"PARSER_FIX_REQUEST_{stamp}.md"
    latest_md_path = support_dir / "PARSER_FIX_REQUEST_LATEST.md"
    md_path.write_text(md_text, encoding="utf-8")
    latest_md_path.write_text(md_text, encoding="utf-8")
    return latest_md_path


def apply_review_decisions(
    state_dir: Path,
    decisions_file: Path,
    request_dir: Path | None = None,
) -> Dict[str, Any]:
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
    fix_request_flags: List[Dict[str, Any]] = []
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
            if decision == "NEEDS PARSER/MAPPING FIX":
                fix_request_flags.append(flag)
            updated += 1

    _write_json(flags_path, flags)

    unique_fix_flags: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    for flag in fix_request_flags:
        unique_fix_flags[_key_from_flag(flag)] = flag
    fix_request_file = ""
    if unique_fix_flags:
        fix_request_file = str(
            _write_parser_mapping_fix_request(
                state_dir,
                run_id,
                unique_fix_flags.values(),
                request_dir=request_dir,
            )
        )

    result = {
        "run_id": run_id,
        "updated_flags": updated,
        "unmatched_decisions": len(unmatched),
        "parser_mapping_fix_requests": len(unique_fix_flags),
        "fix_request_file": fix_request_file,
        "saved_at": _utc_now(),
    }
    _append_history(state_dir, {"event": "REVIEW_DECISIONS_SAVED", **result})
    return result

from __future__ import annotations

import hashlib
import json
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Sequence, Tuple

ENGINE_SCHEMA_VERSION = "1"
DEFAULT_PARSER_VERSION = "cycle3-extractor-v1"
NON_OVERRIDABLE_CONFLICT_CODES = {
    "SOURCE_HASH_ERROR",
    "PARSER_ERROR",
    "DUPLICATE_RECORD_KEY",
}

Record = Dict[str, Any]
Processor = Callable[[Path, str], Sequence[Mapping[str, Any]]]


@dataclass(frozen=True)
class SourceDecision:
    relative_path: str
    change_type: str
    action: str
    reason: str


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _write_json(path: Path, payload: Mapping[str, Any] | Sequence[Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, indent=2, ensure_ascii=False, sort_keys=True) + "\n"
    with path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def _write_jsonl(path: Path, rows: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as f:
        for row in rows:
            f.write(json.dumps(dict(row), ensure_ascii=False, sort_keys=True, default=str) + "\n")


def _read_jsonl(path: Path) -> List[Record]:
    if not path.exists():
        return []
    out: List[Record] = []
    with path.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                out.append(dict(json.loads(line)))
    return out


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def fingerprint_paths(paths: Sequence[Path]) -> str:
    h = hashlib.sha256()
    for path in sorted((Path(p) for p in paths), key=lambda p: str(p).casefold()):
        h.update(str(path).encode("utf-8", "surrogatepass"))
        h.update(b"\0")
        if path.exists() and path.is_file():
            h.update(sha256_file(path).encode("ascii"))
        else:
            h.update(b"[MISSING]")
        h.update(b"\0")
    return h.hexdigest()


def _flag(level: str, code: str, message: str, source: str = "", action: str = "") -> Dict[str, str]:
    return {
        "level": level,
        "code": code,
        "message": message,
        "source": source,
        "recommended_action": action,
    }


def scan_sources(data_dir: Path) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    """Scan an arbitrary user-selected folder without modifying it."""
    data_dir = Path(data_dir)
    entries: List[Dict[str, Any]] = []
    flags: List[Dict[str, str]] = []
    if not data_dir.exists() or not data_dir.is_dir():
        raise FileNotFoundError(f"Data folder not found: {data_dir}")

    files = sorted(
        (p for p in data_dir.rglob("*") if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}),
        key=lambda p: p.relative_to(data_dir).as_posix().casefold(),
    )
    for path in files:
        rel = path.relative_to(data_dir).as_posix()
        stat = path.stat()
        entry: Dict[str, Any] = {
            "relative_path": rel,
            "size": int(stat.st_size),
            "modified_ns": int(stat.st_mtime_ns),
            "sha256": "",
            "scan_error": "",
        }
        try:
            entry["sha256"] = sha256_file(path)
        except OSError as exc:
            entry["scan_error"] = f"{exc.__class__.__name__}: {exc}"
            flags.append(
                _flag(
                    "CONFLICT",
                    "SOURCE_HASH_ERROR",
                    f"Could not read source workbook '{rel}' to calculate its content hash.",
                    rel,
                    "Check that the file is accessible and not locked, then run the update again.",
                )
            )
        entries.append(entry)
    return entries, flags


def build_manifest(
    entries: Sequence[Mapping[str, Any]],
    *,
    parser_version: str,
    config_fingerprint: str,
    run_id: str,
    data_dir: Path,
    selection_statuses: Mapping[str, str] | None = None,
) -> Dict[str, Any]:
    statuses = selection_statuses or {}
    files: List[Dict[str, Any]] = []
    for row in entries:
        item = dict(row)
        rel = str(item.get("relative_path", ""))
        item["selection_status"] = str(statuses.get(rel, "SELECTED")).strip().upper() or "SELECTED"
        item["last_processed_run"] = ""
        files.append(item)
    files.sort(key=lambda r: str(r.get("relative_path", "")).casefold())
    return {
        "engine_schema_version": ENGINE_SCHEMA_VERSION,
        "parser_version": parser_version,
        "config_fingerprint": config_fingerprint,
        "run_id": run_id,
        "data_dir": str(Path(data_dir).resolve()),
        "files": files,
    }


def _manifest_map(manifest: Mapping[str, Any]) -> Dict[str, Dict[str, Any]]:
    return {
        str(row.get("relative_path", "")): dict(row)
        for row in manifest.get("files", [])
        if str(row.get("relative_path", ""))
    }


def compare_manifests(
    current: Mapping[str, Any],
    approved: Mapping[str, Any] | None,
    *,
    full_rescan: bool = False,
) -> List[SourceDecision]:
    approved = approved or {}
    cur = _manifest_map(current)
    old = _manifest_map(approved)
    parser_changed = bool(approved) and approved.get("parser_version") != current.get("parser_version")
    config_changed = bool(approved) and approved.get("config_fingerprint") != current.get("config_fingerprint")

    decisions: List[SourceDecision] = []
    for rel in sorted(cur, key=str.casefold):
        now = cur[rel]
        before = old.get(rel)
        if now.get("scan_error"):
            decisions.append(SourceDecision(rel, "ERROR", "KEEP_APPROVED" if before else "BLOCK", "SOURCE_HASH_ERROR"))
            continue
        if before is None:
            decisions.append(SourceDecision(rel, "NEW", "REPROCESS", "SOURCE_NEW"))
            continue
        content_changed = now.get("sha256") != before.get("sha256")
        if full_rescan:
            decisions.append(SourceDecision(rel, "CHANGED" if content_changed else "UNCHANGED", "REPROCESS", "FULL_RESCAN"))
        elif parser_changed:
            decisions.append(SourceDecision(rel, "CHANGED" if content_changed else "UNCHANGED", "REPROCESS", "PARSER_VERSION_CHANGED"))
        elif config_changed:
            decisions.append(SourceDecision(rel, "CHANGED" if content_changed else "UNCHANGED", "REPROCESS", "CONFIG_CHANGED"))
        elif content_changed:
            decisions.append(SourceDecision(rel, "CHANGED", "REPROCESS", "SOURCE_CHANGED"))
        else:
            decisions.append(SourceDecision(rel, "UNCHANGED", "REUSE", "SOURCE_UNCHANGED"))

    for rel in sorted(set(old) - set(cur), key=str.casefold):
        decisions.append(SourceDecision(rel, "REMOVED", "DROP", "SOURCE_REMOVED"))
    return decisions


def _cache_key(relative_path: str, file_hash: str, parser_version: str, config_fingerprint: str) -> str:
    raw = "\0".join((relative_path, file_hash, parser_version, config_fingerprint)).encode("utf-8", "surrogatepass")
    return hashlib.sha256(raw).hexdigest()[:32]


def _canonical_record(record: Mapping[str, Any]) -> str:
    return json.dumps(dict(record), ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _record_identity(record: Mapping[str, Any]) -> str:
    event_key = str(record.get("Event_Key", "")).strip()
    if event_key:
        return "EVENT:" + event_key
    return "ROW:" + hashlib.sha256(_canonical_record(record).encode("utf-8")).hexdigest()


def _index_records(rows: Sequence[Mapping[str, Any]], side: str) -> Tuple[Dict[str, Dict[str, Any]], List[Dict[str, str]]]:
    index: Dict[str, Dict[str, Any]] = {}
    counts: Dict[str, int] = {}
    flags: List[Dict[str, str]] = []
    for row in rows:
        base = _record_identity(row)
        counts[base] = counts.get(base, 0) + 1
        key = base if counts[base] == 1 else f"{base}#DUP{counts[base]}"
        if counts[base] == 2:
            flags.append(
                _flag(
                    "CONFLICT",
                    "DUPLICATE_RECORD_KEY",
                    f"Duplicate canonical record key found in {side} dataset: {base}.",
                    str(row.get("Source File", "")),
                    "Review duplicate document/revision/event rows before approval.",
                )
            )
        index[key] = dict(row)
    return index, flags


def compare_records(
    approved_rows: Sequence[Mapping[str, Any]],
    staged_rows: Sequence[Mapping[str, Any]],
) -> Tuple[Dict[str, Any], List[Dict[str, str]]]:
    old, flags_old = _index_records(approved_rows, "approved")
    new, flags_new = _index_records(staged_rows, "staged")
    added: List[str] = []
    removed: List[str] = []
    modified: List[str] = []
    unchanged: List[str] = []

    for key in sorted(set(old) | set(new)):
        if key not in old:
            added.append(key)
        elif key not in new:
            removed.append(key)
        elif _canonical_record(old[key]) != _canonical_record(new[key]):
            modified.append(key)
        else:
            unchanged.append(key)
    return {
        "added": added,
        "modified": modified,
        "removed": removed,
        "unchanged": unchanged,
        "counts": {
            "added": len(added),
            "modified": len(modified),
            "removed": len(removed),
            "unchanged": len(unchanged),
        },
    }, flags_old + flags_new


def _append_log(state_dir: Path, payload: Mapping[str, Any]) -> None:
    path = Path(state_dir) / "logs" / "history.jsonl"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(dict(payload), ensure_ascii=False, sort_keys=True, default=str) + "\n")


def _set_manifest_processed_run(manifest: Dict[str, Any], rel: str, processed_run: str) -> None:
    for row in manifest.get("files", []):
        if str(row.get("relative_path", "")) == rel:
            row["last_processed_run"] = processed_run
            return


def stage_update(
    *,
    data_dir: Path,
    state_dir: Path,
    processor: Processor,
    config_paths: Sequence[Path] = (),
    parser_version: str = DEFAULT_PARSER_VERSION,
    full_rescan: bool = False,
    selection_statuses: Mapping[str, str] | None = None,
) -> Dict[str, Any]:
    """Create a proposed update without changing the approved dataset."""
    data_dir = Path(data_dir)
    state_dir = Path(state_dir)
    approved_dir = state_dir / "approved"
    approved_manifest = _read_json(approved_dir / "manifest.json", {})
    approved_cache_index = _read_json(approved_dir / "cache_index.json", {})
    approved_rows = _read_jsonl(approved_dir / "records.jsonl")

    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:8]
    config_fingerprint = fingerprint_paths([Path(p) for p in config_paths])
    scanned, flags = scan_sources(data_dir)
    current_manifest = build_manifest(
        scanned,
        parser_version=parser_version,
        config_fingerprint=config_fingerprint,
        run_id=run_id,
        data_dir=data_dir,
        selection_statuses=selection_statuses,
    )
    decisions = compare_manifests(current_manifest, approved_manifest, full_rescan=full_rescan)
    current_map = _manifest_map(current_manifest)
    approved_map = _manifest_map(approved_manifest)
    cache_index: Dict[str, str] = {}
    staged_rows: List[Record] = []
    effective_decisions: List[SourceDecision] = []

    for original_decision in decisions:
        decision = original_decision
        rel = decision.relative_path
        now = current_map.get(rel, {})
        old = approved_map.get(rel, {})
        now_status = str(now.get("selection_status", "SELECTED")).upper()
        old_status = str(old.get("selection_status", "SELECTED")).upper() if old else ""

        if rel in current_map and now_status != "SELECTED":
            if now_status == "REVIEW_REQUIRED":
                flags.append(
                    _flag(
                        "REVIEW",
                        "SOURCE_SELECTION_REVIEW_REQUIRED",
                        f"Source workbook '{rel}' requires source-selection review and is not included in the staged master data.",
                        rel,
                        "Resolve source selection before expecting this workbook to appear in the approved index.",
                    )
                )
            if old_status == "SELECTED":
                flags.append(
                    _flag(
                        "REVIEW",
                        "SOURCE_SELECTION_CHANGED",
                        f"Source workbook '{rel}' is no longer selected; its previously approved records are staged for removal.",
                        rel,
                        "Review the source-selection decision before approval.",
                    )
                )
            decision = SourceDecision(rel, decision.change_type, "DROP", "SOURCE_NOT_SELECTED")

        if decision.reason == "SOURCE_NEW" and now_status == "SELECTED":
            flags.append(_flag("REVIEW", "SOURCE_NEW", f"New source workbook detected: {rel}.", rel, "Review the staged extracted records before approval."))
        elif decision.reason == "SOURCE_CHANGED" and now_status == "SELECTED":
            flags.append(_flag("REVIEW", "SOURCE_CHANGED", f"Source workbook content changed: {rel}.", rel, "Review staged differences before approval."))
        elif decision.reason == "SOURCE_REMOVED" and old_status == "SELECTED":
            flags.append(_flag("CONFLICT", "SOURCE_REMOVED", f"Previously approved source workbook is no longer present: {rel}.", rel, "Confirm that removal is intentional before approval."))
        elif decision.reason == "PARSER_VERSION_CHANGED" and now_status == "SELECTED":
            flags.append(_flag("REVIEW", "PARSER_VERSION_CHANGED", f"Parser version changed; '{rel}' will be reprocessed.", rel, "Review staged differences before approval."))
        elif decision.reason == "CONFIG_CHANGED" and now_status == "SELECTED":
            flags.append(_flag("REVIEW", "CONFIG_CHANGED", f"Configuration changed; '{rel}' will be reprocessed.", rel, "Review staged differences before approval."))

        if decision.action == "DROP":
            effective_decisions.append(decision)
            _set_manifest_processed_run(current_manifest, rel, str(old.get("last_processed_run", "")))
            continue

        if decision.action == "REUSE":
            cache_file = str(approved_cache_index.get(rel, ""))
            cache_path = state_dir / cache_file if cache_file else Path()
            if not cache_file or not cache_path.exists():
                decision = SourceDecision(rel, decision.change_type, "REPROCESS", "CACHE_MISSING")
                flags.append(_flag("REVIEW", "CACHE_MISSING", f"Cached extraction is missing for unchanged source '{rel}'; it will be safely reprocessed.", rel, "Review staged differences before approval."))
            else:
                rows = _read_jsonl(cache_path)
                staged_rows.extend(rows)
                cache_index[rel] = cache_file
                _set_manifest_processed_run(current_manifest, rel, str(old.get("last_processed_run", approved_manifest.get("run_id", ""))))
                effective_decisions.append(decision)
                continue

        if decision.action == "KEEP_APPROVED":
            cache_file = str(approved_cache_index.get(rel, ""))
            cache_path = state_dir / cache_file if cache_file else Path()
            if cache_file and cache_path.exists():
                staged_rows.extend(_read_jsonl(cache_path))
                cache_index[rel] = cache_file
            _set_manifest_processed_run(current_manifest, rel, str(old.get("last_processed_run", approved_manifest.get("run_id", ""))))
            effective_decisions.append(decision)
            continue

        if decision.action == "BLOCK":
            effective_decisions.append(decision)
            continue

        if decision.action == "REPROCESS":
            source_path = data_dir / rel
            try:
                rows = [dict(r) for r in processor(source_path, rel)]
                for row in rows:
                    row.setdefault("Source File", rel)
                key = _cache_key(
                    rel,
                    str(now.get("sha256", old.get("sha256", ""))),
                    parser_version,
                    config_fingerprint,
                )
                cache_rel = f"cache/{key}.jsonl"
                _write_jsonl(state_dir / cache_rel, rows)
                cache_index[rel] = cache_rel
                staged_rows.extend(rows)
                _set_manifest_processed_run(current_manifest, rel, run_id)
            except Exception as exc:
                flags.append(
                    _flag(
                        "CONFLICT",
                        "PARSER_ERROR",
                        f"Could not process source workbook '{rel}': {exc.__class__.__name__}: {exc}",
                        rel,
                        "Keep the approved data unchanged for this source and report the error for correction.",
                    )
                )
                old_cache = str(approved_cache_index.get(rel, ""))
                old_cache_path = state_dir / old_cache if old_cache else Path()
                if old_cache and old_cache_path.exists():
                    staged_rows.extend(_read_jsonl(old_cache_path))
                    cache_index[rel] = old_cache
                _set_manifest_processed_run(current_manifest, rel, str(old.get("last_processed_run", approved_manifest.get("run_id", ""))))
            effective_decisions.append(decision)

    staged_rows.sort(
        key=lambda r: (
            str(r.get("Project No.", "")),
            str(r.get("Document No.", "")),
            str(r.get("Revision", "")),
            str(r.get("Event_Key", "")),
            str(r.get("Source File", "")),
            str(r.get("Source Row", "")),
        )
    )
    record_changes, record_flags = compare_records(approved_rows, staged_rows)
    flags.extend(record_flags)
    blocking = sum(1 for f in flags if f.get("level") == "CONFLICT")

    stage_dir = state_dir / "staging" / run_id
    _write_json(stage_dir / "manifest.json", current_manifest)
    _write_json(stage_dir / "cache_index.json", cache_index)
    _write_jsonl(stage_dir / "records.jsonl", staged_rows)
    _write_json(
        stage_dir / "source_decisions.json",
        [
            {
                "relative_path": d.relative_path,
                "change_type": d.change_type,
                "action": d.action,
                "reason": d.reason,
            }
            for d in effective_decisions
        ],
    )
    _write_json(stage_dir / "record_changes.json", record_changes)
    _write_json(stage_dir / "flags.json", flags)
    summary = {
        "run_id": run_id,
        "mode": "FULL_RESCAN" if full_rescan else "INCREMENTAL",
        "status": "REVIEW_REQUIRED" if blocking else "STAGED",
        "blocking_flags": blocking,
        "review_flags": sum(1 for f in flags if f.get("level") == "REVIEW"),
        "source_counts": {
            key: sum(1 for d in decisions if d.change_type == key)
            for key in ("NEW", "CHANGED", "UNCHANGED", "REMOVED", "ERROR")
        },
        "record_counts": record_changes["counts"],
        "staged_records": len(staged_rows),
        "approved_records_before": len(approved_rows),
    }
    _write_json(stage_dir / "summary.json", summary)
    _write_json(state_dir / "staging" / "latest.json", {"run_id": run_id})
    _append_log(state_dir, {"event": "STAGED", "at": utc_now(), **summary})
    return summary


def _resolve_stage(state_dir: Path, run_id: str | None) -> Tuple[str, Path]:
    state_dir = Path(state_dir)
    if not run_id:
        latest = _read_json(state_dir / "staging" / "latest.json", {})
        run_id = str(latest.get("run_id", ""))
    if not run_id:
        raise FileNotFoundError("No staged update is available.")
    stage_dir = state_dir / "staging" / run_id
    if not stage_dir.exists():
        raise FileNotFoundError(f"Staged update not found: {run_id}")
    return run_id, stage_dir


def approve_stage(state_dir: Path, run_id: str | None = None, *, allow_conflicts: bool = False) -> Dict[str, Any]:
    """Promote one staged package into approved state after explicit user approval."""
    state_dir = Path(state_dir)
    run_id, stage_dir = _resolve_stage(state_dir, run_id)
    summary = _read_json(stage_dir / "summary.json", {})
    flags = _read_json(stage_dir / "flags.json", [])
    blocking = int(summary.get("blocking_flags", 0) or 0)
    non_overridable = sorted(
        {
            str(flag.get("code", ""))
            for flag in flags
            if flag.get("level") == "CONFLICT" and str(flag.get("code", "")) in NON_OVERRIDABLE_CONFLICT_CODES
        }
    )
    if non_overridable:
        raise ValueError(
            "Staged update has non-overridable technical conflict(s): "
            + ", ".join(non_overridable)
            + ". Correct the underlying problem and stage again."
        )
    if blocking and not allow_conflicts:
        raise ValueError(f"Staged update has {blocking} blocking conflict flag(s); approval is not allowed without an explicit override.")

    approved_dir = state_dir / "approved"
    approved_dir.mkdir(parents=True, exist_ok=True)
    for name in ("manifest.json", "cache_index.json", "records.jsonl"):
        src = stage_dir / name
        if not src.exists():
            raise FileNotFoundError(f"Staged package is incomplete: {name}")
        shutil.copy2(src, approved_dir / name)
    approval = {
        "run_id": run_id,
        "decision": "APPROVED",
        "approved_at": utc_now(),
        "allow_conflicts": bool(allow_conflicts),
    }
    _write_json(approved_dir / "approval.json", approval)
    _write_json(stage_dir / "decision.json", approval)
    _append_log(state_dir, {"event": "APPROVED", **approval})
    return approval


def reject_stage(state_dir: Path, run_id: str | None = None, *, note: str = "") -> Dict[str, Any]:
    state_dir = Path(state_dir)
    run_id, stage_dir = _resolve_stage(state_dir, run_id)
    decision = {
        "run_id": run_id,
        "decision": "REJECTED",
        "rejected_at": utc_now(),
        "note": note,
    }
    _write_json(stage_dir / "decision.json", decision)
    _append_log(state_dir, {"event": "REJECTED", **decision})
    return decision


def hold_stage(state_dir: Path, run_id: str | None = None, *, note: str = "") -> Dict[str, Any]:
    state_dir = Path(state_dir)
    run_id, stage_dir = _resolve_stage(state_dir, run_id)
    decision = {
        "run_id": run_id,
        "decision": "HOLD",
        "held_at": utc_now(),
        "note": note,
    }
    _write_json(stage_dir / "decision.json", decision)
    _append_log(state_dir, {"event": "HOLD", **decision})
    return decision

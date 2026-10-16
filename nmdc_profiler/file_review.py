"""Run-bound acceptance of affected files, independent of persistent source scope."""
from __future__ import annotations

import csv
import hashlib
from collections import defaultdict
from pathlib import Path

from .update_engine import (
    _load_approved_state, _manifest_map, _read_json, _read_jsonl,
    _resolve_stage, _write_json_atomic, compare_records,
)

FIELDS = ["Accept Update?", "Project No.", "Source File", "Change Type",
          "Added", "Modified", "Removed", "Owner Note"]


def affected_files(state_dir: Path, run_id: str) -> list[dict]:
    _, stage = _resolve_stage(state_dir, run_id)
    _, _, before = _load_approved_state(state_dir)
    after = _read_jsonl(stage / "records.jsonl")
    old, new = defaultdict(list), defaultdict(list)
    for rows, groups in ((before, old), (after, new)):
        for row in rows:
            groups[str(row.get("Source File", ""))].append(row)
    decisions = {d["relative_path"]: d for d in _read_json(stage / "source_decisions.json", [])}
    result = []
    for source in sorted(set(old) | set(new), key=str.casefold):
        changes, _ = compare_records(old[source], new[source])
        counts = changes["counts"]
        decision = decisions.get(source, {})
        if not any(counts[k] for k in ("added", "modified", "removed")) and decision.get("change_type") not in ("NEW", "CHANGED", "REMOVED"):
            continue
        sample = (new[source] or old[source])[0]
        change_type = decision.get("change_type", "CHANGED")
        if change_type == "UNCHANGED":
            change_type = "CHANGED"
        result.append(dict(zip(FIELDS, [True, sample.get("Project No.", ""), source,
            change_type, counts["added"], counts["modified"], counts["removed"], ""])))
    return result


def _baseline(state_dir: Path, stage: Path) -> str:
    """Bind saved choices to both immutable proposal bytes and approved pointer."""
    digest = hashlib.sha256()
    for path in (stage / "manifest.json", stage / "records.jsonl", stage / "source_decisions.json",
                 Path(state_dir) / "approved" / "current.json"):
        digest.update(path.name.encode())
        digest.update(path.read_bytes() if path.exists() else b"MISSING")
    return digest.hexdigest()


def _check_base(state_dir: Path, stage: Path) -> None:
    summary = _read_json(stage / "summary.json", {})
    approved, _, _ = _load_approved_state(state_dir)
    if summary.get("approved_base_run_id", "") != str(approved.get("run_id", "")):
        raise ValueError("Approved data changed after staging. Stage a new proposal before file review.")


def save_file_review(state_dir: Path, run_id: str, choices_file: Path) -> dict:
    state_dir = Path(state_dir)
    _, stage = _resolve_stage(state_dir, run_id)
    _check_base(state_dir, stage)
    if _read_json(state_dir / "staging" / "latest.json", {}).get("run_id") != run_id:
        raise ValueError("The proposal changed. Review the current run before saving file choices.")
    if _read_json(stage / "decision.json", {}).get("decision") in ("APPROVED", "REJECTED"):
        raise ValueError("This proposal is no longer awaiting file review.")
    files = affected_files(state_dir, run_id)
    expected = {r["Source File"] for r in files}
    choices = {}
    with Path(choices_file).open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not {"Source File", "Accept Update?", "Owner Note"} <= set(reader.fieldnames or []):
            raise ValueError("File review CSV is missing required columns.")
        for row in reader:
            source = row.get("Source File", "")
            if source not in expected or source in choices:
                raise ValueError("Unknown or duplicate file in review choices: " + source)
            value = row.get("Accept Update?", "").strip().upper()
            if value not in ("TRUE", "FALSE"):
                raise ValueError("Accept Update? must be TRUE or FALSE.")
            choices[source] = {"accept": value == "TRUE", "note": row.get("Owner Note", "")}
    if set(choices) != expected:
        raise ValueError("Save a choice for every affected file before approval.")
    payload = {"run_id": run_id, "baseline": _baseline(state_dir, stage), "choices": choices, "files": files}
    _write_json_atomic(stage / "file_review.json", payload)
    return payload


def validated_review(state_dir: Path, run_id: str, *, required: bool = False) -> dict:
    _, stage = _resolve_stage(state_dir, run_id)
    review = _read_json(stage / "file_review.json", {})
    if not review:
        if required:
            raise ValueError("Save File Choices before approval.")
        return {}
    _check_base(state_dir, stage)
    if _read_json(Path(state_dir) / "staging" / "latest.json", {}).get("run_id") != run_id:
        raise ValueError("The proposal was replaced. Review the current run before approval.")
    if review.get("run_id") != run_id or review.get("baseline") != _baseline(state_dir, stage):
        raise ValueError("Saved file review is stale. Review and save the current proposal again.")
    if set(review.get("choices", {})) != {r["Source File"] for r in affected_files(state_dir, run_id)}:
        raise ValueError("Saved file review does not cover the affected files.")
    if any(not isinstance(choice, dict) or type(choice.get("accept")) is not bool
           or not isinstance(choice.get("note"), str) for choice in review["choices"].values()):
        raise ValueError("Saved file choices contain invalid values.")
    return review


def export_file_review(state_dir: Path, output_dir: Path, run_id: str) -> None:
    rows = affected_files(state_dir, run_id) if run_id else []
    if run_id:
        _, stage = _resolve_stage(state_dir, run_id)
        saved = _read_json(stage / "file_review.json", {})
        completed = _read_json(stage / "decision.json", {}).get("decision") == "APPROVED"
        audit = _read_json(Path(state_dir) / "approved" / "versions" / run_id / "file_review.json", {}) if completed else {}
        if audit:
            # Completion deliberately changes the approved pointer. Preserve the
            # reviewed file list and choices from its immutable version audit.
            saved = audit
            rows = [dict(row) for row in audit.get("files", [])]
        # Export only a current baseline; an invalid saved review cannot hide stale choices.
        if saved and (audit or saved.get("baseline") == _baseline(state_dir, stage)):
            for row in rows:
                choice = saved.get("choices", {}).get(row["Source File"], {})
                row["Accept Update?"] = choice.get("accept", True)
                row["Owner Note"] = choice.get("note", "")
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    with (Path(output_dir) / "file_review.csv").open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def merge_deferred_files(state_dir: Path, stage: Path, review: dict) -> tuple[dict, dict, list]:
    approved_manifest, approved_cache, approved_rows = _load_approved_state(state_dir)
    manifest = _read_json(stage / "manifest.json", {})
    cache = _read_json(stage / "cache_index.json", {})
    rows = _read_jsonl(stage / "records.jsonl")
    deferred = {source for source, choice in review.get("choices", {}).items() if not choice["accept"]}
    rows = [r for r in rows if r.get("Source File", "") not in deferred]
    rows.extend(r for r in approved_rows if r.get("Source File", "") in deferred)
    files, old_files = _manifest_map(manifest), _manifest_map(approved_manifest)
    for source in deferred:
        if source in old_files:
            files[source] = dict(old_files[source])
            # Mixed approvals must not falsely claim skipped files use the new parser/config.
            files[source]["parser_version"] = old_files[source].get("parser_version", approved_manifest.get("parser_version"))
            files[source]["config_fingerprint"] = old_files[source].get("config_fingerprint", approved_manifest.get("config_fingerprint"))
        else:
            files.pop(source, None)
        if source in approved_cache:
            cache[source] = approved_cache[source]
        else:
            cache.pop(source, None)
    manifest["files"] = [files[k] for k in sorted(files, key=str.casefold)]
    rows.sort(key=lambda r: tuple(str(r.get(k, "")) for k in
        ("Project No.", "Document No.", "Revision", "Event_Key", "Source File", "Source Row")))
    return manifest, cache, rows

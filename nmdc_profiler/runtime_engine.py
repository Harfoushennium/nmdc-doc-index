from __future__ import annotations

import json
import hashlib
import os
import sys
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Sequence

from .classification import classification_rows, enrich_profiles_with_classification
from .excel_bridge import create_support_request, export_excel_exchange
from .full_extractor import FullWorkItem, build_full_worklist, run_full_extraction
from .ooxml import profile_workbook
from .project_identity import apply_project_identity_overrides, load_project_identity_overrides
from .reporting import write_outputs
from .rules import Rule, load_rules
from .selection import assign_selection
from .source_cache import prepare_local_source_cache
from .source_selection import apply_source_exclusions
from .update_engine import (
    DEFAULT_PARSER_VERSION,
    approve_stage,
    hold_stage,
    reject_stage,
    stage_update,
    utc_now,
    sha256_file,
)


def bundled_root() -> Path:
    frozen_root = getattr(sys, "_MEIPASS", "")
    if frozen_root:
        return Path(frozen_root)
    return Path(__file__).resolve().parents[1]


def resolve_config_file(config_dir: Path, name: str) -> Path:
    external = Path(config_dir) / name
    if external.exists():
        return external
    bundled = bundled_root() / "config" / name
    if bundled.exists():
        return bundled
    raise FileNotFoundError(f"Required configuration file not found: {name}")


def _flag_from_review(review: Mapping[str, Any]) -> Dict[str, str]:
    reason = str(review.get("Reason", "") or "UNRECOGNIZED_LAYOUT")
    return {
        "level": "REVIEW",
        "code": "UNRECOGNIZED_LAYOUT",
        "message": f"The worksheet could not be extracted safely ({reason}).",
        "recommended_action": "Review the worksheet layout and mapping. Do not guess its structure.",
        "source": str(review.get("Source File", "")),
        "source_sheet": str(review.get("Worksheet", "")),
        "project_no": str(review.get("Project No.", "")),
        "document_no": "",
        "revision": "",
        "source_row": "",
        "source_cell": "",
        "event_key": "",
        "resolution_status": "OPEN",
    }


def _inventory_row(workbook: Mapping[str, Any]) -> Dict[str, str]:
    projects = workbook.get("inferred_project_numbers", []) or []
    return {
        "relative_path": str(workbook.get("relative_path", "")),
        "source_family": str(workbook.get("source_family", "")),
        "inferred_project_numbers": ", ".join(str(value) for value in projects),
        "selected_excluded_status": str(workbook.get("selected_excluded_status", "")),
        "selection_exclusion_reason": str(workbook.get("selection_exclusion_reason", "")),
    }


def _normalized_overrides(path: Path) -> Dict[str, Dict[str, str]]:
    raw = load_project_identity_overrides(path)
    normalized: Dict[str, Dict[str, str]] = {}
    for source, decision in raw.items():
        clean = source.replace("\\", "/").lstrip("/")
        normalized[clean] = decision
        if clean.casefold().startswith("data/"):
            normalized[clean[5:]] = decision
    return normalized


@dataclass
class RuntimeCatalog:
    data_dir: Path
    rules: Sequence[Rule]
    selected_inventory: Mapping[str, Mapping[str, str]]
    selection_statuses: Mapping[str, str]
    work_items: Mapping[str, Sequence[FullWorkItem]]
    initial_flags: Sequence[Mapping[str, Any]]
    owner_excluded_sources: Sequence[str]
    _processor_flags: List[Dict[str, str]] = field(default_factory=list)

    def process(self, source_path: Path, relative_path: str) -> Sequence[Mapping[str, Any]]:
        items = list(self.work_items.get(relative_path, ()))
        if not items:
            return []
        inventory = {relative_path: self.selected_inventory[relative_path]}
        records, _reconciliation, reviews = run_full_extraction(
            self.data_dir,
            items,
            inventory,
            [],
            self.rules,
        )
        self._processor_flags.extend(_flag_from_review(review) for review in reviews)
        return records

    def drain_processor_flags(self) -> Sequence[Mapping[str, Any]]:
        flags = list(self._processor_flags)
        self._processor_flags.clear()
        return flags


def _profile_workbooks_cached(files, data_dir, cache_dir, version, refresh=False):
    """Reuse only raw READABLE profiles of exactly matching local file bytes.
    Classification and selection always run fresh on independent profile objects.
    Cache is disposable, never an approved record store; invalid entries fall back.
    """
    cache_dir.mkdir(parents=True, exist_ok=True)
    rows=[]; hits=0
    for path in files:
        rel=path.relative_to(data_dir).as_posix()
        digest=sha256_file(path)
        key=hashlib.sha256((version+"\0"+rel+"\0"+digest).encode()).hexdigest()
        entry=cache_dir/(key+".json")
        profile=None
        try:
            if refresh:raise FileNotFoundError("Full rescan bypasses cached profile")
            payload=json.loads(entry.read_text(encoding="utf-8"))
            candidate=payload["profile"]
            encoded=json.dumps(candidate,sort_keys=True,separators=(",",":"),ensure_ascii=False).encode()
            if (payload["checksum"]==hashlib.sha256(encoded).hexdigest()
                and candidate.get("sha256")==digest and candidate.get("relative_path")==rel
                and candidate.get("readability_status")=="READABLE"):
                profile=candidate
                hits+=1
        except (OSError,ValueError,KeyError,TypeError,AttributeError):
            pass
        if profile is None:
            profile=profile_workbook(path,data_dir,data_dir)
            if profile.get("readability_status")=="READABLE" and profile.get("sha256")==digest:
                encoded=json.dumps(profile,sort_keys=True,separators=(",",":"),ensure_ascii=False).encode()
                temporary=entry.with_suffix(".tmp-"+uuid.uuid4().hex)
                try:
                    temporary.write_text(json.dumps({"profile":profile,"checksum":hashlib.sha256(encoded).hexdigest()},ensure_ascii=False),encoding="utf-8")
                    os.replace(temporary,entry)
                finally:temporary.unlink(missing_ok=True)
        rows.append(profile)
    return rows,hits

def build_runtime_catalog(
    data_dir: Path,
    rules_path: Path,
    overrides_path: Path,
    exclusions_path: Path,
    profile_output_dir: Path,
    *, profile_cache_dir: Path | None = None, profile_cache_version: str = "", timings: dict | None = None, refresh_profiles: bool = False,
) -> RuntimeCatalog:
    data_dir = Path(data_dir).resolve()
    rules = load_rules(rules_path)
    files = sorted(
        (
            path
            for path in data_dir.rglob("*")
            if path.is_file()
            and path.suffix.lower() in {".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
        ),
        key=lambda path: path.relative_to(data_dir).as_posix().casefold(),
    )
    start = time.perf_counter()
    if profile_cache_dir is None:
        workbooks = [profile_workbook(path, data_dir, data_dir) for path in files]; hits=0
    else:
        workbooks,hits = _profile_workbooks_cached(files,data_dir,profile_cache_dir,profile_cache_version,refresh_profiles)
    if timings is not None:timings.update(profile_seconds=round(time.perf_counter()-start,3),profile_cache_hits=hits,profile_cache_misses=len(files)-hits)
    start = time.perf_counter()
    exact_groups, version_groups = assign_selection(workbooks, rules)
    owner_excluded_sources = sorted(apply_source_exclusions(workbooks, exclusions_path), key=str.casefold)
    discovery = classification_rows(workbooks, rules)
    enrich_profiles_with_classification(workbooks, discovery)
    write_outputs(workbooks, discovery, exact_groups, version_groups, profile_output_dir)

    inventory_rows = [_inventory_row(workbook) for workbook in workbooks]
    items, selected_inventory, workbook_review = build_full_worklist(inventory_rows, discovery)
    items, _override_audit = apply_project_identity_overrides(items, _normalized_overrides(overrides_path))

    by_source: Dict[str, List[FullWorkItem]] = {}
    for item in items:
        by_source.setdefault(item.source_file, []).append(item)

    initial_flags: List[Mapping[str, Any]] = [_flag_from_review(row) for row in workbook_review]
    selected_sources = {
        str(workbook.get("relative_path", ""))
        for workbook in workbooks
        if str(workbook.get("selected_excluded_status", "")).upper() == "SELECTED"
    }
    for row in discovery:
        if (
            str(row.get("workbook_path", "")) in selected_sources
            and str(row.get("proposed_action", "")).upper() == "REVIEW_REQUIRED"
        ):
            initial_flags.append(
                _flag_from_review(
                    {
                        "Reason": row.get("warning_code", "REVIEW_REQUIRED"),
                        "Source File": row.get("workbook_path", ""),
                        "Worksheet": row.get("worksheet_name", ""),
                        "Project No.": row.get("project_number", ""),
                    }
                )
            )

    if timings is not None:timings["catalog_selection_seconds"]=round(time.perf_counter()-start,3)
    return RuntimeCatalog(
        data_dir=data_dir,
        rules=rules,
        selected_inventory=selected_inventory,
        selection_statuses={
            str(workbook.get("relative_path", "")): str(workbook.get("selected_excluded_status", ""))
            for workbook in workbooks
        },
        work_items=by_source,
        initial_flags=initial_flags,
        owner_excluded_sources=owner_excluded_sources,
    )


def _restore_original_data_path(state_dir: Path, run_id: str, original_data_dir: Path) -> None:
    """Keep the user-facing manifest bound to the real DATA folder, not the local cache."""
    if not run_id:
        return
    manifest_path = Path(state_dir) / "staging" / run_id / "manifest.json"
    if not manifest_path.exists():
        return
    with manifest_path.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    payload["data_dir"] = str(Path(original_data_dir).resolve())
    with manifest_path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def _suppress_expected_owner_selection_flag(state_dir: Path, run_id: str, owner_excluded_sources: Sequence[str]) -> None:
    """Do not ask the owner to review the exact exclusion decision they just made."""
    if not run_id or not owner_excluded_sources:
        return
    excluded = {str(value).replace("\\", "/").casefold() for value in owner_excluded_sources}
    stage_dir = Path(state_dir) / "staging" / run_id
    flags_path = stage_dir / "flags.json"
    summary_path = stage_dir / "summary.json"
    if not flags_path.exists() or not summary_path.exists():
        return

    with flags_path.open(encoding="utf-8") as handle:
        flags = json.load(handle)
    filtered = [
        flag
        for flag in flags
        if not (
            str(flag.get("code", "")) == "SOURCE_SELECTION_CHANGED"
            and str(flag.get("source", "")).replace("\\", "/").casefold() in excluded
        )
    ]
    if len(filtered) == len(flags):
        return

    with flags_path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(filtered, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")

    with summary_path.open(encoding="utf-8") as handle:
        summary = json.load(handle)
    summary["review_flags"] = sum(1 for flag in filtered if flag.get("level") == "REVIEW")
    summary["blocking_flags"] = sum(1 for flag in filtered if flag.get("level") == "CONFLICT")
    summary["status"] = "REVIEW_REQUIRED" if summary["blocking_flags"] else "STAGED"
    with summary_path.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(summary, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def stage_runtime_update(
    *,
    data_dir: Path,
    state_dir: Path,
    config_dir: Path,
    full_rescan: bool,
    parser_version: str = DEFAULT_PARSER_VERSION,
    reuse_raw_profiles: bool = False,
    checkpoint=None,
) -> Dict[str, Any]:
    phase_times={}
    original_data_dir = Path(data_dir).resolve()
    state_dir = Path(state_dir).resolve()
    rules_path = resolve_config_file(config_dir, "classification_rules.csv")
    overrides_path = resolve_config_file(config_dir, "project_identity_overrides.csv")
    exclusions_path = Path(config_dir) / "source_exclusions.csv"

    # Process from a persistent local mirror. The OneDrive/source tree is only
    # opened when a workbook is new or changed; profiling, hashing and extraction
    # then run against the local cache.
    if checkpoint is not None: checkpoint("Prepare local source cache", 0, 0)
    start=time.perf_counter()
    working_data_dir, cache_stats = prepare_local_source_cache(
        original_data_dir,
        state_dir / "source_cache",
    )
    if checkpoint is not None: checkpoint("Profile and classify", 0, 0)
    phase_times["mirror_seconds"]=round(time.perf_counter()-start,3)
    start=time.perf_counter()
    catalog = build_runtime_catalog(
        working_data_dir,
        rules_path,
        overrides_path,
        exclusions_path,
        state_dir / "profile" / "current",
        profile_cache_dir=(state_dir/"profile"/"raw-v1") if reuse_raw_profiles else None,
        profile_cache_version="raw-v1:"+parser_version,
        timings=phase_times,
        refresh_profiles=full_rescan,
    )
    phase_times["catalog_seconds"]=round(time.perf_counter()-start,3)
    start=time.perf_counter()
    result = stage_update(
        data_dir=working_data_dir,
        state_dir=state_dir,
        processor=catalog.process,
        # Owner source inclusion/exclusion is selection metadata, not parser
        # configuration. It intentionally does not invalidate extraction caches.
        config_paths=[rules_path, overrides_path],
        parser_version=parser_version,
        full_rescan=full_rescan,
        selection_statuses=catalog.selection_statuses,
        initial_flags=catalog.initial_flags,
        post_process_flags=catalog.drain_processor_flags,
        checkpoint=checkpoint,
    )
    phase_times["stage_update_seconds"]=round(time.perf_counter()-start,3)
    run_id = str(result.get("run_id", ""))
    _restore_original_data_path(state_dir, run_id, original_data_dir)
    _suppress_expected_owner_selection_flag(state_dir, run_id, catalog.owner_excluded_sources)
    # Reload the persisted summary after presentation-safe flag suppression.
    summary_path = state_dir / "staging" / run_id / "summary.json"
    if summary_path.exists():
        with summary_path.open(encoding="utf-8") as handle:
            result = json.load(handle)
    result["source_cache"] = cache_stats
    result["phase_times"] = phase_times
    result["data_dir"] = str(original_data_dir)
    result["owner_excluded_sources"] = list(catalog.owner_excluded_sources)
    return result


def _append_jsonl(path: Path, rows: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as handle:
        for row in rows:
            handle.write(json.dumps(dict(row), ensure_ascii=False, sort_keys=True) + "\n")


def resolution_route(current_field: str, flag_code: str) -> str:
    field = current_field.strip().casefold()
    code = flag_code.strip().upper()
    if field in {"project no.", "project no", "project"}:
        return "OWNER_OVERRIDE_REVIEW"
    if field in {"discipline", "category", "subcategory"}:
        return "CONFIGURATION_CHANGE_REVIEW"
    if code in {"PARSER_ERROR", "UNRECOGNIZED_LAYOUT", "MISSING_DOCUMENT_NUMBER"}:
        return "CODE_CHANGE_REQUIRED"
    return "REVIEW_REQUIRED"


def record_user_flag(state_dir: Path, **context: str) -> Dict[str, Any]:
    code = context.get("flag_code", "") or "USER_FLAGGED_WRONG_DATA"
    support_path = create_support_request(state_dir, flag_code=code, **{k: v for k, v in context.items() if k != "flag_code"})
    row = {
        "created_at": utc_now(),
        "level": "REVIEW",
        "code": "USER_FLAGGED_WRONG_DATA",
        "message": context.get("message", "User reported that this record is wrong."),
        "recommended_action": "Review the expected value and apply an approved override, configuration update, or parser correction.",
        "project_no": context.get("project_no", ""),
        "document_no": context.get("document_no", ""),
        "revision": context.get("revision", ""),
        "source": context.get("source_file", ""),
        "source_sheet": context.get("worksheet", ""),
        "source_row": context.get("source_row", ""),
        "source_cell": context.get("source_cell", ""),
        "event_key": context.get("event_identity", ""),
        "user_comment": context.get("message", ""),
        "resolution_status": "OPEN",
        "resolution_route": resolution_route(context.get("current_field", ""), code),
        "support_file": str(support_path),
    }
    _append_jsonl(Path(state_dir) / "user_flags" / "flags.jsonl", [row])
    return row

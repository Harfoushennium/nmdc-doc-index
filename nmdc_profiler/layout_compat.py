from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Callable, List, Mapping, Tuple

from .core import norm_text

_INSTALLED = False

# These are normal lifecycle notifications, not extraction findings. They are
# already represented by Pending Update, dashboard counts and Update History.
_NON_ACTIONABLE_STAGE_CODES = {
    "SOURCE_NEW",
    "SOURCE_CHANGED",
    "PARSER_VERSION_CHANGED",
    "CONFIG_CHANGED",
    "CACHE_MISSING",
}


def _extended_document_header(text: str) -> bool:
    n = norm_text(text)
    if not n:
        return False
    exact = {
        "procedure no",
        "procedure number",
        "method statement no",
        "method statement number",
        "setup plan no",
        "setup plan number",
        "anchor pattern no",
        "anchor pattern number",
        "cut list no",
        "cut list number",
        "cut-list no",
        "cut-list number",
        "nmdc energy number",
        "nmdc energy no",
        "nmdc document no",
        "nmdc document number",
        "nmdc doc no",
        "nmdc doc number",
        "npcc document no",
        "npcc document number",
        "npcc doc no",
        "npcc doc number",
        "contractor document no",
        "contractor document number",
        "contractor doc no",
        "contractor doc number",
    }
    if n in exact:
        return True
    if any(label in n for label in ("procedure", "method statement", "setup plan", "anchor pattern", "cut list", "cut-list")):
        return any(token in n for token in (" no", " number", " ref", " reference"))
    return False


def _review_diagnostics(review: Mapping[str, Any]) -> str:
    reason = str(review.get("Reason", "") or "")
    warnings = review.get("Warnings", "")
    if isinstance(warnings, (list, tuple)):
        warning_text = ";".join(str(value) for value in warnings if value)
    else:
        warning_text = str(warnings or "")
    return f"{reason};{warning_text}".upper()


def _date_like_identifier(value: str) -> bool:
    text = str(value or "").strip()
    return bool(
        re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", text)
        or re.fullmatch(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", text)
        or re.fullmatch(r"\d{1,2}-[A-Za-z]{3}-\d{2,4}", text)
    )


def _header_candidate_score(model, row: int, extractor) -> int:
    """Prefer a real table header over document-number labels in workbook title blocks."""
    score = 0
    for (candidate_row, _col), value in model.cells.items():
        if candidate_row not in {row, row + 1}:
            continue
        weight = 2 if candidate_row == row else 1
        if extractor._is_title_header(value):
            score += 12 * weight
        if extractor._is_company_doc_header(value):
            score += 10 * weight
        if extractor._is_revision_header(value):
            score += 8 * weight
        n = norm_text(value)
        if any(token in n for token in ("status", "submission plan", "issue date", "schedule date", "transmittal")):
            score += 2 * weight
    return score


def _safe_empty_register(root: Path, review: Mapping[str, Any], extractor, full_extractor) -> bool:
    """Return True only when recognized identity columns contain no real record values."""
    if "LAYOUT_FIRST_DATA_ROW_NOT_FOUND" not in _review_diagnostics(review):
        return False

    source_file = str(review.get("Source File", "") or "").strip()
    worksheet = str(review.get("Worksheet", "") or "").strip()
    if not source_file or not worksheet:
        return False

    source = Path(root) / source_file
    if not source.exists():
        return False

    try:
        actual_sheet = full_extractor.resolve_sheet_name(source, worksheet)
        model = extractor.read_sheet_model(source, Path(root), actual_sheet)
    except Exception:
        return False

    headers = extractor._header_candidates(model, extractor._is_doc_header)
    if not headers:
        return False
    header_start, document_col = headers[0]
    band_end = min(model.max_row, header_start + 4)
    company_candidates = [
        (row, col)
        for (row, col), value in model.cells.items()
        if header_start <= row <= band_end and extractor._is_company_doc_header(value)
    ]
    identity_cols = {document_col}
    if company_candidates:
        identity_cols.add(sorted(company_candidates)[0][1])

    placeholder_values = {"-", "--", "n/a", "na", "nil", "none", "tbd", "to be advised", "to be confirmed"}
    for row in range(header_start + 1, model.max_row + 1):
        for col in identity_cols:
            text = str(model.value(row, col) or "").strip()
            if not text:
                continue
            if norm_text(text) in placeholder_values or _date_like_identifier(text):
                continue
            if extractor._looks_identifier(text):
                return False
            if row > band_end and not (
                extractor._is_doc_header(text)
                or extractor._is_company_doc_header(text)
                or extractor._is_title_header(text)
                or extractor._is_revision_header(text)
            ):
                return False
    return True


def _rewrite_latest_stage_without_notifications(state_dir: Path, summary: Mapping[str, Any]) -> dict[str, Any]:
    """Keep Review Flags for actionable anomalies only.

    New/changed/reprocessed-source notifications remain fully visible in Pending
    Update, dashboard counts, source decisions and Update History. They are not
    parser findings and therefore must not ask the user for a Review Flags
    decision.
    """
    corrected = dict(summary)
    run_id = str(corrected.get("run_id", ""))
    if not run_id:
        return corrected
    stage_dir = Path(state_dir) / "staging" / run_id
    flags_path = stage_dir / "flags.json"
    summary_path = stage_dir / "summary.json"
    if not flags_path.exists() or not summary_path.exists():
        return corrected

    try:
        flags = json.loads(flags_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return corrected
    kept = [
        dict(flag)
        for flag in flags
        if str(flag.get("code", "")).strip().upper() not in _NON_ACTIONABLE_STAGE_CODES
    ]
    if len(kept) == len(flags):
        return corrected

    blocking = sum(1 for flag in kept if str(flag.get("level", "")).upper() == "CONFLICT")
    review_count = sum(1 for flag in kept if str(flag.get("level", "")).upper() == "REVIEW")
    corrected["blocking_flags"] = blocking
    corrected["review_flags"] = review_count
    corrected["status"] = "REVIEW_REQUIRED" if blocking else "STAGED"

    flags_path.write_text(json.dumps(kept, indent=2, ensure_ascii=False, sort_keys=True) + "\n", encoding="utf-8")
    summary_path.write_text(json.dumps(corrected, indent=2, ensure_ascii=False, sort_keys=True) + "\n", encoding="utf-8")

    # Keep the most recent STAGED history row consistent with the corrected
    # summary so dashboard/history counts never disagree with Review Flags.
    history_path = Path(state_dir) / "logs" / "history.jsonl"
    if history_path.exists():
        try:
            lines = history_path.read_text(encoding="utf-8").splitlines()
            for index in range(len(lines) - 1, -1, -1):
                row = json.loads(lines[index])
                if str(row.get("event", "")) == "STAGED" and str(row.get("run_id", "")) == run_id:
                    row["blocking_flags"] = blocking
                    row["review_flags"] = review_count
                    row["status"] = corrected["status"]
                    lines[index] = json.dumps(row, ensure_ascii=False, sort_keys=True)
                    history_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
                    break
        except (OSError, json.JSONDecodeError):
            pass
    return corrected


def install_layout_compatibility() -> None:
    """Install conservative runtime layout and clean-review compatibility.

    Known engineering identifier labels are recognized, real table headers are
    preferred over title-block labels, and empty registers are accepted only
    after identity columns are proven to contain no document rows. Review Flags
    are reserved for actionable anomalies; ordinary source lifecycle notices
    remain in Pending Update/history instead. The known standalone barge-sketch
    register is selected without inventing a project number because its document
    identity is independent of a project folder/name.
    """
    global _INSTALLED
    if _INSTALLED:
        return

    from . import extractor, full_extractor, runtime_engine

    original_is_doc_header = extractor._is_doc_header
    original_header_candidates = extractor._header_candidates
    original_flag_from_review = runtime_engine._flag_from_review
    original_runtime_run_full_extraction = runtime_engine.run_full_extraction
    original_stage_update = runtime_engine.stage_update
    original_assign_selection = runtime_engine.assign_selection

    def is_doc_header(text: str) -> bool:
        return original_is_doc_header(text) or _extended_document_header(text)

    def header_candidates(model, predicate: Callable[[str], bool]) -> List[Tuple[int, int]]:
        found = original_header_candidates(model, predicate)
        if not found:
            found = []
            for (row, col), value in model.cells.items():
                if row <= min(model.max_row, 60) and predicate(value):
                    found.append((row, col))
        if predicate is is_doc_header and found:
            return sorted(
                found,
                key=lambda item: (-_header_candidate_score(model, item[0], extractor), item[0], item[1]),
            )
        return sorted(found)

    def runtime_assign_selection(workbooks, rules):
        exact_groups, version_groups = original_assign_selection(workbooks, rules)
        for workbook in workbooks:
            if (
                str(workbook.get("selected_excluded_status", "")).upper() == "REVIEW_REQUIRED"
                and str(workbook.get("logical_register_identity", "")).upper() == "BARGE_SKETCH_REGISTER"
                and not list(workbook.get("inferred_project_numbers", []) or [])
            ):
                workbook["selected_excluded_status"] = "SELECTED"
                workbook["selection_exclusion_reason"] = (
                    "Standalone barge sketch register; project identity is not required for document extraction"
                )
                workbook["warnings"] = [
                    warning
                    for warning in list(workbook.get("warnings", []) or [])
                    if warning not in {"PROJECT_ID_NOT_INFERRED", "PROJECT_ID_INTERNAL_NOT_FOUND"}
                ]
        return exact_groups, version_groups

    def review_flag(review: Mapping[str, Any]) -> dict[str, str]:
        flag = original_flag_from_review(review)
        diagnostics = _review_diagnostics(review)
        if "LAYOUT_DOCUMENT_HEADER_NOT_FOUND" in diagnostics:
            flag["code"] = "UNRECOGNIZED_LAYOUT_HEADER"
            flag["message"] = (
                "The worksheet contains candidate register content, but the parser could not find a recognized "
                "document-number/identifier header in the header area."
            )
            flag["recommended_action"] = (
                "NEEDS PARSER/MAPPING FIX: this is an extraction defect requiring parser/mapping correction; "
                "do not manually classify the source as bad."
            )
        elif "LAYOUT_FIRST_DATA_ROW_NOT_FOUND" in diagnostics:
            flag["code"] = "UNRECOGNIZED_LAYOUT_DATA"
            flag["message"] = (
                "The worksheet header was recognized, but no safe first data row could be confirmed."
            )
            flag["recommended_action"] = (
                "This is an extraction defect unless the register is empty; the engine should resolve known empty registers automatically."
            )
        return flag

    def runtime_run_full_extraction(root, work_items, selected_inventory, initial_review, rules):
        records, reconciliation, reviews = original_runtime_run_full_extraction(
            root, work_items, selected_inventory, initial_review, rules
        )

        accepted_empty: set[tuple[str, str]] = set()
        kept_reviews = []
        for review in reviews:
            if _safe_empty_register(Path(root), review, extractor, full_extractor):
                accepted_empty.add(
                    (
                        str(review.get("Source File", "")).casefold(),
                        re.sub(r"\s+", " ", str(review.get("Worksheet", ""))).strip().casefold(),
                    )
                )
            else:
                kept_reviews.append(review)

        if accepted_empty:
            reconciliation = dict(reconciliation)
            worksheets = [dict(row) for row in reconciliation.get("worksheets", [])]
            for row in worksheets:
                key = (
                    str(row.get("source_file", "")).casefold(),
                    re.sub(r"\s+", " ", str(row.get("requested_worksheet", row.get("worksheet", "")))).strip().casefold(),
                )
                if key in accepted_empty:
                    row["status"] = "EMPTY_ACCEPTED"
                    row["reason"] = "Recognized register contains no document rows"
                    warnings = list(row.get("warnings", []) or [])
                    if "EMPTY_REGISTER_NO_DATA" not in warnings:
                        warnings.append("EMPTY_REGISTER_NO_DATA")
                    row["warnings"] = warnings
            reconciliation["worksheets"] = worksheets

            summary = dict(reconciliation.get("summary", {}))
            summary["worksheets_review_required"] = sum(
                1
                for row in worksheets
                if str(row.get("status", "")) not in {"INCLUDE", "EMPTY_ACCEPTED"}
                or (
                    str(row.get("status", "")) == "INCLUDE"
                    and int(row.get("event_records", 0) or 0) <= 0
                )
            )
            summary["empty_worksheets_accepted"] = len(accepted_empty)
            summary["review_queue_items"] = len(kept_reviews)
            reconciliation["summary"] = summary

        return records, reconciliation, kept_reviews

    def clean_stage_update(*args, **kwargs):
        summary = original_stage_update(*args, **kwargs)
        state_dir = kwargs.get("state_dir")
        if state_dir is None:
            return summary
        return _rewrite_latest_stage_without_notifications(Path(state_dir), summary)

    extractor._is_doc_header = is_doc_header
    extractor._header_candidates = header_candidates
    runtime_engine.assign_selection = runtime_assign_selection
    runtime_engine._flag_from_review = review_flag
    runtime_engine.run_full_extraction = runtime_run_full_extraction
    runtime_engine.stage_update = clean_stage_update
    _INSTALLED = True

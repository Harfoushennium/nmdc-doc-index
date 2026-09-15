from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Callable, List, Mapping, Tuple

from .core import norm_text

_INSTALLED = False


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
    # Only inspect the candidate row and the immediately following row. Looking
    # four rows ahead can make a title-block label inherit the score of the real
    # table header below it (as happened in the 7279 NAVAL MARINE sheet).
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


def install_layout_compatibility() -> None:
    """Install conservative runtime-only layout compatibility extensions.

    Existing recognized layouts are untouched. The fallback broadens known
    document-identifier labels, prefers table headers over title-block labels,
    and retries header discovery through row 60. Runtime review filtering closes
    a worksheet automatically only when its recognized identity columns contain
    no real records. Populated unfamiliar layouts remain in Review Flags.
    """
    global _INSTALLED
    if _INSTALLED:
        return

    from . import extractor, full_extractor, runtime_engine

    original_is_doc_header = extractor._is_doc_header
    original_header_candidates = extractor._header_candidates
    original_flag_from_review = runtime_engine._flag_from_review
    original_runtime_run_full_extraction = runtime_engine.run_full_extraction

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

    def review_flag(review: Mapping[str, Any]) -> dict[str, str]:
        flag = original_flag_from_review(review)
        diagnostics = _review_diagnostics(review)
        if "LAYOUT_DOCUMENT_HEADER_NOT_FOUND" in diagnostics:
            flag["code"] = "UNRECOGNIZED_LAYOUT_HEADER"
            flag["message"] = (
                "The worksheet contains candidate register content, but the parser could not find a recognized "
                "document-number/identifier header in the header area. The runtime also checks common engineering "
                "labels such as NMDC Energy Number, NPCC Doc No., Contractor Document No., Procedure No., Setup Plan No., "
                "Anchor Pattern No. and Cut List No."
            )
            flag["recommended_action"] = (
                "Open the Source File and named Source Sheet. If it contains register data, compare the identifier "
                "column heading with an extracted project and choose NEEDS PARSER/MAPPING FIX with a short comment. "
                "If the sheet is intentionally empty, choose NO ACTION REQUIRED."
            )
        elif "LAYOUT_FIRST_DATA_ROW_NOT_FOUND" in diagnostics:
            flag["code"] = "UNRECOGNIZED_LAYOUT_DATA"
            flag["message"] = (
                "The worksheet header was recognized, but no safe first data row with a document/company identifier "
                "could be confirmed. This commonly means the sheet is empty, uses placeholder-only rows, or its "
                "document identifiers use a pattern the parser does not yet recognize."
            )
            flag["recommended_action"] = (
                "Open the Source File and named Source Sheet. If there are real document rows, choose "
                "NEEDS PARSER/MAPPING FIX and add an example document number in User Comment. If there are no real "
                "records, choose NO ACTION REQUIRED."
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

    extractor._is_doc_header = is_doc_header
    extractor._header_candidates = header_candidates
    runtime_engine._flag_from_review = review_flag
    runtime_engine.run_full_extraction = runtime_run_full_extraction
    _INSTALLED = True

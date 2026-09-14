from __future__ import annotations

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


def install_layout_compatibility() -> None:
    """Install conservative runtime-only layout compatibility extensions.

    Existing recognized layouts are untouched. The fallback only broadens known
    document-identifier header labels and, when the original first-30-row scan
    finds nothing, retries the same predicate through row 60. Runtime review
    flags also preserve the concrete parser reason so Excel tells the owner why
    a worksheet needs review instead of showing only a generic layout message.
    """
    global _INSTALLED
    if _INSTALLED:
        return

    from . import extractor, runtime_engine

    original_is_doc_header = extractor._is_doc_header
    original_header_candidates = extractor._header_candidates
    original_flag_from_review = runtime_engine._flag_from_review

    def is_doc_header(text: str) -> bool:
        return original_is_doc_header(text) or _extended_document_header(text)

    def header_candidates(model, predicate: Callable[[str], bool]) -> List[Tuple[int, int]]:
        found = original_header_candidates(model, predicate)
        if found:
            return found
        out: List[Tuple[int, int]] = []
        for (row, col), value in model.cells.items():
            if row <= min(model.max_row, 60) and predicate(value):
                out.append((row, col))
        return sorted(out)

    def review_flag(review: Mapping[str, Any]) -> dict[str, str]:
        flag = original_flag_from_review(review)
        diagnostics = _review_diagnostics(review)
        if "LAYOUT_DOCUMENT_HEADER_NOT_FOUND" in diagnostics:
            flag["code"] = "UNRECOGNIZED_LAYOUT_HEADER"
            flag["message"] = (
                "The worksheet contains candidate register content, but the parser could not find a recognized "
                "document-number/identifier header in the header area. The runtime also checks common engineering "
                "labels such as Procedure No., Setup Plan No., Anchor Pattern No. and Cut List No."
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

    extractor._is_doc_header = is_doc_header
    extractor._header_candidates = header_candidates
    runtime_engine._flag_from_review = review_flag
    _INSTALLED = True

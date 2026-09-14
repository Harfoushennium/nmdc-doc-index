from __future__ import annotations

from typing import Callable, List, Tuple

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


def install_layout_compatibility() -> None:
    """Install conservative runtime-only layout compatibility extensions.

    Existing recognized layouts are untouched. The fallback only broadens known
    document-identifier header labels and, when the original first-30-row scan
    finds nothing, retries the same predicate through row 60.
    """
    global _INSTALLED
    if _INSTALLED:
        return

    from . import extractor

    original_is_doc_header = extractor._is_doc_header
    original_header_candidates = extractor._header_candidates

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

    extractor._is_doc_header = is_doc_header
    extractor._header_candidates = header_candidates
    _INSTALLED = True

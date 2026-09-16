from __future__ import annotations

from typing import Dict, List


# User-facing order follows the review workflow: identify the record first,
# understand the business content second, then show provenance/technical keys.
DOCUMENT_FIELDS: List[str] = [
    "Project No.",
    "Document No.",
    "Document Title",
    "Latest Revision",
    "Latest Event Status",
    "Latest Event Date",
    "Discipline",
    "Category",
    "Subcategory",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Document Link",
    "Company Document No.",
    "Source Family",
    "Flag Level",
    "Global Document Key",
    "Source Cell",
]

REVISION_FIELDS: List[str] = [
    "Project No.",
    "Document No.",
    "Revision",
    "Is Latest Revision",
    "Document Title",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Flag Level",
    "Revision Key",
    "Source Cell",
]

EVENT_FIELDS: List[str] = [
    "Project No.",
    "Document No.",
    "Revision",
    "Event Type",
    "Event Date",
    "Event Status",
    "Event Reference",
    "Document Title",
    "Source File",
    "Source Sheet",
    "Source Row",
    "Document Link",
    "Is Latest Event",
    "Flag Level",
    "Event Key",
    "Event Values JSON",
    "Source Cell",
]

PENDING_FIELDS: List[str] = [
    "Change Type",
    "Project No.",
    "Document No.",
    "Revision",
    "Event Type",
    "Plain-English Summary",
    "Source File",
    "Review Required",
    "Record Identity",
]

FLAG_FIELDS: List[str] = [
    "Flag Level",
    "Plain-English Problem",
    "Recommended User Action",
    "User Decision",
    "User Comment",
    "Resolution Status",
    "Source File",
    "Source Sheet",
    "Project No.",
    "Document No.",
    "Revision",
    "Flag Code",
    "Source Row",
    "Source Cell",
    "Event Key",
]


def preferred_orders() -> Dict[str, List[str]]:
    return {
        "master_documents.csv": list(DOCUMENT_FIELDS),
        "revisions.csv": list(REVISION_FIELDS),
        "events.csv": list(EVENT_FIELDS),
        "pending_update.csv": list(PENDING_FIELDS),
        "flags.csv": list(FLAG_FIELDS),
    }


def install_review_first_column_order() -> None:
    """Apply the user-facing field order without changing canonical record data."""
    from . import excel_bridge

    excel_bridge.DOCUMENT_FIELDS = list(DOCUMENT_FIELDS)
    excel_bridge.REVISION_FIELDS = list(REVISION_FIELDS)
    excel_bridge.EVENT_FIELDS = list(EVENT_FIELDS)
    excel_bridge.PENDING_FIELDS = list(PENDING_FIELDS)
    excel_bridge.FLAG_FIELDS = list(FLAG_FIELDS)

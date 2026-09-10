from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DISCOVERY = ROOT / "outputs" / "cycle1" / "classification_discovery.csv"


def load_rows():
    with DISCOVERY.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def find_one(rows, *, path=None, worksheet=None, section=None):
    matches = []
    for row in rows:
        if path is not None and row["workbook_path"] != path:
            continue
        if worksheet is not None and row["worksheet_name"] != worksheet:
            continue
        if section is not None and row["original_section"] != section:
            continue
        matches.append(row)
    if len(matches) != 1:
        raise AssertionError(
            f"Expected exactly one row for path={path!r}, worksheet={worksheet!r}, "
            f"section={section!r}; found {len(matches)}"
        )
    return matches[0]


def assert_taxonomy(row, discipline, category, subcategory):
    actual = (row["discipline"], row["category"], row["subcategory"])
    expected = (discipline, category, subcategory)
    if actual != expected:
        raise AssertionError(
            f"Unexpected taxonomy for {row['workbook_path']} | {row['worksheet_name']} | "
            f"{row['original_section']}: expected {expected}, got {actual}"
        )


def main():
    rows = load_rows()
    if len(rows) != 211:
        raise AssertionError(f"Expected 211 classification rows, got {len(rows)}")

    review = [r for r in rows if r["proposed_action"] == "REVIEW_REQUIRED"]
    if review:
        raise AssertionError(f"Expected 0 classification REVIEW_REQUIRED rows, got {len(review)}")

    # Broad TECH document worksheets must stay at safe discovery bases. Sampled
    # report/procedure/analysis titles must never collapse the whole worksheet.
    pipeline = [r for r in rows if r["source_family"] == "TECH" and r["normalized_worksheet"] == "documents pipeline and cable" and r["proposed_action"] == "INCLUDE"]
    if not pipeline:
        raise AssertionError("No Pipeline & Cable document discovery rows found")
    for row in pipeline:
        assert_taxonomy(row, "PIPELINE & CABLE", "DOCUMENT", "GENERAL TECHNICAL DOCUMENT")
        if "TITLE:" in row["match_basis"] or "DOC_NUMBER:" in row["match_basis"]:
            raise AssertionError("Worksheet discovery must not use TITLE/DOC_NUMBER refinement evidence")

    naval = [r for r in rows if r["source_family"] == "TECH" and r["normalized_worksheet"] in {"documents naval marine", "documents naval and marine", "naval marine"} and r["proposed_action"] == "INCLUDE"]
    if not naval:
        raise AssertionError("No Naval & Marine document discovery rows found")
    for row in naval:
        assert_taxonomy(row, "NAVAL & MARINE", "DOCUMENT", "GENERAL TECHNICAL DOCUMENT")
        if "TITLE:" in row["match_basis"] or "DOC_NUMBER:" in row["match_basis"]:
            raise AssertionError("Worksheet discovery must not use TITLE/DOC_NUMBER refinement evidence")

    generic_docs = [r for r in rows if r["source_family"] == "TECH" and r["normalized_worksheet"] == "documents" and r["proposed_action"] == "INCLUDE"]
    if not generic_docs:
        raise AssertionError("No generic TECH Documents discovery rows found")
    for row in generic_docs:
        assert_taxonomy(row, "GENERAL / MULTIDISCIPLINE", "DOCUMENT", "GENERAL TECHNICAL DOCUMENT")

    methods_sketch = [r for r in rows if r["source_family"] == "METHODS" and r["normalized_worksheet"] == "sketches" and r["proposed_action"] == "INCLUDE"]
    if not methods_sketch:
        raise AssertionError("No METHODS Sketches discovery rows found")
    for row in methods_sketch:
        assert_taxonomy(row, "OFFSHORE INSTALLATION", "SKETCH", "ENGINEERING SKETCH")

    setup_rows = [r for r in rows if r["source_family"] == "METHODS" and r["normalized_worksheet"] == "setup plans and anchor patterns" and r["proposed_action"] == "INCLUDE"]
    if not setup_rows:
        raise AssertionError("No METHODS setup/anchor discovery rows found")
    for row in setup_rows:
        assert_taxonomy(row, "MARINE OPERATIONS", "DRAWING", "METHOD DRAWING")
        if row["matched_rule_ids"] != "M010":
            raise AssertionError(f"Expected worksheet-only M010 discovery, got {row['matched_rule_ids']}")

    incoming_path = "DATA/METHODS/1 Completed Project  Deliverables/2891- BU HASEER Delivarables.xlsx"
    incoming_doc = find_one(rows, path=incoming_path, worksheet="Incomming DOC and DRG", section="DOCUMENTS")
    assert_taxonomy(incoming_doc, "EXTERNAL / INPUT", "DOCUMENT", "INCOMING TECHNICAL DOCUMENT")
    incoming_drg = find_one(rows, path=incoming_path, worksheet="Incomming DOC and DRG", section="DRAWINGS")
    assert_taxonomy(incoming_drg, "EXTERNAL / INPUT", "DRAWING", "INCOMING DRAWING")

    special = find_one(
        rows,
        path="DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx",
        worksheet="2171-2172",
        section="",
    )
    assert_taxonomy(special, "OFFSHORE INSTALLATION", "PROCEDURE", "INSTALLATION PROCEDURE")
    if special["matched_rule_ids"] != "M002" or "WORKSHEET:M002" not in special["match_basis"]:
        raise AssertionError("2171-2172 must use the narrow structural M002 worksheet exception")

    client = find_one(
        rows,
        path="DATA/TECH/3291 DOCUMENT REGISTER Latest.xlsx",
        worksheet="CLIENT",
        section="",
    )
    if client["proposed_action"] != "EXCLUDE" or client["matched_rule_ids"] != "X017":
        raise AssertionError("3291 CLIENT must be excluded only by scoped X017")

    print("Real-data semantic sentinels: PASS")
    print(f"Rows checked: {len(rows)}")
    print(f"Pipeline broad worksheets: {len(pipeline)}")
    print(f"Naval broad worksheets: {len(naval)}")
    print(f"Generic TECH Documents worksheets: {len(generic_docs)}")
    print(f"METHODS Sketches worksheets: {len(methods_sketch)}")
    print(f"METHODS setup/anchor worksheets: {len(setup_rows)}")


if __name__ == "__main__":
    main()

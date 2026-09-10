from __future__ import annotations

import csv
import json
from pathlib import Path


def fail(message: str) -> None:
    raise SystemExit(f"Cycle 3 real-data verification FAILED: {message}")


def read_csv(path: Path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    inventory = read_csv(root / "outputs" / "cycle1" / "source_inventory.csv")
    discovery = read_csv(root / "outputs" / "cycle1" / "classification_discovery.csv")
    records = read_csv(root / "outputs" / "cycle3" / "full_records.csv")
    review = read_csv(root / "outputs" / "cycle3" / "review_queue.csv")
    reconciliation = json.loads((root / "outputs" / "cycle3" / "full_reconciliation.json").read_text(encoding="utf-8"))
    summary = reconciliation["summary"]

    selected = {r["relative_path"] for r in inventory if r["selected_excluded_status"].strip().upper() == "SELECTED"}
    if len(selected) != 46:
        fail(f"expected 46 selected workbooks, got {len(selected)}")
    if int(summary["selected_workbooks"]) != len(selected):
        fail(f"summary selected count mismatch: {summary['selected_workbooks']} vs {len(selected)}")

    record_sources = {r["Source File"] for r in records}
    leaked = sorted(record_sources - selected)
    if leaked:
        fail(f"non-selected workbooks leaked into full records: {leaked[:10]}")

    forbidden_fragments = [
        "2820-DOCUMENT REGISTER.xlsx",
        "3291 DOCUMENT REGISTER revised.XLSX",
        "Harfoush/ANCHOR PATTERN Procedure status",
        "1. Delivarables  FORMAT.xlsx",
        "2722 Deliverable.xlsx",
    ]
    for fragment in forbidden_fragments:
        if any(fragment.casefold() in src.casefold() for src in record_sources):
            fail(f"forbidden/excluded/superseded source leaked into records: {fragment}")

    if any(r["Source File"].endswith("3291 DOCUMENT REGISTER Latest.xlsx") and r["Source Sheet"].strip().upper() == "CLIENT" for r in records):
        fail("scoped 3291 CLIENT worksheet leaked into full records")

    # Every selected workbook must be represented either by extracted records or the explicit review queue.
    review_sources = {r["Source File"] for r in review}
    accounted = record_sources | review_sources
    missing_selected = sorted(selected - accounted)
    if missing_selected:
        fail(f"selected workbooks not accounted for: {missing_selected}")

    # Every unique INCLUDE worksheet for a selected workbook must have reconciliation evidence.
    def canon(value: str) -> str:
        return " ".join((value or "").split()).casefold()

    include_sheets = {
        (r["workbook_path"], canon(r["worksheet_name"]))
        for r in discovery
        if r["workbook_path"] in selected and r["proposed_action"].strip().upper() == "INCLUDE" and r["worksheet_name"] != "[UNREADABLE]"
    }
    reconciled_sheets = {
        (str(r.get("source_file", "")), canon(str(r.get("requested_worksheet", r.get("worksheet", "")))))
        for r in reconciliation["worksheets"]
    }
    missing_sheets = sorted(include_sheets - reconciled_sheets)
    if missing_sheets:
        fail(f"included worksheets missing reconciliation: {missing_sheets[:10]}")

    event_keys = [r["Event_Key"] for r in records]
    if len(event_keys) != len(set(event_keys)):
        fail("Event_Key is not unique")
    source_docs = {r["Source_Document_Key"] for r in records}
    revisions = {r["Revision_Key"] for r in records}
    if sum(int(r["Document_Row_Flag"] or 0) for r in records) != len(source_docs):
        fail("Document_Row_Flag does not select exactly one row per source document")
    if sum(int(r["Revision_Row_Flag"] or 0) for r in records) != len(revisions):
        fail("Revision_Row_Flag does not select exactly one row per revision")

    if int(summary["event_key_duplicates"]) != 0:
        fail(f"summary reports duplicate event keys: {summary['event_key_duplicates']}")
    if summary["leaked_nonselected_sources"]:
        fail(f"summary reports non-selected leaks: {summary['leaked_nonselected_sources']}")
    if summary["unaccounted_selected_workbooks"]:
        fail(f"summary reports unaccounted selected workbooks: {summary['unaccounted_selected_workbooks']}")

    # Cycle-2 sentinel semantics must still be present in the full extraction.
    p2369 = [r for r in records if r["Source File"].endswith("2369 - NMGL Delivarables.xlsx") and r["Document No."] == "2369-PP-OF-003"]
    revisions_2369 = {r["Revision"] for r in p2369}
    if not {"A1", "1", "2", "3"}.issubset(revisions_2369):
        fail(f"2369 sentinel revision history missing from full run: {sorted(revisions_2369)}")

    incoming = [r for r in records if r["Source File"].endswith("2891- BU HASEER Delivarables.xlsx") and r["Source Sheet"].strip() == "Incomming DOC and DRG"]
    if not {"DOCUMENTS", "DRAWINGS"}.issubset({r["Original Section"] for r in incoming}):
        fail("2891 incoming DOCUMENTS/DRAWINGS sections are not both preserved")

    p2171 = [r for r in records if r["Source File"].endswith("2171-2172 -Document Deliverables LATEST.xlsx") and r["Document No."] == "2171-2172-PP-OF-012" and r["Revision"] == "A1"]
    if not p2171:
        fail("2171-2172 known A1 record missing from full extraction")
    first_2171 = p2171[0]
    if first_2171["Event Type"] != "Construction and Installation Procedure" or first_2171["Event Date"] != "2021-12-23" or first_2171["Event Reference"] != "T-553/21":
        fail(f"2171-2172 consolidated event semantics changed: {first_2171}")

    p2820 = [r for r in records if r["Source File"].endswith("2820-DOCUMENT REGISTER-NEW 30-04-2026.xlsx") and r["Document No."] == "2820-NN-RP-001"]
    if not p2820:
        fail("selected 2820 sentinel document missing")
    if any(r["Discipline"] != "PIPELINE & CABLE" or r["Category"] != "DOCUMENT" for r in p2820):
        fail("2820 classification changed in full extraction")

    # Review queue is allowed in Cycle 3, but it must be explicit and internally consistent.
    summary_review = int(summary["review_queue_items"])
    if summary_review != len(review):
        fail(f"review queue count mismatch: summary={summary_review}, csv={len(review)}")

    final_xlsx_candidates = list(root.glob("**/NMDC_DOCUMENT_INDEX.xlsx"))
    if final_xlsx_candidates:
        fail(f"final XLSX must not be produced in Cycle 3: {final_xlsx_candidates}")

    print("Cycle 3 real-data verification: PASS")
    print(
        f"selected={len(selected)} processed={summary['processed_workbooks']} "
        f"sheets={summary['included_worksheets_attempted']} extracted={summary['included_worksheets_extracted']} "
        f"sheet_review={summary['worksheets_review_required']} review_items={len(review)}"
    )
    print(
        f"records={len(records)} source_docs={len(source_docs)} global_docs={summary['distinct_global_documents']} "
        f"revisions={len(revisions)} hyperlinks={summary['hyperlinks_preserved']} row_review={summary['row_level_review_records']}"
    )
    if review:
        print("Review queue reasons:")
        reasons = {}
        for item in review:
            reasons[item["Reason"]] = reasons.get(item["Reason"], 0) + 1
        for reason, count in sorted(reasons.items()):
            print(f"  {reason}: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

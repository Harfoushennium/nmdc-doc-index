from __future__ import annotations

import csv
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "cycle2"


def fail(msg: str) -> None:
    raise SystemExit(msg)


def main() -> int:
    records_path = OUT / "sentinel_records.csv"
    recon_path = OUT / "sentinel_reconciliation.json"
    report_path = OUT / "sentinel_report.md"
    for path in (records_path, recon_path, report_path):
        if not path.exists():
            fail(f"Missing Cycle 2 output: {path.relative_to(ROOT)}")

    with records_path.open(newline="", encoding="utf-8-sig") as f:
        records = list(csv.DictReader(f))
    reconciliation = json.loads(recon_path.read_text(encoding="utf-8"))
    by_case = {r["case_id"]: r for r in reconciliation}

    required_cases = {"M2369_PROC", "M2891_INCOMING", "M2171_NONSTANDARD", "T2820_PIPELINE", "T3291_CLIENT"}
    if set(by_case) != required_cases:
        fail(f"Unexpected sentinel case set: {sorted(by_case)}")

    for case_id in required_cases - {"T3291_CLIENT"}:
        item = by_case[case_id]
        if item["status"] != "INCLUDE" or int(item["event_records"]) <= 0 or int(item["distinct_documents"]) <= 0:
            fail(f"{case_id} did not produce included document events: {item}")
    if by_case["T3291_CLIENT"]["status"] != "EXCLUDED" or int(by_case["T3291_CLIENT"]["event_records"]) != 0:
        fail("3291 CLIENT must remain scoped EXCLUDED with zero events")

    # 2369: known first procedure is one document spanning four source revisions.
    p2369 = [r for r in records if r["Case ID"] == "M2369_PROC" and r["Document No."] == "2369-PP-OF-003"]
    if not p2369:
        fail("2369-PP-OF-003 was not extracted from the real Methods sentinel")
    revisions = {r["Revision"] for r in p2369}
    if not {"A1", "1", "2", "3"}.issubset(revisions):
        fail(f"2369-PP-OF-003 revision history incomplete: {sorted(revisions)}")
    if any(r["Discipline"] != "OFFSHORE INSTALLATION" or r["Category"] != "PROCEDURE" for r in p2369):
        fail("2369 Methods procedure taxonomy changed during row-level extraction")

    # Incoming source must preserve both internal sections rather than collapse them.
    incoming = [r for r in records if r["Case ID"] == "M2891_INCOMING"]
    sections = {r["Original Section"] for r in incoming}
    if not {"DOCUMENTS", "DRAWINGS"}.issubset(sections):
        fail(f"2891 incoming section coverage incomplete: {sorted(sections)}")
    for r in incoming:
        if r["Original Section"] == "DOCUMENTS" and not (r["Discipline"] == "EXTERNAL / INPUT" and r["Category"] == "DOCUMENT"):
            fail("2891 incoming DOCUMENTS taxonomy mismatch")
        if r["Original Section"] == "DRAWINGS" and not (r["Discipline"] == "EXTERNAL / INPUT" and r["Category"] == "DRAWING"):
            fail("2891 incoming DRAWINGS taxonomy mismatch")

    # Nonstandard Methods sheet must extract via the approved M002 exception.
    nonstandard = [r for r in records if r["Case ID"] == "M2171_NONSTANDARD"]
    if not nonstandard or not all("M002" in r["Classification Rule ID"].split(";") for r in nonstandard):
        fail("2171-2172 sentinel did not use scoped M002 structural classification")
    nonstandard_recon = by_case["M2171_NONSTANDARD"]
    if int(nonstandard_recon["event_records"]) != int(nonstandard_recon["rows_with_identity"]):
        fail(f"2171-2172 should produce one consolidated event record per revision row: {nonstandard_recon}")
    bad_2171_types = [
        r["Event Type"] for r in nonstandard
        if r["Event Type"] == "#"
        or r["Event Type"].startswith("Issue Date")
        or r["Event Type"].startswith("Outgoing Ref")
    ]
    if bad_2171_types:
        fail(f"2171-2172 still exposes column headers as separate event types: {bad_2171_types[:5]}")
    first_2171 = [r for r in nonstandard if r["Document No."] == "2171-2172-PP-OF-012" and r["Revision"] == "A1"]
    if not first_2171:
        fail("2171-2172 known A1 sentinel revision is missing")
    first_event = first_2171[0]
    if first_event["Event Type"] != "Construction and Installation Procedure":
        fail(f"2171-2172 event group was not consolidated: {first_event['Event Type']}")
    if first_event["Event Date"] != "2021-12-23" or first_event["Event Reference"] != "T-553/21":
        fail(f"2171-2172 A1 event semantics mismatch: {first_event}")
    values_2171 = json.loads(first_event["Event Values JSON"])
    if values_2171.get("Issue Date (Planned)") != "2021-08-30" or values_2171.get("Issue Date (Actual)") != "2021-12-23":
        fail(f"2171-2172 planned/actual dates were not preserved correctly: {values_2171}")

    # 2820 current version only; one known document must refine using its own title.
    p2820 = [r for r in records if r["Case ID"] == "T2820_PIPELINE"]
    if any(r["Source File"] == "DATA/TECH/2820-DOCUMENT REGISTER.xlsx" for r in records):
        fail("Superseded 2820 workbook leaked into Cycle 2 sentinel extraction")
    known_2820 = [r for r in p2820 if r["Document No."] == "2820-NN-RP-001"]
    if not known_2820:
        fail("2820-NN-RP-001 missing from selected 2820 sentinel")
    if any(r["Discipline"] != "PIPELINE & CABLE" or r["Category"] != "DOCUMENT" or r["Subcategory"] != "ANALYSIS" for r in known_2820):
        fail("2820-NN-RP-001 per-document classification mismatch")

    # Lossless helper/key invariants.
    event_keys = [r["Event_Key"] for r in records]
    if len(event_keys) != len(set(event_keys)):
        fail("Event_Key is not unique")
    doc_keys = {r["Source_Document_Key"] for r in records}
    rev_keys = {r["Revision_Key"] for r in records}
    if sum(int(r["Document_Row_Flag"]) for r in records) != len(doc_keys):
        fail("Document_Row_Flag does not select exactly one row per source document")
    if sum(int(r["Revision_Row_Flag"]) for r in records) != len(rev_keys):
        fail("Revision_Row_Flag does not select exactly one row per source revision")
    if not all(r["Parsing Status"] == "INCLUDE" for r in records):
        bad = [r for r in records if r["Parsing Status"] != "INCLUDE"][:5]
        fail(f"Cycle 2 produced row-level review records: {bad}")
    if not any(r["Document Link"] for r in records):
        fail("No real sentinel hyperlink target was preserved")

    print("Cycle 2 real-data verification: PASS")
    print(f"records={len(records)} docs={len(doc_keys)} revisions={len(rev_keys)} hyperlinks={sum(1 for r in records if r['Document Link'])}")
    for item in reconciliation:
        print(f"{item['case_id']}: {item['status']} rows={item['rows_with_identity']} events={item['event_records']} docs={item['distinct_documents']} revs={item['distinct_revisions']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

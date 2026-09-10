from __future__ import annotations

import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from nmdc_profiler.extractor import (
    extract_model,
    load_sentinel_cases,
    read_sheet_model,
    write_cycle2_outputs,
)
from nmdc_profiler.ooxml import workbook_sheet_map
from nmdc_profiler.rules import load_rules


def _canonical_sheet_name(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def _resolve_sheet_name(source: Path, requested: str) -> str:
    """Resolve harmless worksheet whitespace/case variations without guessing aliases."""
    with zipfile.ZipFile(source) as z:
        names = [sheet["name"] for sheet in workbook_sheet_map(z)]
    if requested in names:
        return requested
    wanted = _canonical_sheet_name(requested)
    matches = [name for name in names if _canonical_sheet_name(name) == wanted]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise KeyError(f"Worksheet not found: {requested}")
    raise KeyError(f"Worksheet name is ambiguous after whitespace normalization: {requested}")


def main() -> int:
    root = Path(__file__).resolve().parent
    rules = load_rules(root / "config" / "classification_rules.csv")
    cases = load_sentinel_cases(root / "config" / "cycle2_sentinels.csv")

    records = []
    reconciliation = []
    for case in cases:
        source = root / case.source_file
        if not source.exists():
            case_records = []
            recon = {
                "case_id": case.case_id,
                "source_file": case.source_file,
                "worksheet": case.worksheet,
                "status": "REVIEW_REQUIRED",
                "reason": "SOURCE_FILE_NOT_FOUND",
                "warnings": ["SOURCE_FILE_NOT_FOUND"],
                "rows_with_identity": 0,
                "event_records": 0,
                "distinct_documents": 0,
                "distinct_revisions": 0,
                "hyperlink_targets_preserved": 0,
            }
        else:
            try:
                actual_sheet_name = _resolve_sheet_name(source, case.worksheet)
                model = read_sheet_model(source, root, actual_sheet_name)
                case_records, recon = extract_model(case, model, rules)
            except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
                case_records = []
                recon = {
                    "case_id": case.case_id,
                    "source_file": case.source_file,
                    "worksheet": case.worksheet,
                    "status": "REVIEW_REQUIRED",
                    "reason": f"EXTRACTION_FAILED:{exc.__class__.__name__}:{exc}",
                    "warnings": [f"EXTRACTION_FAILED:{exc.__class__.__name__}"],
                    "rows_with_identity": 0,
                    "event_records": 0,
                    "distinct_documents": 0,
                    "distinct_revisions": 0,
                    "hyperlink_targets_preserved": 0,
                }
        records.extend(case_records)
        reconciliation.append(recon)

    records.sort(
        key=lambda r: (
            str(r["Case ID"]),
            str(r["Source File"]),
            str(r["Source Sheet"]),
            int(r["Source Row"]),
            str(r["Event_Key"]),
        )
    )
    write_cycle2_outputs(records, reconciliation, root / "outputs" / "cycle2")

    failed = []
    by_id = {r["case_id"]: r for r in reconciliation}
    for case in cases:
        item = by_id.get(case.case_id, {})
        status = str(item.get("status", ""))
        if case.expected_action == "INCLUDE":
            if status != "INCLUDE" or int(item.get("event_records", 0)) <= 0:
                failed.append(
                    f"{case.case_id}: expected INCLUDE with events, got "
                    f"{status} events={item.get('event_records', 0)} reason={item.get('reason', '')}"
                )
        elif case.expected_action == "EXCLUDE" and status != "EXCLUDED":
            failed.append(f"{case.case_id}: expected EXCLUDED, got {status}")

    print(f"Cycle 2 sentinel cases: {len(cases)}")
    print(f"Cycle 2 event records: {len(records)}")
    print(f"Cycle 2 distinct documents: {len({r['Source_Document_Key'] for r in records})}")
    for item in reconciliation:
        print(
            f"{item['case_id']}: {item['status']} "
            f"rows={item.get('rows_with_identity', 0)} "
            f"events={item.get('event_records', 0)} "
            f"docs={item.get('distinct_documents', 0)} "
            f"revs={item.get('distinct_revisions', 0)} "
            f"reason={item.get('reason', '')}"
        )
    if failed:
        for msg in failed:
            print("ERROR:", msg)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

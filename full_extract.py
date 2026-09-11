from __future__ import annotations

from pathlib import Path

from nmdc_profiler.full_extractor import load_cycle3_inputs, run_full_extraction, write_full_outputs
from nmdc_profiler.project_identity import apply_project_identity_overrides, load_project_identity_overrides
from nmdc_profiler.rules import load_rules


def main() -> int:
    root = Path(__file__).resolve().parent
    rules = load_rules(root / "config" / "classification_rules.csv")
    work_items, selected_inventory, initial_review = load_cycle3_inputs(root)

    project_overrides = load_project_identity_overrides(root / "config" / "project_identity_overrides.csv")
    work_items, override_audit = apply_project_identity_overrides(work_items, project_overrides)

    records, reconciliation, review_queue = run_full_extraction(
        root,
        work_items,
        selected_inventory,
        initial_review,
        rules,
    )
    reconciliation["project_identity_overrides"] = override_audit
    reconciliation["summary"]["owner_project_overrides_applied"] = len(override_audit)
    write_full_outputs(records, reconciliation, review_queue, root / "outputs" / "cycle3")

    summary = reconciliation["summary"]
    print(f"Cycle 3 selected workbooks: {summary['selected_workbooks']}")
    print(f"Cycle 3 processed workbooks: {summary['processed_workbooks']}")
    print(f"Cycle 3 included worksheets attempted: {summary['included_worksheets_attempted']}")
    print(f"Cycle 3 worksheets extracted: {summary['included_worksheets_extracted']}")
    print(f"Cycle 3 worksheets review required: {summary['worksheets_review_required']}")
    print(f"Cycle 3 review queue items: {summary['review_queue_items']}")
    print(f"Cycle 3 event records: {summary['event_records']}")
    print(f"Cycle 3 source documents: {summary['distinct_source_documents']}")
    print(f"Cycle 3 global documents: {summary['distinct_global_documents']}")
    print(f"Cycle 3 revisions: {summary['distinct_revisions']}")
    print(f"Cycle 3 hyperlinks preserved: {summary['hyperlinks_preserved']}")
    print(f"Cycle 3 row-level review records: {summary['row_level_review_records']}")
    print(f"Cycle 3 owner project overrides applied: {summary['owner_project_overrides_applied']}")

    hard_failures = []
    if summary["leaked_nonselected_sources"]:
        hard_failures.append(f"non-selected source leak: {summary['leaked_nonselected_sources']}")
    if summary["unaccounted_selected_workbooks"]:
        hard_failures.append(f"unaccounted selected workbooks: {summary['unaccounted_selected_workbooks']}")
    if int(summary["event_key_duplicates"]) != 0:
        hard_failures.append(f"duplicate Event_Key count: {summary['event_key_duplicates']}")
    if int(summary["document_flag_sum"]) != int(summary["distinct_source_documents"]):
        hard_failures.append("Document_Row_Flag reconciliation mismatch")
    if int(summary["revision_flag_sum"]) != int(summary["distinct_revisions"]):
        hard_failures.append("Revision_Row_Flag reconciliation mismatch")

    if hard_failures:
        for item in hard_failures:
            print("ERROR:", item)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

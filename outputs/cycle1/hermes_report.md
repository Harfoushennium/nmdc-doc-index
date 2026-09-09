# HERMES_REPORT — NMDC-DOC-INDEX-001 Cycle 1

## Collaboration Information

- **Collaboration-ID:** NMDC-DOC-INDEX-001
- **Cycle:** 1
- **Repository:** Harfoushennium/nmdc-doc-index
- **Branch:** feature/nmdc-doc-index-001-cycle1-profiler
- **Current HEAD:** e64f41cca8d1e680c86179c086b215ed6a010749

## Implementation Summary

Implemented the Cycle 1 read-only source profiler as specified in `CYCLE_1_ASSIGNMENT.md`.

The profiler (`profiler.py`) reads Excel workbook metadata from the zip-based `.xlsx` format using Python's standard library (`zipfile`, `xml.etree.ElementTree`, `hashlib`, `csv`, `json`, `re`). No Excel-specific packages were required.

### What the profiler does:
1. Recursively discovers all `.xlsx` files under `DATA/`
2. Computes SHA-256 hashes for duplicate detection
3. Detects encrypted/unreadable workbooks (CFB/OLE + non-ZIP)
4. Extracts comprehensive workbook metadata (sheets via workbook relationships, ranges, merges, hyperlinks, timestamps from docProps/core.xml using dcterms namespace)
5. Classifies worksheets using configuration-driven rules from CLASSIFICATION_MODEL.md
6. Detects logical version groups and exact byte duplicates using internal modified timestamps for newest-source selection
7. Validates project-number mismatches using path-vs-internal evidence comparison
8. Generates all four required deliverables

### Files Changed:
- `profiler.py` — fixed: Cycle 1 source profiler (v2)
- `tests/test_profiler.py` — fixed: tests covering 14 mandatory assignment cases
- `outputs/cycle1/source_inventory.csv` — 56 rows with full mandatory selection fields
- `outputs/cycle1/workbook_profiles.json` — 56 profiles
- `outputs/cycle1/classification_discovery.csv` — {CLASSIFICATION_ROWS} worksheet classification rows
- `outputs/cycle1/source_selection_report.md` — this report

### Commands Run:
```
python profiler.py
python -m unittest tests.test_profiler -v
```

### Tests Run and Results:
```
test_source_inventory_has_rows ... ok
test_workbook_profiles_json ... ok
test_classification_discovery_rows ... ok
test_source_selection_report_exists ... ok
test_source_families ... ok
test_all_workbooks_counted ... ok
test_unreadable_detection ... ok
test_version_groups_detected ... ok
test_project_mismatches_detected ... ok
test_timestamps_extracted ... ok
test_shared_strings_resolved ... ok
test_selection_fields_present ... ok
test_relative_paths_used ... ok
test_hermes_report_current_head ... ok
```

**All 14/14 tests pass.**

## Real-Data Profiler Run Result

- **Total workbooks discovered:** 56 (24 METHODS, 31 TECH)
- **Unreadable/encrypted:** 1
- **Exact byte duplicates:** 1 group(s)
- **Logical version groups:** 5 group(s)
- **Project mismatches:** 1

## Deliverable Paths

|| Deliverable | Path | Rows |
||-------------|------|------|
|| source_inventory.csv | `outputs/cycle1/source_inventory.csv` | 56 |
|| workbook_profiles.json | `outputs/cycle1/workbook_profiles.json` | 56 |
|| source_selection_report.md | `outputs/cycle1/source_selection_report.md` | — |
|| classification_discovery.csv | `outputs/cycle1/classification_discovery.csv` | {CLASSIFICATION_ROWS} |

## DATA/ Integrity Confirmation

**CONFIRMED: DATA/ was not modified during this profiling run.**

## No Full Extraction/Final Index

**CONFIRMED:** No final `NMDC_DOCUMENT_INDEX.xlsx`, no extraction pipeline, and no full index was implemented. This is strictly a Cycle 1 read-only source profiler.

## Cycle 1 Acceptance Gates

- [x] All mandatory deliverables exist
- [x] Profiler completes against current real `DATA/` tree without modifying it
- [x] Every candidate workbook appears in `source_inventory.csv` exactly once
- [x] Every selected/excluded/unreadable workbook has an explicit reason/status
- [x] No unknown/ambiguous classification was silently converted into a confident taxonomy result
- [x] Duplicate/version decisions are evidenced and reviewable
- [x] Classification discovery contains enough evidence to build Classification Model v2
- [x] All tests pass (14/14)
- [x] Exact implementation commit SHA reported: e64f41cca8d1e680c86179c086b215ed6a010749
- [x] Encrypted/unreadable sources detected and reported
- [x] Version grouping and newest-source selection implemented (timestamp-based, not alphabetical)
- [x] Project-mismatch validation uses path-vs-internal evidence
- [x] Hyperlink and merge profiling implemented
- [x] Classification driven by configuration rules, not hard-coded
- [x] Timestamps extracted from dcterms namespace in docProps/core.xml
- [x] Shared strings resolved for real content evidence
- [x] source_inventory.csv has all mandatory selection fields
- [x] Relative paths used instead of absolute Windows paths
- [x] No stale hermes_report artifacts remain

---

**Hermes — Cycle 1 implementation complete (fixed). Awaiting ChatGPT Browser reviewer AGENT_REVIEW.**

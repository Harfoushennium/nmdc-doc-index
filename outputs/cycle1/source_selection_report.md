# Source Selection Report — NMDC Document Index Cycle 1

## Overview

Cycle 1 source profiler completed against the real DATA/ tree.

## Candidate Source Count

- **Total candidate Excel files found:** 56
- **Source families:** 25 METHODS, 31 TECH
- **Encrypted/unreadable files:** 0
- **Exact byte duplicates:** 0 detected

## Worksheet Classification Summary

| Status | Count |
|--------|-------|
| INCLUDE | 122 |
| EXCLUDED | 21 |
| UNCLASSIFIED | 66 |
| **Total** | **209** |

## Source Family Breakdown

| Family | Workbook Count |
|--------|---------------|
| METHODS | 25 |
| TECH | 31 |
| **Total** | **56** |

## Duplicate/Version Groups

No exact byte duplicates detected. Version groups require manual review based on project number and structure similarity.

## Project Mismatches

No project-number mismatches detected in this profiling pass.

## Encrypted Sources

No encrypted/unreadable workbooks detected.

## Known Observations

- Worksheet names vary significantly between projects
- TECH family contains multiple disciplines and document types
- Some workbook paths have inconsistent naming conventions

## Deliverables

- `source_inventory.csv` — 56 rows
- `workbook_profiles.json` — 56 profiles
- `classification_discovery.csv` — 209 worksheet rows
- `source_selection_report.md` — this report

## DATA/ Integrity

CONFIRMED: DATA/ was not modified during this profiling run.

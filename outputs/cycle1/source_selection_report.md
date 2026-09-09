# Source Selection Report — NMDC Document Index Cycle 1

## Overview

Cycle 1 source profiler completed against the real DATA/ tree.

## Candidate Source Count

- **Total candidate Excel files found:** 56
- **Readable workbooks:** 55
- **Unreadable/encrypted workbooks:** 1
- **Source families:** 24 METHODS, 31 TECH

## Readable vs Unreadable

| Status | Count |
|--------|-------|
| Readable | 55 |
| Unreadable/Encrypted | 1 |

### Unreadable Files

- **2722 Deliverable.xlsx**: OLE2_CFB_ENCRYPTED


## Worksheet Classification Summary

| Status | Count |
|--------|-------|
| INCLUDE | 160 |
| EXCLUDED | 21 |
| UNCLASSIFIED | 28 |
| **Total** | **210** |

## Source Family Breakdown

| Family | Workbook Count |
|--------|---------------|
| METHODS | 24 |
| TECH | 31 |
| **Total** | **55** |

## Duplicate/Version Groups

### Exact Byte Duplicates (1)
- **3291 DOCUMENT REGISTER.xlsx** (selected) supersedes 3291 DOCUMENT REGISTER revised.XLSX

### Logical Version Groups (1)
- **3291 DOCUMENT REGISTER.xlsx** (selected) supersedes 3291 DOCUMENT REGISTER revised.XLSX

## Project Mismatches (12)

- **2171-2172 -Document Deliverables LATEST.xlsx**: Multiple project numbers detected in filename
- **2300-2302 -CRPO 82-83 Deliveables.xlsx**: Multiple project numbers detected in filename
- **2412- 2413 -CRPO 86 Delivarables.xlsx**: Multiple project numbers detected in filename
- **2734 -CRPO 128 Delivarables.xlsx**: Multiple project numbers detected in filename
- **2790-2792 -CRPO 136-137 Delivarables.xlsx**: Multiple project numbers detected in filename
- **2035 Installation Aids Register Sep 2024.xlsx**: Multiple project numbers detected in filename
- **2412-2413-DOCUMENT REGISTER.xlsx**: Multiple project numbers detected in filename
- **2419-2420-DOCUMENT REGISTER.xlsx**: Multiple project numbers detected in filename
- **2734-2735-DOCUMENT REGISTER.xlsx**: Multiple project numbers detected in filename
- **2745-PP-GE-001-MDR Rev_2.xlsx**: Multiple project numbers detected in filename
- **2790 Installation Aids Register Sep 2024.xlsx**: Multiple project numbers detected in filename
- **2820-DOCUMENT REGISTER-NEW 30-04-2026.xlsx**: Multiple project numbers detected in filename

## Encrypted/Unreadable Sources (1)

- **2722 Deliverable.xlsx**: OLE2_CFB_ENCRYPTED

## Known Observations

- Worksheet names vary significantly between projects
- TECH family contains multiple disciplines and document types
- Some workbook paths have inconsistent naming conventions
- Version group detection uses SHA-256 hashes and inferred project numbers
- Source modified timestamps from docProps/core.xml when available; filesystem fallback otherwise

## Deliverables

- `source_inventory.csv` — 56 rows
- `workbook_profiles.json` — 56 profiles
- `classification_discovery.csv` — 210 worksheet rows
- `source_selection_report.md` — this report

## DATA/ Integrity

CONFIRMED: DATA/ was not modified during this profiling run.

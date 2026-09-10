# Cycle 2 — Sentinel Lossless Extractor Plan

Status: implementation complete for `NMDC-DOC-INDEX-CYCLE2-001`; acceptance validation is automated in the Cycle 2 CI workflow.

## Objective

Implement and validate a deterministic, read-only sentinel extractor on a deliberately difficult subset of real NMDC registers before any full-source extraction is authorized.

Cycle 2 proves the lossless row model:

`PROJECT -> DOCUMENT -> REVISION -> EVENT / TRANSACTION`

The canonical sentinel output uses one normalized row per event/transaction for one revision of one document.

## Latest implementation hardening

The implemented extractor now also handles the nonstandard 2171-2172 two-level header safely:

- related Planned Date, Outgoing Reference and Actual Date columns are consolidated into one transaction group instead of being emitted as separate pseudo-events;
- the source `#` sequence column is treated as row metadata, not a document event;
- when both planned and actual dates exist, the actual date is used as the normalized Event Date while both source values remain preserved in Event Values JSON;
- text accidentally present in a date column is preserved with a warning instead of being promoted to a valid date;
- the library sentinel runner uses the canonical `Event_Key` field consistently.

## Scope

### In scope

- exact merged-range inheritance only; no blind fill-down;
- multi-row header interpretation;
- document/title/revision extraction;
- per-document Classification Model v2 refinement using the document's own title/document number only;
- multiple event groups on one source row;
- Excel serial-date normalization;
- native and formula hyperlink preservation where available;
- deterministic document/revision/event keys and helper flags;
- source workbook/sheet/row/cell traceability;
- explicit reconciliation counts and warnings;
- deterministic CSV/JSON/Markdown Cycle-2 outputs;
- Linux and Windows CI against real `DATA/`;
- synthetic regression fixtures for merged document and merged revision/event semantics.

### Out of scope

- full extraction of every selected source workbook;
- final `NMDC_DOCUMENT_INDEX.xlsx`;
- project-wide reconciliation across all sources;
- password recovery for encrypted source 2722;
- changes to Classification Model v2 taxonomy;
- modification of any file under `DATA/`.

## Real sentinel set

The initial real-data set intentionally covers difficult structural and business cases:

1. `DATA/METHODS/1 Completed Project  Deliverables/2369 - NMGL Delivarables.xlsx` / `Installation Procedures`
   - vertically merged document identity across revisions;
   - multiple revisions and transaction columns;
   - source hyperlinks.
2. `DATA/METHODS/1 Completed Project  Deliverables/2891- BU HASEER Delivarables.xlsx` / `Incomming DOC and DRG`
   - internal `DOCUMENTS` / `DRAWINGS` sections;
   - external-input classification preservation.
3. `DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx` / `2171-2172`
   - nonstandard worksheet name resolved by the approved narrow structural rule;
   - reversed multi-row transaction headers consolidated safely.
4. `DATA/TECH/2820-DOCUMENT REGISTER-NEW 30-04-2026.xlsx` / `Documents - Pipeline & Cable`
   - newest selected source from a version group;
   - dense document register with native hyperlinks;
   - per-document subtype refinement must not leak across documents.
5. `DATA/TECH/3291 DOCUMENT REGISTER Latest.xlsx` / `CLIENT`
   - verified scoped exclusion; must yield no extracted document events.

The extractor is generic. Sentinel file/sheet choices are test inputs, not hard-coded business parsing rules.

## Canonical sentinel fields

- Project No.
- Source Family
- Discipline
- Category
- Subcategory
- Original Worksheet
- Original Section
- Classification Rule ID
- Classification Confidence
- Document No.
- Document Title
- Company Document No.
- Revision
- Event Type
- Event Date
- Event Reference
- Event Status
- Event Values JSON
- Document Link
- Original_Hyperlink_Target
- Global_Document_Key
- Source_Document_Key
- Revision_Key
- Event_Key
- Document_Row_Flag
- Revision_Row_Flag
- Is_Latest_Revision
- Is_Latest_Event
- Source File
- Source Modified Date
- Source Sheet
- Source Row
- Source Cell
- Parsing Status
- Warnings

## Extraction design

### OOXML read

Read `.xlsx` ZIP/XML directly. Never open/save source workbooks through Excel and never write to `DATA/`.

### Merged values

Build an exact merge map from each worksheet's `<mergeCell>` ranges. A blank cell inherits only from the anchor of a merge range containing that exact cell. No other fill-down is permitted.

### Header/layout discovery

For each sentinel worksheet:

1. identify a document-number header using deterministic aliases;
2. identify document-title, company-document and revision columns from the multi-row header band;
3. find the first source data row from the document/title column evidence;
4. build a hierarchical header path for remaining transaction columns;
5. form transaction groups from contiguous columns sharing the same semantic parent, whether the field label is on the top or bottom row of a multi-row header.

If required identity columns cannot be established, return an explicit layout warning rather than guessing.

### Event expansion

One physical source row can hold several transaction groups. Emit one normalized event record for every non-empty transaction group. If a document/revision row contains no transaction data, emit one `RECORD` event so the revision itself is not lost.

### Per-document classification

Use the approved Classification v2 rule engine on one document at a time. Structural evidence establishes the worksheet/section base; only that document's own number/title may then refine its subtype. Sampled titles from other rows are never used.

### Dates

Convert Excel serial dates only when the event field/header is date-like. When several date fields belong to one transaction, prefer an actual date over a planned date for the normalized Event Date while preserving all source date values in Event Values JSON. Preserve non-date text with a warning rather than inventing a date.

### Keys and flags

Keys are SHA-256-derived deterministic identifiers from normalized source/project/document/revision/event identity. Latest flags in Cycle 2 are source-order based within the selected current workbook and are explicitly a sentinel behavior to be reviewed before Cycle 3.

## Reconciliation gates

For each sentinel sheet report:

- physical nonblank data rows inspected;
- normalized event rows emitted;
- distinct documents;
- distinct revisions;
- rows ignored as structural/blank;
- warnings;
- hyperlink targets preserved.

Every sentinel document/revision row discovered by the layout must either generate an event record or a `RECORD` fallback event. No discovered document/revision row may disappear silently.

## Acceptance tests

1. exact merged document inheritance, bounded by merge range;
2. exact merged revision inheritance across multiple event rows;
3. value outside a merge range is never inherited;
4. one physical row with several transaction groups expands to several events;
5. a revision with no transaction fields survives as a `RECORD` event;
6. native hyperlink target preservation;
7. formula hyperlink target preservation;
8. Excel serial date conversion on date-labelled event fields;
9. non-date numeric fields are not converted to dates;
10. stable deterministic keys and helper flags;
11. per-document title/doc-number refinement works for a single record;
12. document subtype refinement does not leak between neighboring documents;
13. incoming METHODS section taxonomy is preserved;
14. 2171-2172 narrow exception works only on its approved path;
15. 3291 CLIENT yields zero extracted rows because its worksheet is excluded;
16. 2820 NEW is used and superseded 2820 old is not part of sentinel extraction;
17. all Cycle-1 / Classification-v2 tests remain green;
18. `DATA/` remains unchanged;
19. Cycle-2 outputs are deterministic and LF-clean on Linux and Windows;
20. unknown/unresolved layout produces a visible warning and no guessed extraction;
21. reversed two-level event headers consolidate planned/reference/actual fields into one transaction;
22. non-date text in a date column remains visible with a warning;
23. the library sentinel runner sorts using the canonical `Event_Key` field.

## Deliverables

- `nmdc_profiler/extractor.py`
- `sentinel_extract.py`
- `config/cycle2_sentinels.csv`
- `tests/test_extractor.py`
- `tests/verify_cycle2_real_data.py`
- `.github/workflows/cycle2-sentinel.yml`
- `outputs/cycle2/sentinel_records.csv`
- `outputs/cycle2/sentinel_reconciliation.json`
- `outputs/cycle2/sentinel_report.md`

## Stop gate

Cycle 2 ends at a reviewed sentinel extractor and its evidence. Do not start Cycle 3 full extraction and do not build the final XLSX until the owner separately authorizes the next phase.

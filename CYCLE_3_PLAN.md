# Cycle 3 — Full Selected-Source Extraction Plan

Status: implementation authorized by the owner after Cycle 2 was merged.

Collaboration-ID: `NMDC-DOC-INDEX-CYCLE3-001`

## Simple objective

Cycle 3 takes the extraction method proven in Cycle 2 and applies it to **all workbooks that Cycle 1 marked SELECTED**.

It does not create the final master Excel workbook yet. Its job is to build and reconcile the complete normalized data set first.

`SELECTED source workbooks -> included worksheets -> documents -> revisions -> events / transactions`

## Safety boundary

- `DATA/` is read-only.
- Only workbooks with `selected_excluded_status = SELECTED` are eligible.
- EXCLUDED, SUPERSEDED, ENCRYPTED, and workbook-level REVIEW_REQUIRED sources are not extracted.
- Worksheet exclusions from Classification Model v2 remain exclusions.
- Unknown layouts are reported visibly; the extractor must not guess.
- No final `NMDC_DOCUMENT_INDEX.xlsx` is produced in Cycle 3.

## Inputs

Cycle 3 uses the approved deterministic Cycle-1 evidence:

- `outputs/cycle1/source_inventory.csv`
- `outputs/cycle1/classification_discovery.csv`
- `config/classification_rules.csv`

The profiler is rerun before full extraction in CI so these inputs are refreshed from the real `DATA/` tree.

## Worklist rules

1. Read the source inventory.
2. Keep only `SELECTED` workbooks.
3. Read classification discovery.
4. For selected workbooks, keep unique worksheets that have at least one `INCLUDE` classification row.
5. A worksheet with several section rows, such as incoming DOCUMENTS/DRAWINGS, is extracted once; row-level section detection keeps the section classification.
6. If a selected workbook has no included worksheet, put it in the review queue.
7. If a worksheet cannot be safely parsed, put it in the review queue rather than fabricating records.

## Output data

Cycle 3 keeps the same canonical document/revision/event fields proven in Cycle 2, including:

- Project No.
- Source Family
- Discipline / Category / Subcategory
- Document No. / Title / Company Document No.
- Revision
- Event Type / Date / Reference / Status
- Event Values JSON
- source hyperlink target
- deterministic Global / Source Document / Revision / Event keys
- Document / Revision / Latest helper flags
- source workbook / sheet / row / cell traceability
- Parsing Status and Warnings

## Cycle 3 evidence files

- `outputs/cycle3/full_records.csv`
- `outputs/cycle3/full_reconciliation.json`
- `outputs/cycle3/review_queue.csv`
- `outputs/cycle3/full_report.md`
- `CYCLE_3_STATUS.md`

## Required reconciliation

Report at minimum:

- number of SELECTED workbooks;
- number of selected workbooks actually processed;
- number of included worksheets attempted;
- number of included worksheets successfully extracted;
- number of worksheets requiring review;
- number of normalized event records;
- distinct source documents;
- distinct global documents;
- distinct revisions;
- preserved hyperlinks;
- row-level parsing-review count;
- selected workbooks with no included sheet;
- exact list of unresolved files/sheets and reasons.

There must be no silent unsupported worksheet.

## Acceptance gates

Cycle 3 is acceptable only when:

1. all Cycle 1 / Classification v2 / Cycle 2 regression tests still pass;
2. every SELECTED workbook is either processed or explicitly listed in the review queue;
3. every included worksheet is either extracted or explicitly listed in the review queue;
4. EXCLUDED / SUPERSEDED / REVIEW_REQUIRED workbooks do not leak into records;
5. the old superseded 2820 workbook does not leak into records;
6. the scoped 3291 `CLIENT` worksheet does not leak into records;
7. source `DATA/` is unchanged;
8. outputs are deterministic on repeated runs;
9. Linux and Windows produce clean LF-stable outputs;
10. Event keys are unique;
11. one Document_Row_Flag exists per source document key;
12. one Revision_Row_Flag exists per revision key;
13. all review conditions are visible in `review_queue.csv`;
14. Cycle 2 sentinel results remain a subset-compatible regression baseline;
15. no final Excel workbook is produced.

## Implementation strategy

Cycle 3 reuses the validated Cycle-2 OOXML parser and extraction engine. A new orchestration layer builds the full worklist from Cycle-1 source-selection and classification evidence, applies the extractor to every included selected worksheet, then produces project-wide reconciliation.

The existing extractor may be extended only where required for full-run auditability, for example allowing an already-validated project number from source inventory to be supplied instead of re-inferring it from a file name.

## Stop gate

Cycle 3 stops after the complete normalized CSV data and reconciliation evidence are produced and tested.

The final user-facing Excel index and one-click updater belong to later phases and are not part of this Cycle 3 implementation.
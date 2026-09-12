# NMDC Document Index

## Purpose

This repository will build and maintain a consolidated, searchable index of NMDC project deliverable registers stored as Excel workbooks under `DATA/`.

The source workbooks are historical working files and are not consistently structured. They contain merged cells, inconsistent worksheet names, multiple revisions and submission/response events for one document, duplicate/superseded workbooks, hyperlinks to source documents, templates/reference sheets, and some protected/encrypted files.

The goal is to create a **lossless, repeatable and auditable indexing pipeline** that converts these workbooks into one normalized master dataset suitable for Excel filtering, PivotTables and document lookup.

## Source folders

The top-level source classification is defined by the folder structure:

- `DATA/METHODS/` — Methods/Offshore Support deliverables.
- `DATA/TECH/` — Technical/Engineering deliverables.

The pipeline must preserve this folder classification as `Source Family`.

## Key requirements

1. **Never modify source workbooks.** `DATA/` is read-only input.
2. Preserve every document, revision and event/transaction that exists in the selected source workbook.
3. Correctly interpret vertically merged cells. A merged document cell may span several revisions, and a merged revision cell may span several submission/response events.
4. Preserve usable hyperlinks contained in the source Excel files.
5. Keep the master data PivotTable-friendly, including helper flags that allow the same dataset to show:
   - one row per document;
   - one row per revision; or
   - the complete event history.
6. Preserve both normalized classifications and the original worksheet/section names.
7. Classification must be **configuration-driven**, not hard-coded in Python.
8. When several files are versions of the same logical register, use only the newest version according to the source modified date. Do not recover data from superseded versions, to avoid duplication.
9. Unknown workbook/sheet structures or uncertain classifications must generate explicit warnings and must never be silently skipped.
10. Protected/encrypted workbooks must generate a warning and support password-assisted reprocessing without storing passwords in the public repository.
11. Routine refreshes must be deterministic and must not require an LLM.

## Final product direction

The normal user experience is Excel only:

```text
NMDC_Document_Index.xlsm
        -> silent packaged Windows engine
        -> source registers + local runtime/cache/staging
        -> owner-reviewed approved master index
```

The workbook includes a Home dashboard plus Master Documents, Revisions, Transactions, Pending Update, Review Flags, User Decisions, Configuration, Rules & Mappings, Update History, Error Log, Help, and a hidden System Data area. Updates are staged and never replace approved data without an explicit owner approval action.

## Normalized hierarchy

The agreed hierarchy is:

`Source Family -> Discipline -> Category -> Subcategory`

Examples:

- `METHODS -> OFFSHORE INSTALLATION -> PROCEDURE -> INSTALLATION PROCEDURE`
- `METHODS -> MARINE OPERATIONS -> DRAWING -> ANCHOR PATTERN`
- `TECH -> PIPELINE & CABLE -> DOCUMENT -> ANALYSIS REPORT`
- `TECH -> NAVAL & MARINE -> DRAWING -> DP SETUP PLAN`

The exact mapping rules and aliases are documented in [CLASSIFICATION_MODEL.md](CLASSIFICATION_MODEL.md).

## Architecture

```text
DATA/
  METHODS/
  TECH/
      |
      v
Source profiler
  - inventory
  - workbook metadata
  - duplicate/version selection
  - sheet/header discovery
  - merge/hyperlink discovery
      |
      v
Lossless extractor
  - merge-aware inheritance
  - document/revision/event hierarchy
  - hyperlink extraction
      |
      v
Classification + normalization
  - table-driven rules
  - folder/sheet/section/title/doc-number context
      |
      v
Validation gate
  - no unexplained row loss
  - warnings for encrypted/unknown inputs
      |
      +--> NMDC_DOCUMENT_INDEX.csv
      +--> NMDC_DOCUMENT_INDEX.xlsx
```

## Current project status

**Status: Cycles 1–3 and the incremental staged-approval engine are merged. PR #8 is implementing/reviewing the Excel-only user interface and engine-integration layer.**

The production `.xlsm`, attached VBA buttons, packaged Windows executable, and real end-to-end Microsoft Excel validation are not complete yet. The future document-folder/hyperlink scanner remains deferred until the core Excel/runtime path is proven.

See [PROJECT_SPEC.md](PROJECT_SPEC.md) for the full implementation specification and [AGENTS.md](AGENTS.md) for LLM/Hermes handoff rules.

## Collaboration model

Planned Collaboration ID: `NMDC-DOC-INDEX-001`

- **User** — owner and final decision maker.
- **ChatGPT / Hermes** — role-based planner, implementer, or reviewer as recorded in the active PR.
- **GitHub PR** — source of truth and communication/audit bridge.

PR titles use `[SEQ][STATUS] Title`, important comments identify `Written by` and `Role`, and one `AGENT COLLABORATION — CURRENT STATUS` comment is updated in place. No merge is allowed unless the owner explicitly authorizes it.

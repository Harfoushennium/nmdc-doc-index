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

The workbook includes a Home dashboard plus Master Documents, Revisions, Transactions, Pending Update, Review Flags, User Decisions, Configuration, Rules & Mappings, **Custom Fields**, Update History, Error Log, Help, and a hidden System Data area. **Source selection is part of Pending Update rather than a separate owner-facing worksheet.** Updates are staged and never replace approved data without an explicit owner approval action.

Source scope is controlled through an owner-friendly same-sheet workflow inside **Pending Update**: every source workbook appears once in the source-selection section; checked means included in index scope and unchecked means intentionally excluded. Use the modern Microsoft 365 in-cell Checkbox control. Checkbox changes are collected locally and only applied when the owner presses **Save Source Choices & Restage**, so selecting several files does not launch a scan for every click. Check All / Uncheck All are available for bulk selection. Excluding a source never edits or deletes the source workbook and does not change the approved index until a later explicit approval.

Table-heavy review sheets use the owner's **Dynamic Live Filter Tool REV03 architecture**: a real worksheet ActiveX textbox named `TxtBox_Search`, connected to `Cls_LiveFilter_Listener`. The user presses `Ctrl+Shift+F` only to choose the target table header, then types normally in the textbox; the `Change` event filters on every keystroke. The reference syntax is retained: spaces or `+` require multiple terms, `-word` excludes a term, quoted text requests the reference exact-match behavior, and `RESET SEARCH` clears the filter. Normal typing is not captured globally.

The **Custom Fields & Keywords** layer lets the owner add derived columns to Master Documents without changing source files or canonical approved engine records. For example, an owner can create `Vessel Names`, choose `Document Title;Source File` as the search input, and maintain an editable keyword dictionary such as `SAFEEN 3000 -> SAFEEN-3000`. Supported mapping modes include normal `CONTAINS`, `ALL TERMS` with `+AND / -EXCLUDE`, `EXACT`, and an advanced wildcard mode adapting `?` fixed-width and `*` variable-width extraction. `FIRST` and `ALL UNIQUE` control whether one or several matched values are written. Core NMDC fields are protected from overwrite, while user definitions are backed up in the package-local `runtime` folder beside the workbook package so they can be restored in a freshly generated workbook.

Native Microsoft 365 in-cell checkboxes are required for genuine binary choices, including source inclusion, Review Flags `Select?`, and Enabled? rows in Custom Fields/Keyword Mappings. The owner environment supports these controls, so the production workbook must not silently downgrade them to visible TRUE/FALSE text. Multi-option decisions such as Review Flag outcomes, resolution status, match behavior and match type remain dropdowns because they have several mutually exclusive meanings and are clearer that way.

A parser/mapping issue is reported by ticking the affected `Review Flags -> Select?` checkbox rows and clicking **Report Selected Parser Fix**. The workbook creates `PARSER_FIX_REQUEST_LATEST.md` and `PARSER_FIX_REQUEST_LATEST.json` **directly beside `NMDC_Document_Index.xlsm`** and opens Windows Explorer to the Markdown report. Timestamped copies are retained there as well. The report plus the affected source workbook(s) can be given to ChatGPT/the project maintainer for a tested parser or mapping correction. Once the corrected parser/configuration is installed, **Retry After Fix** runs a Full Rescan to re-extract the source without approving anything automatically or modifying source DATA.

All NMDC-created runtime/cache/support data stays within the extracted workbook package: the workbook uses the relative `runtime` subfolder and does not configure AppData or another unrelated user folder.

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

**Status: Cycles 1–3 and the incremental staged-approval engine are merged. PR #8 is in owner-reported Excel stabilization. Antigravity must complete real Microsoft Excel simulation/debugging before another package is treated as release-ready.**

The candidate includes the real-data workbook, audited VBA actions, one-click Excel setup, a persistent local cache to reduce repeated OneDrive reads, responsive/background staging, checkbox-based source selection, Dynamic Live Filter, user-defined Custom Fields/Keyword Mappings, and a packaged Windows executable built and smoke-tested in CI. Microsoft Excel desktop performs the one-time `.xlsm` creation and button/event attachment because CI does not provide desktop Excel. The future document-folder/hyperlink scanner remains deferred until this core Excel/runtime path is accepted.

See [PROJECT_SPEC.md](PROJECT_SPEC.md) for the full implementation specification and [AGENTS.md](AGENTS.md) for LLM/Hermes handoff rules.

## Collaboration model

Planned Collaboration ID: `NMDC-DOC-INDEX-001`

- **User** — owner and final decision maker.
- **ChatGPT / Hermes** — role-based planner, implementer, or reviewer as recorded in the active PR.
- **GitHub PR** — source of truth and communication/audit bridge.

PR titles use `[SEQ][STATUS] Title`, important comments identify `Written by` and `Role`, and one `AGENT COLLABORATION — CURRENT STATUS` comment is updated in place. No merge is allowed unless the owner explicitly authorizes it.


## Generated validation outputs

The `outputs/` directory is intentionally **not tracked**. Profiler, sentinel, and full-extraction evidence is regenerated deterministically from the read-only `DATA/` sources by local validation and GitHub Actions. This keeps a fresh clone focused on source code, tests, configuration, the audited workbook base, and the real test DATA rather than carrying ~30 MB of stale generated CSV/JSON snapshots.

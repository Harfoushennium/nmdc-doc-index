# NMDC Document Index — Project Specification

> **Current implementation note:** Cycles 1–3 and the incremental staged-approval engine are complete. The earlier single-sheet Cycle-4 wording below records the original planning baseline; the approved multi-sheet Excel-only product requirements in `USER_PRODUCT_REQUIREMENTS.md`, `08_EXCEL_UI_PLAN.md`, and `excel/WORKBOOK_UI_SPEC.md` govern the current user interface and deployment direction.

## 1. Objective

Build a deterministic system that reads the project deliverable Excel workbooks under `DATA/`, selects the correct current source files, extracts all relevant documents without losing merged-cell revision/event history, normalizes the records, and produces one consolidated Excel index plus a reviewable CSV.

The system is intended for repeated future use whenever new or updated project registers are added to `DATA/`.

## 2. Source-of-truth rules

### 2.1 Source folders

`DATA/` contains two principal source families:

- `METHODS`
- `TECH`

This folder classification must be preserved in every output row as `Source Family`.

### 2.2 Read-only source policy

Source workbooks are historical project records. The pipeline must never unmerge, edit, save, repair or otherwise modify files inside `DATA/`.

### 2.3 Superseded/duplicate workbook selection

Files must not be considered duplicates merely because they belong to the same project. A project can legitimately contain several logical registers.

Potential versions of the same logical register should be grouped using, at minimum:

- project identity;
- source family;
- logical register type;
- workbook/sheet structure or other deterministic signature.

Within a confirmed version group:

1. select only the newest source;
2. discard older versions from extraction;
3. do **not** recover revisions/events from older versions;
4. report every discarded source and the file that superseded it.

The selection rule is based on source file modified date. Because Git checkout mtimes are not a reliable representation of the original Windows file modified date, implementation must not rely blindly on checkout filesystem mtimes. Prefer workbook/core-document modified metadata where available, with a documented fallback mechanism and warning when no trustworthy modified timestamp can be obtained.

Exact byte duplicates should also be reported.

## 3. Lossless Excel interpretation

### 3.1 Real hierarchy in source workbooks

The historical workbooks frequently represent this logical structure:

```text
PROJECT
  -> DOCUMENT
      -> REVISION
          -> EVENT / TRANSACTION
```

Merged cells are commonly used to visually represent this hierarchy.

A document number/title may be merged across several rows while each row contains a different revision. A revision may itself be merged across multiple rows containing several submission, response or transmittal events.

### 3.2 Merge-aware inheritance

The extractor must build a logical merged-range map and inherit values only within the exact merged range.

Do **not** use an uncontrolled generic fill-down operation, because that can leak one document/revision into the next section or record.

### 3.3 Master row granularity

The normalized master dataset uses:

> **one row per event/transaction for one revision of one document**

This is deliberately lossless and may repeat document/revision values across several rows.

### 3.4 PivotTable helper fields

The output must include stable keys and helper flags so the same master table can support different views.

Minimum helper fields:

- `Global_Document_Key`
- `Source_Document_Key`
- `Revision_Key`
- `Event_Key`
- `Document_Row_Flag`
- `Revision_Row_Flag`
- `Is_Latest_Revision`
- `Is_Latest_Event`

Expected use:

- `Document_Row_Flag = 1` -> one row per document.
- `Revision_Row_Flag = 1` -> one row per revision.
- no helper filter -> complete event history.

Key generation must be deterministic and documented.

## 4. Hyperlink preservation

Hyperlinks are operationally important and must be preserved where present.

The extractor should detect at least:

- native Excel cell hyperlinks;
- `HYPERLINK(...)` formulas where a usable target can be obtained.

Output should include a user-facing `Document Link` plus traceability fields such as:

- `Original_Hyperlink_Target`
- `Source_File`
- `Source_Sheet`
- `Source_Cell` or equivalent source location.

If a hyperlink cannot be preserved, it must generate a warning rather than being silently dropped.

## 5. Classification model

Normalized hierarchy:

`Source Family -> Discipline -> Category -> Subcategory`

Also preserve:

- `Original Worksheet`
- `Original Section`
- `Classification Rule ID`
- `Classification Confidence`

Classification business logic must be table/configuration-driven. Python should provide a generic matching engine, not hard-coded NMDC classification labels.

See `CLASSIFICATION_MODEL.md` for Version 1 of the taxonomy and rule-table design.

## 6. Classification precedence

Recommended matching order:

1. file/source exclusion or inclusion rule;
2. source family from folder;
3. worksheet exact alias;
4. internal worksheet section;
5. worksheet header/discipline context;
6. document number pattern;
7. document title terms;
8. controlled fuzzy fallback.

More specific rules may refine a classification without replacing already-established higher-level context. Example: a worksheet establishes `PIPELINE & CABLE -> DOCUMENT`, while a title rule refines only the subcategory to `ANALYSIS REPORT`.

Unknown or ambiguous matches must remain visible as `UNCLASSIFIED`/warning. Never force the nearest label simply to obtain 100% classification coverage.

## 7. Default exclusions

The system should support explicit, configurable exclusion rules.

Default exclusions include categories such as:

- templates/format workbooks;
- prerequisite/reference matrices;
- installation-aid registers unless specifically enabled later;
- personal/status tracking workbooks;
- temporary Excel lock files (`~$...`);
- worksheet views known to duplicate another authoritative worksheet (for example a client/alternate view where verified);
- administrative/support sheets such as number-allocation, deleted-document or schedule-data sheets;
- superseded workbook versions.

Every exclusion must appear in a source-selection/exclusion report with the reason. Nothing should disappear silently.

## 8. Project mismatch validation

Historical Excel files can contain copied/stale sheets.

Where a project identity can be inferred both from the selected source file/path and from content inside the worksheet, compare them.

If they conflict, emit a `PROJECT_MISMATCH` warning and do not silently include the sheet as if the project identity were certain.

## 9. Encrypted/protected workbooks

At least one protected/encrypted workbook was observed during planning (`2722 Deliverable.xlsx`).

Required behavior:

1. detect encrypted/unreadable workbook;
2. report it explicitly as not indexed;
3. provide a supported way to rerun with a password;
4. never commit passwords to this public repository.

Recommended local mechanism:

- a gitignored local password file, environment variable or interactive prompt.

Recommended GitHub Actions mechanism:

- repository/action secret.

If an encrypted file cannot be processed, the generated report must state that the index is not fully complete for that source.

## 10. Planned output schema

The exact set of project-specific/event columns will be finalized from profiling, but the core schema is frozen as follows.

### Project

- `Project No.`
- `Project Name`

### Classification

- `Source Family`
- `Discipline`
- `Category`
- `Subcategory`
- `Original Worksheet`
- `Original Section`
- `Classification Rule ID`
- `Classification Confidence`

### Document

- `Document No.`
- `Document Title`

### Revision

- `Revision`

### Event / transaction

Fields discovered from source workbooks, expected to include some combination of:

- event type;
- submission/issue date;
- transmittal/reference;
- response/status;
- response date/reference;
- comments/remarks.

### Navigation

- `Document Link`
- `Original_Hyperlink_Target`

### Pivot/helper keys

- `Global_Document_Key`
- `Source_Document_Key`
- `Revision_Key`
- `Event_Key`
- `Document_Row_Flag`
- `Revision_Row_Flag`
- `Is_Latest_Revision`
- `Is_Latest_Event`

### Traceability / QA

- `Source File`
- `Source Modified Date`
- `Source Sheet`
- source row/cell locator where practical;
- parsing/classification status;
- warning field(s).

## 11. Planned outputs

### User-facing

`NMDC_DOCUMENT_INDEX.xlsx`

Requirements:

- one visible worksheet named `INDEX`;
- Excel Table;
- autofilter;
- frozen header;
- readable column widths/formats;
- real Excel date values where applicable;
- preserved hyperlinks;
- PivotTable-friendly helper fields.

### Review/machine output

`NMDC_DOCUMENT_INDEX.csv`

This is important because GitHub can diff/review CSV data much more effectively than a binary `.xlsx` file.

### Validation/profiling outputs

Planned initial reports:

- `source_inventory.csv`
- `workbook_profiles.json`
- `source_selection_report.md`
- `classification_discovery.csv`

Additional validation reports may be added during implementation.

## 12. Runtime architecture

The routine refresh process must be deterministic and LLM-free.

```text
DATA/**
   |
   v
Source profiler / selector
   |
   v
Merge-aware extractor
   |
   v
Config-driven classifier + normalizer
   |
   v
Validation / reconciliation gate
   |
   +--> CSV
   +--> XLSX
```

An LLM may be used to design, implement and review the system, but not as a required runtime dependency for normal index refreshes.

## 13. Implementation phases

### Cycle 1 — Read-only Source Profiler

Implement only the profiler and rule discovery layer.

Required outputs:

- `source_inventory.csv`
- `workbook_profiles.json`
- `source_selection_report.md`
- `classification_discovery.csv`

Required discovery:

- all source files;
- trustworthy modified timestamps;
- duplicate/superseded version groups;
- exact duplicates;
- encrypted sources;
- worksheet names;
- used ranges/header candidates;
- merged-cell ranges/patterns;
- hyperlinks;
- likely classification and rule match;
- unknown/ambiguous classifications;
- exclusions and reasons.

**No final consolidated index in Cycle 1.**

### Cycle 2 — Sentinel lossless extractor

Implement extraction against a deliberately difficult representative subset covering:

- METHODS and TECH;
- merged document cells across revisions;
- merged revision cells across several events;
- hyperlinks;
- inconsistent worksheet names;
- date conversion;
- unknown-layout warning behavior.

Every known revision/event in the sentinel cases must survive extraction.

### Cycle 3 — Full extraction

Run the accepted extractor across every selected source workbook.

Perform reconciliation between:

- source data rows/blocks;
- extracted normalized rows;
- structural/header/blank rows intentionally ignored;
- warning/error rows.

No unexplained row loss is acceptable.

### Cycle 4 — Final Excel index

Generate the user-facing XLSX plus the canonical CSV and verify formatting, hyperlinks, keys and PivotTable helpers.

### Cycle 5 — Routine automation

Add local Windows refresh convenience and GitHub automation.

Planned local entry point may be an `UPDATE_INDEX.cmd` or equivalent wrapper.

GitHub Actions should rebuild when relevant `DATA/**` inputs or configuration rules change, while avoiding output-trigger loops.

## 14. Acceptance principles

The system is not accepted merely because it runs without errors.

Acceptance requires evidence that:

- selected sources are correct;
- superseded sources are excluded for the correct reason;
- merged-cell revision/event history is preserved;
- hyperlinks survive where technically available;
- unknown inputs are visible;
- no source rows disappear without reconciliation;
- output classifications are traceable to rule IDs;
- runtime refresh requires no LLM;
- source files remain untouched.

## 15. Collaboration governance

Planned Collaboration ID: `NMDC-DOC-INDEX-001`.

Roles:

- User: owner/final decision maker.
- ChatGPT: architect/planner and independent reviewer.
- Hermes: implementer.
- GitHub PR: source of truth and communication bridge.

Hermes must not approve its own implementation. No merge is authorized unless the user explicitly approves it.

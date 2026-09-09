# AGENTS.md — NMDC Document Index Handoff Guide

## 1. Why this file exists

This repository is intended to be worked on by multiple LLM/agent sessions over time. This file tells a fresh agent what to read, what is already decided, what is still pending, and how to continue without re-inventing the project or asking the user to repeat settled requirements.

## 2. Required reading order

Before proposing or changing implementation, read:

1. `README.md`
2. `PROJECT_SPEC.md`
3. `CLASSIFICATION_MODEL.md`
4. the current open collaboration PR and its complete discussion, if one exists
5. the actual current repository tree and changed files

Treat GitHub as the source of truth for implementation state.

## 3. Project mission

Build a deterministic, repeatable document-indexing pipeline for the Excel project deliverable registers in `DATA/`.

The pipeline must consolidate relevant records into one normalized master dataset while preserving:

- Source Family (`METHODS` / `TECH`);
- document classification hierarchy;
- document numbers/titles;
- all revisions;
- all submission/response/event rows;
- source hyperlinks where available;
- original worksheet/section identity;
- traceability back to source workbook/cell;
- explicit warnings for anything that could not be processed safely.

Routine refresh must not depend on an LLM.

## 4. Roles and authority

Planned Collaboration ID: `NMDC-DOC-INDEX-001`

### User

- Product owner and final decision maker.
- Approves architecture changes and merges.

### ChatGPT

- Planner/architect.
- Independent reviewer of Hermes implementation.
- Reviews actual diffs, tests and generated outputs.
- Does not delegate final approval to Hermes.

### Hermes

- Implementer/executor, preferably using a low-cost model.
- Must follow the current agreed specification.
- Must not approve its own work.
- Must not merge unless the user explicitly authorizes it.

### GitHub

- Source of truth and communication/audit bridge.
- Implementation should occur through a feature branch and PR once collaboration starts.

## 5. Current status

**Planning/specification is complete. Implementation has not started.**

The next intended implementation cycle is:

### Cycle 1 — Read-only Source Profiler

Do not build the final index yet.

Required outputs are expected to include:

- `source_inventory.csv`
- `workbook_profiles.json`
- `source_selection_report.md`
- `classification_discovery.csv`

Cycle 1 must discover and report, at minimum:

- all candidate workbooks;
- source family;
- trustworthy source modified timestamps;
- logical duplicate/superseded groups;
- exact byte duplicates;
- which file is selected and why;
- excluded files/sheets and reasons;
- encrypted/unreadable workbooks;
- worksheets and used ranges;
- likely header rows;
- merged-cell ranges/patterns;
- hyperlink presence;
- candidate classification and matching rule;
- unknown/ambiguous classifications;
- project-number mismatches between workbook/path and internal content.

No final `NMDC_DOCUMENT_INDEX.xlsx` should be produced as the implementation deliverable for Cycle 1.

## 6. Non-negotiable technical rules

### 6.1 DATA is read-only

Never edit, unmerge, repair, resave or normalize source workbooks in `DATA/`.

### 6.2 Preserve merged-cell semantics

Do not use blind fill-down.

A merged document cell can span multiple revision rows. A merged revision cell can span multiple event/transaction rows.

The logical hierarchy is:

```text
PROJECT -> DOCUMENT -> REVISION -> EVENT/TRANSACTION
```

The extractor must inherit merged values only within the exact source merge range.

### 6.3 Row granularity

The normalized canonical dataset is lossless:

> one row per event/transaction for one revision of one document.

Pivot helper flags/keys are used to derive one-row-per-document or one-row-per-revision views.

### 6.4 Hyperlinks

Preserve source hyperlinks where available, including native Excel hyperlinks and usable `HYPERLINK(...)` formulas.

If a hyperlink cannot be preserved, report it.

### 6.5 Duplicate/superseded files

Do not combine historical versions of the same logical register.

For a confirmed version group:

- select the newest source using the trustworthy source modified date;
- exclude older versions;
- do not recover older revisions/events from superseded files;
- report the selection decision.

Do not confuse "same project" with "same logical register". One project may contain several legitimate registers.

### 6.6 File modified dates

Do not assume Git checkout filesystem mtimes are original file modified dates.

Prefer original workbook/core-document modified metadata where accessible. Use a documented fallback and warning when it is unavailable or unreliable.

### 6.7 Classification

Normalized hierarchy:

`Source Family -> Discipline -> Category -> Subcategory`

Classification must be configuration-driven. See `CLASSIFICATION_MODEL.md`.

Do not hard-code NMDC taxonomy labels throughout parser logic.

Unknown/ambiguous inputs must remain visible as warnings/`UNCLASSIFIED` rather than being silently forced to a guessed category.

### 6.8 Original context

Always preserve original worksheet and section identity separately from normalized classification.

### 6.9 Encrypted files

Encrypted/unreadable files must be explicitly reported.

A password-assisted reprocessing path is required later, but credentials must never be committed to this public repository.

### 6.10 No silent data loss

Every selected source row/block must be explainable as one of:

- extracted;
- structural/header/blank and intentionally ignored;
- explicitly excluded by rule;
- warning/error requiring review.

No unexplained row loss is acceptable.

## 7. Classification governance

`CLASSIFICATION_MODEL.md` defines Classification Model v1.

The implementation should ultimately use an editable table such as:

`config/classification_rules.csv`

A non-programmer should be able to change aliases and normalized classification results without editing Python.

Cycle 1 may discover additional aliases or subcategories, but Hermes must propose them through the PR/report. Do not silently invent new business taxonomy without review.

## 8. Default exclusions

Default exclusion classes include:

- templates/format workbooks;
- reference/prerequisite matrices;
- installation-aid registers unless explicitly enabled later;
- personal/status trackers;
- temporary Excel files;
- support/administrative sheets;
- verified duplicate/alternate worksheet views;
- superseded workbook versions.

All exclusions must be reported with reasons.

## 9. Known planning observations

The following issues were observed during pre-implementation analysis and should be treated as test cases, not assumptions to ignore:

- historical Methods workbooks use vertically merged cells to represent one document across multiple revisions;
- some revision cells are themselves merged across multiple submission/response events;
- worksheet names vary significantly between projects;
- TECH includes multiple disciplines and document types rather than only Procedures/Sketches/Drawings;
- a protected/encrypted `2722 Deliverable.xlsx` was observed and must not be silently skipped;
- some projects have multiple source versions; newest-only selection is required;
- duplicate/alternate worksheets can exist inside a selected workbook;
- at least one historical worksheet appeared to contain a project number inconsistent with the source workbook, motivating `PROJECT_MISMATCH` validation.

Do not hard-code these individual filenames/observations as the only cases. Build generic logic and use them as sentinel tests.

## 10. Expected implementation style

Prefer:

- Python;
- deterministic parsing;
- explicit configuration;
- small testable modules;
- CSV/JSON/Markdown profiler outputs;
- unit/integration tests using representative workbook fixtures or safe source reads;
- clear warnings and reason codes;
- reproducible output ordering;
- minimal runtime dependencies.

Avoid:

- runtime LLM classification;
- opaque heuristic guessing;
- source workbook mutation;
- silent exception swallowing;
- one giant project-specific script full of special cases;
- treating successful execution as proof of data completeness.

## 11. Planned implementation sequence

1. Cycle 1 — read-only Source Profiler.
2. Independent review of profiler diff and outputs.
3. Cycle 2 — sentinel lossless extractor on difficult representative workbooks.
4. Independent review and row-level verification.
5. Cycle 3 — full selected-source extraction and reconciliation.
6. Cycle 4 — final XLSX + canonical CSV.
7. Cycle 5 — local refresh wrapper + GitHub Actions automation.

Do not skip directly to a later cycle unless the user and planner explicitly change the plan.

## 12. PR/review expectations

When collaboration starts, Hermes should provide a structured report in the PR containing:

- Collaboration ID;
- cycle number;
- implementation summary;
- files changed;
- commands/tests run;
- output artifacts generated;
- warnings/known limitations;
- any proposed changes to requirements/taxonomy;
- exact commit SHA.

The reviewer should verify the actual repository state, diff, test results and generated outputs rather than trusting the report alone.

## 13. Merge policy

No agent may assume merge permission.

A review result may say the implementation is acceptable, but **the PR must not be merged unless the user explicitly authorizes it**.

## 14. How a fresh LLM should resume

If you are a fresh LLM/agent arriving at this repository:

1. Read the required documents listed in Section 2.
2. Inspect the current repo/PR state from GitHub.
3. Determine the current collaboration cycle from the PR discussion and repository artifacts.
4. Continue from the last accepted state; do not re-ask the user for requirements already recorded here.
5. If implementation and documentation conflict, stop and surface the conflict to the planner/user rather than guessing.
6. If no collaboration PR exists yet, the next planned action is to start `NMDC-DOC-INDEX-001`, Cycle 1, with the read-only profiler scope defined above.

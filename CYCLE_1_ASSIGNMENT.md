# NMDC-DOC-INDEX-001 — Cycle 1 Assignment

## Collaboration identity

- Collaboration-ID: `NMDC-DOC-INDEX-001`
- Cycle: `1`
- Repository: `Harfoushennium/nmdc-doc-index`
- Role split: User = owner/merge authority; ChatGPT = planner/reviewer; Hermes = implementer; GitHub PR = communication bridge.

## Cycle 1 objective

Build a **read-only source profiler and classification-discovery tool only**.

The purpose of Cycle 1 is to understand every candidate Excel workbook and worksheet under `DATA/`, validate source-selection logic, and produce the evidence needed for the user and planner to approve **Classification Model v2** before any full extraction begins.

Cycle 1 is a discovery/validation cycle. It is **not** the final indexing cycle.

## Mandatory reading before implementation

Read, in order:

1. `README.md`
2. `PROJECT_SPEC.md`
3. `CLASSIFICATION_MODEL.md`
4. `AGENTS.md`
5. this file
6. the complete current collaboration PR discussion

If these sources conflict, stop and report the conflict in the PR instead of choosing one silently.

## Hard scope boundaries

### Allowed

Hermes may:

- add profiler source code;
- add tests;
- add non-secret configuration needed for profiling;
- read Excel metadata, workbook structures, styles, merges, formulas, hyperlinks, worksheets and representative cell content;
- calculate hashes for duplicate detection;
- infer candidate project/register/version groups;
- propose candidate classification mappings and aliases;
- generate CSV/JSON/Markdown discovery outputs;
- document warnings and unsupported structures.

### Forbidden in Cycle 1

Hermes must **not**:

- modify, unmerge, repair, resave, rename, delete or otherwise mutate any file under `DATA/`;
- produce the final `NMDC_DOCUMENT_INDEX.xlsx`;
- build the full normalized document/revision/event extraction pipeline;
- silently invent or approve new business taxonomy;
- silently exclude unknown workbooks or worksheets;
- recover data from superseded historical register versions after a newer version is selected;
- commit passwords or credentials;
- merge the PR;
- force-push.

No work beyond Cycle 1 is authorized until the user and planner explicitly approve Classification Model v2 and authorize Cycle 2.

## Required deliverables

Cycle 1 must generate these four primary outputs in a reproducible location such as `outputs/cycle1/`:

1. `source_inventory.csv`
2. `workbook_profiles.json`
3. `source_selection_report.md`
4. `classification_discovery.csv`

Hermes may add supporting reports where useful, but these four are mandatory.

## 1. source_inventory.csv

One row per candidate Excel workbook found recursively under `DATA/`.

Required fields should include at least:

- relative path;
- filename;
- source family inferred from folder (`METHODS`, `TECH`, or `UNKNOWN`);
- file extension;
- file size;
- content hash (prefer SHA-256);
- Excel/core-document created timestamp if readable;
- Excel/core-document modified timestamp if readable;
- timestamp source/reliability flag;
- encryption/readability status;
- inferred project number(s);
- inferred project/register identity;
- duplicate/version-group identifier if applicable;
- selected/excluded status;
- selection/exclusion reason;
- selected replacement file when superseded;
- warning codes.

Do not use Git checkout filesystem mtime as if it were the authoritative original modified date.

## 2. workbook_profiles.json

Produce a structured profile for every candidate workbook, including excluded and unreadable files where possible.

For each workbook capture at least:

- workbook path and source family;
- selected/excluded/unreadable status;
- workbook metadata;
- worksheet list in source order;
- hidden/visible sheet state;
- approximate used range / max row / max column;
- likely header row(s);
- representative non-empty header values;
- merged ranges and summary counts/patterns;
- hyperlink counts/types where detectable;
- formulas of interest, especially `HYPERLINK(...)` formulas;
- candidate internal project number(s);
- project mismatch findings;
- candidate worksheet role/classification;
- support/admin/duplicate-view suspicion;
- warnings/errors.

For large sheets, profile safely without dumping excessive cell contents or reproducing the whole workbook.

## 3. source_selection_report.md

Human-readable report explaining source selection and exclusions.

It must clearly show:

- candidate source count by family;
- selected source count;
- excluded source count;
- encrypted/unreadable source count;
- exact byte-duplicate groups;
- probable version groups;
- selected newest file for each confirmed version group;
- timestamps used to select it;
- older files discarded and why;
- template/reference/admin files excluded and why;
- duplicate/alternate worksheets excluded or proposed for exclusion and why;
- any uncertain duplicate/version decisions that require human review;
- all project-number mismatches;
- all unknown workbook/sheet layouts.

Selection rule: **same project is not enough to declare duplication**. Group versions only when there is evidence they are versions of the same logical register.

For a confirmed version group, select the newest trustworthy source and exclude older versions completely from later extraction.

## 4. classification_discovery.csv

This is the most important Cycle-1 output for Classification Model v2.

Provide at least one row per relevant worksheet or internal section. Where a single worksheet contains distinct sections, represent the sections separately.

Required fields should include at least:

- workbook path;
- project number;
- source family;
- exact original worksheet name;
- original section/header identity, if any;
- normalized worksheet-name form;
- sample document-number patterns;
- sample title keywords/patterns;
- proposed discipline;
- proposed category;
- proposed subcategory;
- matched Classification Model v1 rule/alias, if any;
- match basis (`WORKSHEET`, `SECTION`, `HEADER`, `DOC_NUMBER`, `TITLE`, etc.);
- confidence;
- proposed include/exclude action;
- exclusion reason where relevant;
- suggested new aliases/keywords;
- warning/review-required flag;
- notes.

Classification discovery must preserve the distinction between:

`Source Family -> Discipline -> Category -> Subcategory`

and must always preserve original worksheet/section names separately.

### Classification governance

Classification Model v1 in `CLASSIFICATION_MODEL.md` is a starting taxonomy, not permission for Hermes to silently expand it.

Cycle 1 may **propose**:

- new worksheet aliases;
- new section aliases;
- new document-number tokens;
- new title keywords;
- additional disciplines/categories/subcategories when genuinely observed;
- changes to exclusions.

But these proposals must be visible in the discovery output/report and are not considered approved until the user and planner approve Classification Model v2.

Unknown or ambiguous inputs must remain `UNCLASSIFIED` / `REVIEW_REQUIRED` rather than being forced to the nearest label.

## Required duplicate/version logic to investigate

The profiler must support evidence-based grouping using signals such as:

- normalized filename similarity;
- project number;
- workbook/sheet structure similarity;
- internal register title/header identity;
- exact file hash;
- workbook modified metadata;
- source family;
- logical register type.

Do not rely on filename alone.

Known examples to validate generically include projects with multiple register versions such as 2820 and 3291, but do not hard-code those projects as special cases.

## Required workbook/sheet exclusion discovery

Report likely exclusions, including but not limited to:

- templates/format workbooks;
- prerequisite/reference matrices;
- installation-aid registers unless later enabled;
- temporary Excel files;
- administrative/support sheets;
- deleted-document sheets;
- scheduling/P6 data sheets;
- blank sheets;
- alternate/duplicate worksheet views such as client copies;
- superseded workbook versions.

All exclusion decisions must have reason codes and remain reviewable.

## Encrypted/unreadable source behavior

An encrypted or unreadable workbook must never be silently skipped.

Record at least:

- file path;
- status `ENCRYPTED` or `UNREADABLE`;
- failure reason;
- whether metadata was available;
- whether a password-assisted retry would be possible later.

Cycle 1 does not need to implement the final password UI, but the profiler design must make later password-assisted reprocessing possible without committing credentials.

Known sentinel: `2722 Deliverable.xlsx` was previously observed as encrypted. Treat it as a generic encrypted-file test, not a one-off hard-coded exception.

## Merged cells and hyperlink discovery

Cycle 1 is not the extraction cycle, but the profiler must identify structures needed for Cycle 2.

For each relevant sheet report:

- merged-range count;
- representative merge dimensions/patterns;
- whether merges span likely document/revision/event rows;
- hyperlink count;
- native hyperlink count where distinguishable;
- `HYPERLINK(...)` formula count where distinguishable.

Do not unmerge or fill down the source workbook.

## Project mismatch validation

Compare project identity inferred from the path/workbook with project numbers/titles found internally.

If there is a mismatch, emit a `PROJECT_MISMATCH` warning with evidence and do not silently trust either side.

## Implementation expectations

Prefer:

- Python;
- deterministic behavior;
- minimal dependencies;
- small testable modules;
- stable sorted outputs;
- explicit warning/reason codes;
- configuration rather than project-specific branching;
- standard library where practical plus appropriate Excel libraries where needed.

Avoid:

- runtime LLM calls;
- source mutation;
- broad exception swallowing;
- uncontrolled fuzzy matching;
- project-number hard-coding;
- treating successful execution as proof of complete profiling.

## Required tests

Tests must cover at minimum:

1. recursive discovery of Excel files under both source families;
2. source family preservation;
3. exact hash duplicate detection;
4. modified-date extraction without depending on checkout mtime;
5. selection of newest source within a confirmed version group;
6. no conflation of different logical registers from the same project;
7. encrypted/unreadable workbook reporting;
8. worksheet/section classification discovery;
9. ambiguous classification producing review-required/unclassified state;
10. merged-range profiling;
11. native and/or formula hyperlink profiling;
12. project mismatch warning;
13. explicit exclusion reporting;
14. deterministic output ordering.

Tests may use small synthetic fixtures where appropriate, but the implementation must also be run read-only against the real `DATA/` tree before reporting completion.

## Cycle 1 acceptance gates

Cycle 1 is ready for planner review only when:

- all mandatory deliverables exist;
- the profiler completes against the current real `DATA/` tree without modifying it;
- every candidate workbook appears in `source_inventory.csv` exactly once;
- every selected/excluded/unreadable workbook has an explicit reason/status;
- no unknown/ambiguous classification is silently converted into a confident taxonomy result;
- duplicate/version decisions are evidenced and reviewable;
- classification discovery contains enough evidence to build Classification Model v2;
- tests pass;
- the exact implementation commit SHA is reported.

## Required HERMES_REPORT in PR

When Cycle 1 implementation is complete, post **one structured `HERMES_REPORT`** to the PR containing:

- Collaboration-ID: `NMDC-DOC-INDEX-001`
- Cycle: `1`
- exact HEAD commit SHA;
- implementation summary;
- files changed;
- commands run;
- tests run and results;
- real-data profiler run result;
- deliverable paths;
- counts: discovered / selected / excluded / unreadable / unclassified / warnings;
- duplicate/version groups found;
- project mismatches found;
- encrypted sources found;
- proposed taxonomy/alias additions for Classification Model v2;
- known limitations;
- explicit statement confirming `DATA/` was not modified;
- explicit statement confirming no full extraction/final index was implemented.

Do not self-approve and do not merge.

## Stop condition

After posting the Cycle-1 `HERMES_REPORT`, stop.

Do not begin Cycle 2 until the user and ChatGPT reviewer approve Classification Model v2 and explicitly authorize the next cycle.

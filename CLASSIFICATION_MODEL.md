# NMDC Document Index — Classification Model v2

## 1. Status and purpose

Classification Model v2 is the controlled taxonomy and rule-governance model derived from approved Cycle-1 profiling evidence.

It preserves the four normalized levels:

```text
Source Family
  -> Discipline
      -> Category
          -> Subcategory
```

Cycle-1 evidence showed that the 34 original classification `REVIEW_REQUIRED` rows could be resolved without inventing new taxonomy values. Version 2 therefore focuses on safe normalization, missing aliases, controlled fallbacks, narrow exclusions, context-guarded refinements, and a strict boundary between worksheet discovery and per-document refinement.

This model does **not** authorize Cycle 2 extraction by itself.

## 2. Required source context

Always preserve separately from normalized classification:

- original worksheet;
- original section;
- classification rule ID;
- classification confidence;
- source workbook path;
- warning/review-required state;
- sampled document numbers/titles used only as audit evidence during discovery.

Classification normalization must never overwrite source evidence.

## 3. Source Family

Allowed values:

- `METHODS`
- `TECH`

Source Family is derived primarily from the top-level `DATA/` folder.

## 4. METHODS hierarchy

### 4.1 OFFSHORE INSTALLATION

```text
METHODS
└── OFFSHORE INSTALLATION
    ├── PROCEDURE
    │   └── INSTALLATION PROCEDURE
    └── SKETCH
        └── ENGINEERING SKETCH
```

Supported worksheet context includes:

- `Installation Procedures`
- `Installation Procedure`
- `Procedures`
- `Sketch`
- `Sketches`

The real `2171-2172 -Document Deliverables LATEST.xlsx` workbook contains one verified nonstandard worksheet named `2171-2172` that is an installation-procedure register. Because the profiler's structural header extraction does not expose a generic procedure phrase reliably for this sheet, v2 uses one narrow path-qualified worksheet rule bound to this exact workbook and worksheet. The rule must not apply to an unrelated workbook that happens to contain a worksheet named `2171-2172`.

### 4.2 MARINE OPERATIONS

```text
METHODS
└── MARINE OPERATIONS
    └── DRAWING
        ├── ANCHOR PATTERN
        ├── DP SETUP PLAN
        ├── SETUP PLAN
        └── METHOD DRAWING
```

Typical context:

- `Setup Plans & Anchor Patterns`
- `ANCHOR PATTERN`
- `DP SETUP`, `DP SET-UP`, `DP SET UP`
- `SETUP PLAN`, `SET-UP PLAN`

`METHOD DRAWING` remains the safe worksheet-level fallback for setup/anchor-plan drawing registers.

`ANCHOR PATTERN`, `DP SETUP PLAN`, and `SETUP PLAN` title refinements are **per-document refinements** only. They may run only when the individual document row already has `MARINE OPERATIONS -> DRAWING` context. They must never convert a worksheet-wide `SKETCH` or `EXTERNAL / INPUT` classification.

### 4.3 EXTERNAL / INPUT

```text
METHODS
└── EXTERNAL / INPUT
    ├── DRAWING
    │   └── INCOMING DRAWING
    └── DOCUMENT
        └── INCOMING TECHNICAL DOCUMENT
```

Reliable internal `DRAWINGS` / `DOCUMENTS` section evidence overrides the broad mixed-sheet label.

## 5. TECH hierarchy

### 5.1 PIPELINE & CABLE

```text
TECH
└── PIPELINE & CABLE
    ├── DOCUMENT
    │   ├── GENERAL TECHNICAL DOCUMENT
    │   ├── REPORT
    │   ├── ANALYSIS
    │   ├── ANALYSIS REPORT
    │   ├── CALCULATION
    │   ├── PROCEDURE
    │   ├── SPECIFICATION
    │   └── TECHNICAL NOTE
    ├── DRAWING
    │   ├── ENGINEERING DRAWING
    │   └── CUT LIST
    └── SKETCH
        └── ENGINEERING SKETCH
```

Supported worksheet aliases include:

- `Documents - Pipeline & Cable`
- `PIPELINE & CABLE`
- `PIPELINE & CABLE doc`
- bare `Pipeline`
- `TN-PL`

A broad Pipeline document worksheet is discovered as:

```text
PIPELINE & CABLE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
```

Individual document rows may later refine to REPORT / ANALYSIS / PROCEDURE / TECHNICAL NOTE, etc.

### 5.2 NAVAL & MARINE

```text
TECH
└── NAVAL & MARINE
    ├── DOCUMENT
    │   ├── GENERAL TECHNICAL DOCUMENT
    │   ├── REPORT
    │   ├── ANALYSIS
    │   ├── ANALYSIS REPORT
    │   └── TECHNICAL NOTE
    ├── DRAWING
    │   ├── ENGINEERING DRAWING
    │   ├── ANCHOR PATTERN
    │   └── DP SETUP PLAN
    └── SKETCH
        └── ENGINEERING SKETCH
```

Supported worksheet aliases include:

- `Documents - Naval Marine`
- `Documents - Naval & Marine`
- `Naval & Marine doc`
- bare `Naval Marine`
- bare `NAVAL MARINE`
- `TN-NA`

A broad Naval & Marine document worksheet is discovered as `GENERAL TECHNICAL DOCUMENT`; per-document refinement happens only after row extraction.

### 5.3 STRUCTURAL

```text
TECH
└── STRUCTURAL
    └── DOCUMENT
        └── TECHNICAL NOTE
```

Known controlled pattern: `TN-ST`.

### 5.4 GENERAL / MULTIDISCIPLINE

```text
TECH
└── GENERAL / MULTIDISCIPLINE
    ├── DOCUMENT
    │   ├── GENERAL TECHNICAL DOCUMENT
    │   ├── REPORT
    │   ├── ANALYSIS
    │   ├── ANALYSIS REPORT
    │   ├── CALCULATION
    │   ├── PROCEDURE
    │   ├── SPECIFICATION
    │   └── TECHNICAL NOTE
    ├── DRAWING
    │   ├── ENGINEERING DRAWING
    │   └── CUT LIST
    └── SKETCH
        └── ENGINEERING SKETCH
```

Use this when document type is valid but engineering discipline cannot be safely established.

A generic TECH worksheet named `Documents` is discovered as:

```text
GENERAL / MULTIDISCIPLINE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
```

Do not let one sampled report/procedure/analysis title collapse the whole worksheet to one subtype.

### 5.5 COMMISSIONING

```text
TECH
└── COMMISSIONING
    └── PROCEDURE
        └── OPERATIONAL TEST PROCEDURE
```

Known controlled context includes `COMMISSION - List of OTP`.

## 6. Administrative/support exclusions

Examples include:

- `Available Numbers`
- `Deleted Documents`
- `to be deleted`
- `P6_Data`
- generic support `Sheet1`, `Sheet2`, `Sheet3`
- `Pre -Req. Matrix`
- aggregate/alternate support views such as `P6_List_Final` / `Data_List_Final_*`
- verified template/format workbooks
- verified personal working trackers

Every exclusion must record its rule ID and reason.

### 6.1 CLIENT worksheet safety rule

Do **not** globally exclude worksheets named `CLIENT`.

Cycle-1 evidence verified the `CLIENT` worksheet in:

```text
DATA/TECH/3291 DOCUMENT REGISTER Latest.xlsx
```

as an alternate/client duplicate view. It may therefore be excluded only by a path-qualified rule bound to that verified workbook.

A `CLIENT` worksheet in any unrelated workbook remains eligible for normal classification/review.

## 7. Configurable rule table

Authoritative rule table:

`config/classification_rules.csv`

Required columns:

| Column | Purpose |
|---|---|
| `Rule_ID` | Stable unique identifier |
| `Enabled` | `YES` / `NO` |
| `Priority` | Evaluation order |
| `Source_Family` | `METHODS`, `TECH`, `ANY` |
| `Match_Scope` | `FILE`, `WORKSHEET`, `HEADER`, `SECTION`, `DOC_NUMBER`, `TITLE` |
| `Match_Type` | `EXACT`, `CONTAINS`, `REGEX`, `FUZZY` |
| `Match_Words` | User-editable aliases/patterns |
| `Exclude_Words` | Terms that invalidate the rule |
| `Path_Qualifier` | Optional narrow source-path constraint |
| `Requires_Discipline` | Optional required current discipline before refinement |
| `Requires_Category` | Optional required current category before refinement |
| `Discipline` | Normalized result or refinement instruction |
| `Category` | Normalized result or refinement instruction |
| `Subcategory` | Normalized result or refinement instruction |
| `Include` | `YES` / `NO` |
| `Min_Confidence` | Fuzzy threshold |
| `Stop_On_Match` | Stops lower-priority rules within the same scope |
| `Notes` | Human-readable explanation |

`Path_Qualifier` must remain exceptional and auditable.

`Requires_Discipline` and `Requires_Category` are taxonomy safety gates. Leave them blank for true base-classification rules. Use them whenever a refinement is valid only inside an already-established taxonomy branch.

Examples:

- `M011` / `M012` / `M013` require `MARINE OPERATIONS` + `DRAWING`;
- `T070`–`T076` require category `DOCUMENT`;
- `T043`–`T045` require category `DOCUMENT` before technical-note document-number refinement;
- `M002` is an exceptional path-qualified `WORKSHEET` rule for the verified 2171-2172 procedure register.

## 8. Discovery classification versus per-document refinement

This distinction is mandatory.

### 8.1 Worksheet/section discovery

Cycle-1 / Classification-v2 discovery classifies source structure. It may use only:

```text
FILE
WORKSHEET
SECTION
HEADER
```

Sample document numbers and sample titles may be displayed in discovery outputs for audit, but they must **not** be joined and passed to `DOC_NUMBER` / `TITLE` rules for worksheet-wide classification.

Reason: a worksheet can contain many different document types. One sampled Analysis Report is not proof that every document in that worksheet is an Analysis Report.

### 8.2 Per-document refinement

`DOC_NUMBER` and `TITLE` rules apply only when the evidence belongs to one individual document record.

They may refine a valid base `DOCUMENT` classification to:

- `REPORT`
- `ANALYSIS`
- `ANALYSIS REPORT`
- `CALCULATION`
- `PROCEDURE`
- `SPECIFICATION`
- `TECHNICAL NOTE`

Methods drawing title rules may similarly refine one individual `MARINE OPERATIONS -> DRAWING` document to `ANCHOR PATTERN`, `DP SETUP PLAN`, or `SETUP PLAN`.

The current PR preserves this rule-engine capability but does not implement the row extractor. That belongs to a later authorized cycle.

## 9. Rule precedence and context safety

The general rule engine supports:

```text
FILE
  -> WORKSHEET
  -> SECTION
  -> HEADER
  -> DOC_NUMBER
  -> TITLE
```

However worksheet discovery intentionally stops after `HEADER`.

A later row-level refinement may narrow or clarify a valid classification, but it must not jump across taxonomy branches. Before applying a refinement, the engine evaluates configured `Requires_Discipline` / `Requires_Category` guards.

Forbidden examples:

- a Methods `SKETCH` becoming `ANCHOR PATTERN` because sampled title text mentions anchor pattern;
- incoming `DOCUMENTS` / `DRAWINGS` being overwritten by unrelated title text;
- a TECH drawing/sketch becoming `DOCUMENT -> REPORT/PROCEDURE` from one title;
- one sampled `TN-PL` number converting an entire generic Documents worksheet to Pipeline Technical Note;
- one sampled Analysis Report converting a mixed Pipeline/Naval/General worksheet to `ANALYSIS REPORT`.

## 10. Whitespace and name normalization

Before `WORKSHEET` and `SECTION` matching:

- trim leading/trailing whitespace;
- collapse repeated whitespace.

Preserve the original worksheet/section text separately.

This safely resolves variants such as:

- `Installation Procedures `
- `Cut-lists `
- `Sketch `

without creating workbook-specific aliases.

## 11. Fuzzy matching

Fuzzy matching remains a fallback, not the primary mechanism.

Recommended defaults:

- >= 90% and one unambiguous candidate -> automatic classification;
- 80–89% -> suggestion + manual review;
- < 80% -> `UNCLASSIFIED`.

Thresholds remain configurable.

## 12. Safety rules

1. Never silently skip an unknown worksheet.
2. Never force an uncertain worksheet merely to achieve 100% coverage.
3. Preserve original worksheet and section text.
4. Every classification must be traceable to a rule ID or explicit fallback reason.
5. Business taxonomy belongs in configuration, not scattered Python constants.
6. Unknown future inputs remain `UNCLASSIFIED` / `REVIEW_REQUIRED`.
7. Generic TECH `Documents` uses `GENERAL / MULTIDISCIPLINE` rather than an invented discipline.
8. File-qualified exclusions/exceptions must be narrow, auditable, and regression-tested.
9. Refinement rules must not cross taxonomy branches.
10. Discovery must not use joined sampled rows as row-level classification evidence.
11. Zero `REVIEW_REQUIRED` rows is necessary but not sufficient proof of taxonomy correctness.
12. Classification work must not modify `DATA/`.

## 13. Cycle-1 discovery resolution summary

| Group | Rows | v2 resolution |
|---|---:|---|
| Methods `Installation Procedures ` | 19 | whitespace-safe worksheet match |
| Methods `2171-2172` | 1 | narrow path-qualified worksheet exception |
| TECH `Cut-lists ` | 6 | whitespace-safe match -> CUT LIST |
| TECH `Sketch ` | 2 | whitespace-safe match -> ENGINEERING SKETCH |
| TECH generic `Documents ` | 2 | safe General/Multidiscipline document base |
| TECH `Naval Marine` / `NAVAL MARINE` | 2 | complete T010 aliases |
| TECH bare `Pipeline` | 1 | complete T001 alias |
| 3291 `CLIENT` | 1 | narrow path-qualified exclusion |
| **Total** | **34** | **all explicitly accounted for** |

No new taxonomy value was required.

## 14. Required regression coverage

At minimum verify:

1. trailing-space Methods Installation Procedures;
2. verified 2171-2172 path-qualified structural exception;
3. same worksheet name in another workbook does not inherit the exception;
4. trailing-space Cut-lists;
5. trailing-space Sketch;
6. Pipeline / Naval aliases;
7. generic TECH Documents safe base;
8. 3291 CLIENT scoped exclusion and unrelated CLIENT safety;
9. unknown worksheet remains review-required;
10. original worksheet evidence is preserved;
11. context guards prevent cross-branch refinements;
12. mixed sample titles do not refine a whole worksheet;
13. sampled `TN-*` numbers do not refine a whole worksheet;
14. direct per-document TITLE/DOC_NUMBER refinement still works when explicitly applied to one record;
15. deterministic outputs;
16. Windows/Linux LF consistency;
17. `DATA/` remains read-only.

## 15. Version status

This document is **Classification Model v2**.

The model is ready for independent review only after exact-head CI and Windows validation confirm the complete regression suite, deterministic outputs, zero current classification review rows, semantic mixed-sample safety, and unchanged `DATA/`.

Cycle 2 remains unauthorized until owner approval.

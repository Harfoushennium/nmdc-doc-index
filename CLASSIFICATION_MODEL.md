# NMDC Document Index — Classification Model v2

## 1. Status and purpose

Classification Model v2 is the controlled taxonomy and rule-governance model derived from the approved Cycle-1 profiling evidence.

It supersedes Classification Model v1 for classification decisions while preserving the same four normalized levels:

```text
Source Family
  -> Discipline
      -> Category
          -> Subcategory
```

Cycle-1 evidence showed that the 34 `REVIEW_REQUIRED` rows could be resolved without inventing new taxonomy values. The v2 changes therefore focus on safe normalization, missing aliases, two controlled fallbacks, one narrowly scoped duplicate-view exclusion, and context-guarded refinement rules.

A deeper semantic validation of the first v2 implementation also found that title refinements could cross taxonomy branches even when every row had a confident classification. For example, a Methods `SKETCH` could be incorrectly refined to the drawing-only `ANCHOR PATTERN` subcategory merely because representative title text contained that phrase. Classification Model v2 therefore requires refinement rules to respect already-established taxonomy context.

This model does **not** authorize Cycle 2 extraction by itself.

## 2. Required source context

Always preserve separately from normalized classification:

- `Original Worksheet`
- `Original Section`
- `Classification Rule ID`
- `Classification Confidence`
- source workbook path
- warning/review-required state

Classification normalization must never overwrite source evidence.

## 3. Source Family

Allowed values remain:

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

Supported context includes:

- `Installation Procedures`
- `Installation Procedure`
- `Procedures`
- `Construction and Installation Procedure`
- `Sketch`
- `Sketches`

A title/content fallback may classify a nonstandard Methods worksheet as `INSTALLATION PROCEDURE` only when reliable content explicitly contains an installation-procedure phrase and the current row is still unresolved or already in the procedure branch. It must not overwrite a previously established sketch, drawing, incoming-document, or other classification.

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

`METHOD DRAWING` remains the controlled fallback when a valid Methods setup/anchor-plan drawing cannot be safely refined.

The `ANCHOR PATTERN`, `DP SETUP PLAN`, and `SETUP PLAN` title refinements are valid only after the base classification is already `MARINE OPERATIONS -> DRAWING`. They must never convert `OFFSHORE INSTALLATION -> SKETCH` or `EXTERNAL / INPUT` rows into Marine Operations drawing subcategories.

### 4.3 EXTERNAL / INPUT

```text
METHODS
└── EXTERNAL / INPUT
    ├── DRAWING
    │   └── INCOMING DRAWING
    └── DOCUMENT
        └── INCOMING TECHNICAL DOCUMENT
```

Reliable internal `DRAWINGS` / `DOCUMENTS` section evidence overrides the broad mixed-sheet label. Once this section classification is established, unrelated Methods title refinements must not overwrite it.

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

A generic TECH worksheet named `Documents` is classified to the safe base:

```text
GENERAL / MULTIDISCIPLINE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
```

Title rules may refine the subcategory without inventing a specific discipline, but only while the already-established category is `DOCUMENT`. A drawing or sketch must not become a document merely because its representative title contains words such as `report`, `analysis`, or `procedure`.

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

A `CLIENT` worksheet in any unrelated workbook must remain eligible for normal classification/review.

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
| `Path_Qualifier` | Optional narrow source-path constraint for exceptional rules |
| `Requires_Discipline` | Optional required current discipline before a refinement may apply; multiple allowed values use `|` |
| `Requires_Category` | Optional required current category before a refinement may apply; multiple allowed values use `|` |
| `Discipline` | Normalized result or refinement instruction |
| `Category` | Normalized result or refinement instruction |
| `Subcategory` | Normalized result or refinement instruction |
| `Include` | `YES` / `NO` |
| `Min_Confidence` | Used for fuzzy rules |
| `Stop_On_Match` | Stops lower-priority rules within the same scope |
| `Notes` | Human-readable explanation |

`Path_Qualifier` must remain exceptional. Use it only where Cycle-1 evidence proves a file-specific rule is necessary and safer than a global business rule.

`Requires_Discipline` and `Requires_Category` are taxonomy safety gates. Leave them blank for true base-classification rules. Use them on refinement rules whenever the refinement is valid only inside an already-established taxonomy branch. Multiple allowed contexts are separated with `|`.

Examples:

- `M011` / `M012` / `M013` require `MARINE OPERATIONS` + `DRAWING`;
- `T070`–`T076` require category `DOCUMENT`;
- `T043`–`T045` require category `DOCUMENT` before a technical-note document-number refinement;
- `M002` permits unresolved/procedure contexts so it can classify the verified nonstandard 2171-2172 procedure sheet without overwriting an established sketch or drawing.

## 8. Rule precedence and context safety

Evaluation order remains:

```text
FILE inclusion/exclusion
  -> Source Family
  -> WORKSHEET exact/alias rule
  -> SECTION rule
  -> HEADER / discipline context
  -> DOC_NUMBER refinement
  -> TITLE refinement
  -> controlled FUZZY fallback
```

A later refinement may narrow or clarify a valid classification, but it must not jump across taxonomy branches. Before applying a refinement, the engine evaluates any configured `Requires_Discipline` and `Requires_Category` guards against the classification already established by higher-priority evidence.

Examples of forbidden cross-branch changes:

- `OFFSHORE INSTALLATION -> SKETCH -> ENGINEERING SKETCH` must not become `... -> SKETCH -> ANCHOR PATTERN` because title text happens to mention anchor patterns;
- `EXTERNAL / INPUT -> DOCUMENT -> INCOMING TECHNICAL DOCUMENT` must not become an anchor-pattern drawing classification;
- `GENERAL / MULTIDISCIPLINE -> DRAWING` or `SKETCH` must not become `DOCUMENT -> REPORT/ANALYSIS/PROCEDURE` merely from title words;
- a technical-note document-number token must not convert a drawing into a document.

Specific refinements must preserve already-established higher-level context unless a rule is explicitly authorized to set a different context and its guards permit that transition.

## 9. Whitespace and name normalization

Before `WORKSHEET` and `SECTION` matching, normalize presentation-only whitespace:

- trim leading/trailing whitespace;
- collapse repeated whitespace.

Preserve the original worksheet/section text separately in discovery outputs.

This resolves harmless source variants such as:

- `Installation Procedures `
- `Cut-lists `
- `Sketch `

without creating workbook-specific aliases.

Other controlled normalization may include:

- case;
- punctuation;
- `&` versus `AND`;
- selected singular/plural variants;
- controlled abbreviations.

## 10. Title/document-number refinement

Title/document-number rules may refine a valid base `DOCUMENT` classification to:

- `REPORT`
- `ANALYSIS`
- `ANALYSIS REPORT`
- `CALCULATION`
- `PROCEDURE`
- `SPECIFICATION`
- `TECHNICAL NOTE`

Methods drawing title refinements may similarly refine only a valid `MARINE OPERATIONS -> DRAWING` base classification.

`KEEP EXISTING` means preserve the already-established higher-level context. It does not authorize a rule to apply outside the context declared by its guard columns.

## 11. Fuzzy matching

Fuzzy matching remains a fallback, not the primary mechanism.

Recommended defaults:

- >= 90% and one unambiguous candidate -> automatic classification;
- 80-89% -> suggestion + manual review;
- < 80% -> `UNCLASSIFIED`.

Thresholds remain configurable.

## 12. Safety rules

1. Never silently skip an unknown worksheet.
2. Never force an uncertain worksheet merely to achieve 100% coverage.
3. Preserve original worksheet and section text.
4. Every classification must be traceable to a rule ID or explicit fallback reason.
5. Business taxonomy belongs in configuration, not scattered Python constants.
6. User edits to the rule table must be testable without parser-code changes.
7. Unknown future inputs remain `UNCLASSIFIED` / `REVIEW_REQUIRED`.
8. Generic TECH `Documents` uses `GENERAL / MULTIDISCIPLINE` rather than an invented specific discipline.
9. File-qualified exclusions must be narrow, auditable, and regression-tested against unrelated files.
10. Refinement rules must not cross taxonomy branches unless the transition is explicitly authorized and context-guarded.
11. Zero `REVIEW_REQUIRED` rows is not, by itself, proof of taxonomy correctness; semantic cross-category tests are required.
12. Classification work must not modify `DATA/`.

## 13. Cycle-1 discovery resolution summary

The 34 Cycle-1 `REVIEW_REQUIRED` rows were grouped as follows:

| Group | Rows | v2 resolution |
|---|---:|---|
| Methods `Installation Procedures ` | 19 | whitespace-safe matching -> Installation Procedure |
| Methods `2171-2172` | 1 | title/content fallback -> Installation Procedure |
| TECH `Cut-lists ` | 6 | whitespace-safe matching -> CUT LIST |
| TECH `Sketch ` | 2 | whitespace-safe matching -> ENGINEERING SKETCH |
| TECH generic `Documents ` | 2 | General/Multidiscipline document base + title refinement |
| TECH `Naval Marine` / `NAVAL MARINE` | 2 | complete T010 aliases |
| TECH bare `Pipeline` | 1 | complete T001 alias |
| 3291 `CLIENT` | 1 | narrow path-qualified duplicate-view exclusion |
| **Total** | **34** | **all explicitly accounted for** |

No new taxonomy value was required.

After these rows were resolved, semantic validation identified and corrected cross-category refinement leakage in other already-classified rows. This correction did not add taxonomy values; it tightened when existing refinement rules are allowed to run.

## 14. Required Classification Model v2 regression tests

At minimum verify:

1. Methods trailing-space Installation Procedures;
2. Methods content fallback;
3. trailing-space Cut-lists;
4. trailing-space Sketch;
5. bare Pipeline alias;
6. bare Naval Marine alias;
7. uppercase NAVAL MARINE alias;
8. generic TECH Documents safe fallback;
9. title refinement preserves General/Multidiscipline;
10. 3291 CLIENT scoped exclusion;
11. unrelated CLIENT is not globally excluded;
12. unknown worksheet remains review-required;
13. original worksheet evidence is preserved;
14. v1 sentinel mappings still pass;
15. Methods sketch does not refine to Anchor Pattern / DP Setup / Setup Plan;
16. incoming document/drawing section classification is not overwritten by unrelated Methods title text;
17. installation-procedure title fallback does not overwrite an established sketch;
18. TECH drawing title does not become a document classification;
19. TECH sketch title does not become a document/procedure classification;
20. technical-note document-number tokens do not convert drawings into documents;
21. context-guard columns are loaded from external configuration;
22. `DATA/` remains read-only.

## 15. Version status

This document is **Classification Model v2**.

It is ready for independent implementation review once:

- all current 34 review rows are resolved by reproducible rules;
- cross-category refinement guards pass their semantic regression tests;
- old Cycle-1 regression tests still pass;
- v2 regression tests pass;
- real-data profiling completes deterministically;
- `DATA/` remains unchanged;
- no Cycle-2 extraction or final index work is introduced.

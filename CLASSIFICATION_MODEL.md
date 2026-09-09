# NMDC Document Index — Classification Model v1

## 1. Purpose

This document defines the initial normalized classification hierarchy and the design of the user-editable classification rule table.

The classification system must be dynamic. Business labels and matching words must live in configuration data rather than being hard-coded throughout the parser.

## 2. Normalized hierarchy

Use four normalized levels:

```text
Source Family
  -> Discipline
      -> Category
          -> Subcategory
```

Also preserve source context separately:

- `Original Worksheet`
- `Original Section`
- `Classification Rule ID`
- `Classification Confidence`

### Why four levels

Folder, engineering discipline and document type are different concepts and must not be collapsed.

Examples:

```text
METHODS -> OFFSHORE INSTALLATION -> PROCEDURE -> INSTALLATION PROCEDURE
METHODS -> MARINE OPERATIONS -> DRAWING -> ANCHOR PATTERN
TECH -> PIPELINE & CABLE -> DOCUMENT -> ANALYSIS REPORT
TECH -> NAVAL & MARINE -> DRAWING -> DP SETUP PLAN
```

This structure is intended to remain clean for filtering and PivotTables.

## 3. Source Family

Source Family is derived primarily from the top-level data folder.

Allowed Version 1 values:

- `METHODS`
- `TECH`

Do not infer a different Source Family from worksheet text unless a future explicit rule requires it.

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

Typical source aliases/context:

- `Installation Procedures`
- `Installation Procedure`
- `Procedures`
- `Construction and Installation Procedure`
- `Sketch`
- `Sketches`
- `SKETCHES`

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

- worksheet aliases similar to `Setup Plans & Anchor Patterns`;
- document-title terms such as `ANCHOR PATTERN`;
- `DP SETUP`, `DP SET-UP`, `DP SET UP`;
- `SETUP PLAN`, `SET-UP PLAN`.

`METHOD DRAWING` is a controlled fallback for an otherwise valid Methods setup/anchor-plan drawing that cannot be safely refined to one of the more specific subcategories.

### 4.3 EXTERNAL / INPUT

```text
METHODS
└── EXTERNAL / INPUT
    ├── DRAWING
    │   └── INCOMING DRAWING
    └── DOCUMENT
        └── INCOMING TECHNICAL DOCUMENT
```

This is used for mixed incoming-document workbooks/sheets where the worksheet itself contains separate `DRAWINGS` and `DOCUMENTS` sections.

Internal section classification should override the broad worksheet label when the section is reliable.

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

Typical aliases/context include:

- `Documents - Pipeline & Cable`
- `PIPELINE & CABLE`
- `PIPELINE & CABLE doc`
- `Pipeline`
- `TN-PL`
- `Drawings` / `DRAWINGS`
- `Cut-list`, `Cutlist`, `Cutlists`, `Cut-lists`

Title/document-number rules may refine a generic document to:

- `REPORT`
- `ANALYSIS`
- `ANALYSIS REPORT`
- `CALCULATION`
- `PROCEDURE`
- `SPECIFICATION`
- `TECHNICAL NOTE`

Refinement must not discard an already-established Pipeline & Cable discipline.

### 5.2 NAVAL & MARINE

```text
TECH
└── NAVAL & MARINE
    ├── DOCUMENT
    │   ├── GENERAL TECHNICAL DOCUMENT
    │   ├── REPORT
    │   ├── ANALYSIS REPORT
    │   └── TECHNICAL NOTE
    ├── DRAWING
    │   ├── ENGINEERING DRAWING
    │   ├── ANCHOR PATTERN
    │   └── DP SETUP PLAN
    └── SKETCH
        └── ENGINEERING SKETCH
```

Typical aliases/context include:

- `Documents - Naval Marine`
- `Documents - Naval & Marine`
- `Naval & Marine doc`
- `Naval Marine`
- `NAVAL MARINE`
- `TN-NA`
- `Anchor Pattern`
- `DP`, `DP SETUP PLAN`, `DP SET-UP PLAN`

### 5.3 STRUCTURAL

```text
TECH
└── STRUCTURAL
    └── DOCUMENT
        └── TECHNICAL NOTE
```

Known Version 1 pattern:

- `TN-ST`

Additional Structural document types can be added later through configuration only.

### 5.4 GENERAL / MULTIDISCIPLINE

```text
TECH
└── GENERAL / MULTIDISCIPLINE
    ├── DOCUMENT
    │   ├── GENERAL TECHNICAL DOCUMENT
    │   ├── REPORT
    │   ├── CALCULATION
    │   └── ANALYSIS
    ├── DRAWING
    │   ├── ENGINEERING DRAWING
    │   └── CUT LIST
    └── SKETCH
        └── ENGINEERING SKETCH
```

Use this only when a valid document can be classified by type but its discipline cannot be safely determined.

Prefer a safe general discipline over an invented specific discipline.

### 5.5 COMMISSIONING

```text
TECH
└── COMMISSIONING
    └── PROCEDURE
        └── OPERATIONAL TEST PROCEDURE
```

Known context:

- worksheet similar to `COMMISSION - List of OTP`.

Inclusion remains controlled by rule-table `Include` settings.

## 6. Administrative/support sheets

Some worksheets are not document classifications and should default to exclusion rather than being forced into the hierarchy.

Examples observed/planned for Version 1 exclusion rules:

- `Available Numbers`
- `Deleted Documents`
- `to be deleted`
- `P6_Data`
- blank/generic support `Sheet1`, `Sheet2`, `Sheet3` where profiling shows no document register;
- `Pre -Req. Matrix`
- aggregate or alternate views such as `P6_List_Final` / `Data_List_Final_*` where they would duplicate authoritative source registers;
- verified alternate/client duplicate views such as `CLIENT` when the same document population is already represented by an authoritative worksheet.

Every exclusion must be recorded with its rule ID and reason.

## 7. Configurable rule table

Planned configuration file:

`config/classification_rules.csv`

It must be directly editable in Excel or a text editor.

Recommended columns:

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
| `Discipline` | Normalized result or refinement instruction |
| `Category` | Normalized result or refinement instruction |
| `Subcategory` | Normalized result or refinement instruction |
| `Include` | `YES` / `NO` |
| `Min_Confidence` | Used for fuzzy rules |
| `Stop_On_Match` | Whether lower-priority rules should stop |
| `Notes` | Human explanation |

Implementation may add technical columns if necessary, but the rule table must remain understandable and editable by a non-programmer.

## 8. Example rules

Illustrative Version 1 rules:

| Rule ID | Family | Scope | Match words | Discipline | Category | Subcategory |
|---|---|---|---|---|---|---|
| `M001` | METHODS | WORKSHEET | installation procedures / procedures | OFFSHORE INSTALLATION | PROCEDURE | INSTALLATION PROCEDURE |
| `M010` | METHODS | WORKSHEET | setup plans & anchor patterns | MARINE OPERATIONS | DRAWING | METHOD DRAWING |
| `M011` | METHODS | TITLE | anchor pattern | KEEP EXISTING | KEEP EXISTING | ANCHOR PATTERN |
| `M012` | METHODS | TITLE | dp setup / dp set-up / dp set up | KEEP EXISTING | KEEP EXISTING | DP SETUP PLAN |
| `M020` | METHODS | WORKSHEET | sketch / sketches | OFFSHORE INSTALLATION | SKETCH | ENGINEERING SKETCH |
| `T001` | TECH | WORKSHEET | documents - pipeline & cable / pipeline & cable doc / pipeline | PIPELINE & CABLE | DOCUMENT | GENERAL TECHNICAL DOCUMENT |
| `T010` | TECH | WORKSHEET | documents - naval marine / naval & marine doc / naval marine | NAVAL & MARINE | DOCUMENT | GENERAL TECHNICAL DOCUMENT |
| `T020` | TECH | WORKSHEET | drawings | GENERAL / MULTIDISCIPLINE | DRAWING | ENGINEERING DRAWING |
| `T030` | TECH | WORKSHEET | cut-list / cutlist / cutlists | GENERAL / MULTIDISCIPLINE | DRAWING | CUT LIST |
| `T040` | TECH | WORKSHEET/DOC_NUMBER | tn-pl | PIPELINE & CABLE | DOCUMENT | TECHNICAL NOTE |
| `T041` | TECH | WORKSHEET/DOC_NUMBER | tn-na | NAVAL & MARINE | DOCUMENT | TECHNICAL NOTE |
| `T042` | TECH | WORKSHEET/DOC_NUMBER | tn-st | STRUCTURAL | DOCUMENT | TECHNICAL NOTE |
| `T050` | TECH | WORKSHEET/TITLE | specification | PIPELINE & CABLE | DOCUMENT | SPECIFICATION |
| `T060` | TECH | WORKSHEET/TITLE | anchor pattern | NAVAL & MARINE | DRAWING | ANCHOR PATTERN |
| `T061` | TECH | WORKSHEET/TITLE | dp setup / dp set-up | NAVAL & MARINE | DRAWING | DP SETUP PLAN |
| `T070` | TECH | TITLE | analysis report | KEEP EXISTING | DOCUMENT | ANALYSIS REPORT |
| `T071` | TECH | TITLE | report | KEEP EXISTING | DOCUMENT | REPORT |
| `T072` | TECH | TITLE | calculation / calc | KEEP EXISTING | DOCUMENT | CALCULATION |

`KEEP EXISTING` means a refining rule should preserve already-established higher-level context.

## 9. Rule precedence

Recommended evaluation order:

```text
FILE inclusion/exclusion
  -> Source Family from folder
  -> WORKSHEET exact/alias rule
  -> SECTION rule
  -> HEADER / discipline context
  -> DOC_NUMBER refinement
  -> TITLE refinement
  -> FUZZY fallback
```

Specific rules should override/refine generic rules only in the fields they explicitly set.

## 10. Name normalization and fuzzy matching

Before alias comparison, normalize obvious superficial differences such as:

- case;
- repeated spaces;
- leading/trailing spaces;
- punctuation;
- `&` versus `AND`;
- selected singular/plural variants;
- controlled abbreviations.

Fuzzy matching is a fallback, not the primary mechanism.

Recommended default behavior:

- score >= 90% and only one unambiguous candidate -> automatic classification;
- score 80-89% -> suggested classification + warning/manual review;
- score < 80% -> `UNCLASSIFIED`.

Thresholds must be configurable.

## 11. Safety rules

1. Never silently skip an unknown worksheet.
2. Never force an uncertain worksheet to the nearest category solely to achieve full coverage.
3. Preserve `Original Worksheet` and `Original Section` even after normalization.
4. Every normalized classification must be traceable to a rule ID or explicit fallback reason.
5. Business taxonomy values should come from configuration, not scattered Python constants.
6. User edits to the rule table must be testable without modifying parser code.
7. When a new workbook/sheet cannot be classified safely, output it in classification discovery/warning reports for the user to map.

## 12. Version status

This is **Classification Model v1**, agreed during project planning before implementation.

Cycle 1 profiling must validate this taxonomy against every selected workbook and produce `classification_discovery.csv`. New aliases or subcategories discovered during profiling should be proposed through the GitHub collaboration review process rather than silently invented by the implementer.

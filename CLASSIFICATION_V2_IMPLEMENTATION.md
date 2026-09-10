# Classification Model v2 — Implementation Record

Collaboration-ID: `NMDC-DOC-INDEX-CLASS-V2-001`

This branch implements Classification Model v2 from the approved Cycle-1 evidence. ChatGPT acted as the implementer under explicit owner authorization.

## Scope

- resolve the 34 Cycle-1 `REVIEW_REQUIRED` classification rows by reproducible rules;
- preserve the existing taxonomy unless new evidence requires expansion;
- preserve original worksheet/section evidence;
- prevent title/document-number refinements from crossing established taxonomy branches;
- prevent mixed representative sample rows from collapsing an entire worksheet into one document subtype;
- keep `DATA/` read-only;
- do not begin Cycle 2 extraction;
- do not build the final document index.

## Implemented controls

- whitespace-safe WORKSHEET/SECTION matching;
- Methods installation-procedure structural/header fallback;
- bare Pipeline and Naval Marine aliases;
- safe General/Multidiscipline fallback for generic TECH `Documents`;
- narrowly path-qualified 3291 `CLIENT` duplicate-view exclusion;
- no global CLIENT exclusion;
- external `Requires_Discipline` / `Requires_Category` rule guards for context-sensitive refinements;
- Methods anchor/DP/setup title refinements restricted to `MARINE OPERATIONS -> DRAWING`;
- TECH report/analysis/calculation/procedure/specification/technical-note title refinements restricted to category `DOCUMENT`;
- technical-note document-number refinements restricted to category `DOCUMENT`;
- dedicated Classification v2 regression tests, including semantic cross-category leakage tests;
- cross-platform LF policy for generated text outputs;
- CI checks for real-data profiling, deterministic outputs, current classification gaps, full regression suite, cache hygiene, and `DATA/` integrity.

## Independent review finding: discovery sample leakage

After the first semantic guard pass, independent review found a second class of semantic error.

`classification_rows()` built one worksheet-level evidence string by joining several sampled document titles and document numbers. Per-document rules such as `ANALYSIS REPORT`, `REPORT`, `PROCEDURE`, and `TN-PL` could then fire against that joined sample and assign one document subtype to the whole worksheet. A single sampled Analysis Report could therefore make a mixed `Documents - Pipeline & Cable` worksheet appear to be an `ANALYSIS REPORT` worksheet.

That behavior is unsafe because the discovery layer is profiling worksheet structure, not classifying individual document rows.

## Corrected discovery boundary

Classification v2 now separates two responsibilities:

### Worksheet/section discovery

`classification_rows()` uses only structural evidence:

```text
FILE
WORKSHEET
SECTION
HEADER
```

Sample document numbers and sample titles remain in `classification_discovery.csv` for audit and reviewer visibility, but they are **not** supplied to the rule engine for worksheet-wide classification.

Consequently, mixed document worksheets remain at their safe structural base, for example:

```text
TECH -> PIPELINE & CABLE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
TECH -> NAVAL & MARINE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
TECH -> GENERAL / MULTIDISCIPLINE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
```

### Per-document refinement

`apply_classification()` still supports:

```text
DOC_NUMBER
TITLE
```

Those scopes are retained for a future per-document-row extractor where a title or document number belongs to one specific document record. This PR does not implement that extractor and does not start Cycle 2.

## Verified nonstandard Methods case

The 2171-2172 worksheet with a nonstandard sheet name still needs to resolve as an installation procedure register. The fallback rule `M002` therefore moved from `TITLE` to `HEADER`, using the structural header phrase:

```text
Construction and Installation Procedure
```

This preserves the verified real-data case without allowing arbitrary sampled document titles to classify an entire worksheet.

## Regression coverage

Classification v2 now includes focused tests for:

- all original 34-gap resolution mechanisms;
- scoped 3291 CLIENT handling and unrelated CLIENT safety;
- unknown-future-input behavior;
- preservation of original worksheet evidence;
- backward compatibility of v1 sentinel mappings;
- Methods sketch/incoming-row protection against cross-branch title leakage;
- TECH drawing/sketch protection against document-title leakage;
- technical-note document-number context protection;
- loading of context guards from the external rule table;
- mixed generic TECH titles remaining at the safe discovery base;
- mixed Pipeline titles remaining at the safe discovery base;
- sampled `TN-*` numbers not refining an entire worksheet;
- 2171-2172 discovery resolving from HEADER evidence;
- an unknown worksheet not being classified from a sampled title;
- direct per-document TITLE/DOC_NUMBER refinement remaining available through `apply_classification()`;
- Windows/Linux LF output consistency.

Expected complete suite after this correction: **53 tests** (31 Classification v2 + 20 Cycle-1 profiler + 2 cross-platform output tests).

## Acceptance gate

This implementation is review-ready only when exact-head CI confirms:

1. profiler runs successfully against real `DATA/`;
2. generated outputs are deterministic;
3. current classification `REVIEW_REQUIRED` count is zero;
4. all 53 tests pass;
5. mixed-sample worksheet discovery remains at safe base classifications;
6. no Python runtime artifacts are tracked;
7. `DATA/` is unchanged;
8. Windows output regeneration leaves no line-ending-only dirty files.

Cycle 2 remains unauthorized until owner approval of Classification Model v2.

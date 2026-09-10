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
- bare Pipeline and Naval Marine aliases;
- safe General/Multidiscipline fallback for generic TECH `Documents`;
- narrowly path-qualified 3291 `CLIENT` duplicate-view exclusion;
- no global CLIENT exclusion;
- external `Requires_Discipline` / `Requires_Category` guards for context-sensitive refinements;
- Methods anchor/DP/setup title refinements restricted to `MARINE OPERATIONS -> DRAWING`;
- TECH report/analysis/calculation/procedure/specification/technical-note title refinements restricted to category `DOCUMENT`;
- technical-note document-number refinements restricted to category `DOCUMENT`;
- cross-platform LF policy for generated text outputs;
- CI checks for real-data profiling, deterministic outputs, current classification gaps, full regression suite, cache hygiene, and `DATA/` integrity.

## Independent review finding: discovery sample leakage

Independent review found that `classification_rows()` built worksheet-level `DOC_NUMBER` / `TITLE` evidence by joining several sampled records. Per-document refinement rules could then assign one document subtype to a mixed worksheet. For example, one sampled Analysis Report could make a broad `Documents - Pipeline & Cable` worksheet appear to be an `ANALYSIS REPORT` worksheet.

That is semantically unsafe because Cycle-1 / Classification-v2 discovery profiles worksheet structure; it is not yet extracting individual document records.

## Corrected discovery boundary

### Worksheet/section discovery

`classification_rows()` now supplies only structural evidence to the rule engine:

```text
FILE
WORKSHEET
SECTION
HEADER
```

Sample document numbers and titles remain present in `classification_discovery.csv` for audit, but they are not used as worksheet-wide refinement evidence.

Broad document worksheets therefore stay at safe base classifications such as:

```text
TECH -> PIPELINE & CABLE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
TECH -> NAVAL & MARINE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
TECH -> GENERAL / MULTIDISCIPLINE -> DOCUMENT -> GENERAL TECHNICAL DOCUMENT
```

### Per-document refinement

`apply_classification()` still supports `DOC_NUMBER` and `TITLE`. These scopes are preserved for a future extractor where the evidence belongs to one specific document row. This PR does not implement that extractor and does not start Cycle 2.

## Verified 2171-2172 nonstandard worksheet

The first implementation attempted to preserve the known 2171-2172 procedure register using a HEADER fallback. Exact-head CI showed that the profiler's structural header extraction did not expose the phrase needed for that rule, leaving exactly one real-data `REVIEW_REQUIRED` row.

Rather than reintroduce sampled-title leakage, the fix uses a narrower structural exception:

- exact worksheet: `2171-2172`;
- exact verified workbook path: `DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx`;
- result: `OFFSHORE INSTALLATION -> PROCEDURE -> INSTALLATION PROCEDURE`.

This exception is intentionally path-qualified and regression-tested so a worksheet with the same name in another workbook does not inherit the rule.

## Regression coverage

Classification v2 now covers:

- all original 34-gap resolution mechanisms;
- scoped 3291 CLIENT handling and unrelated CLIENT safety;
- unknown-future-input behavior;
- preservation of original worksheet evidence;
- backward compatibility of v1 sentinel mappings;
- Methods sketch/incoming-row protection against cross-branch title leakage;
- TECH drawing/sketch protection against document-title leakage;
- technical-note document-number context protection;
- mixed generic TECH titles remaining at the safe discovery base;
- mixed Pipeline titles remaining at the safe discovery base;
- sampled `TN-*` numbers not refining an entire worksheet;
- 2171-2172 discovery using the narrow path-qualified worksheet rule;
- the 2171-2172 exception not applying globally;
- unknown discovery rows not being classified from sampled titles;
- direct per-document TITLE/DOC_NUMBER refinement remaining available through `apply_classification()`;
- Windows/Linux LF output consistency.

Complete suite: **54 tests** (32 Classification v2 + 20 Cycle-1 profiler + 2 cross-platform output tests).

## Validation result

The semantic fix was validated on code HEAD `0cead41a6ec46f601241a0be1c5b42e0f71535b3` by GitHub Actions run `34451393027`:

- Python 3.11.16;
- 56 workbooks profiled;
- 211 classification rows;
- classification `REVIEW_REQUIRED`: 0;
- deterministic output comparison: PASS;
- 54/54 tests: PASS;
- mixed-sample discovery regression tests: PASS;
- Python runtime-artifact hygiene: PASS;
- `DATA/` immutability: PASS.

The workflow then regenerated deterministic outputs in commit `a755c538dfc58ddcb14fec59c12c8654cf5633af`. Those outputs show broad Pipeline/Naval/General document worksheets at safe worksheet-level base classifications and preserve per-document refinements for the future row extractor only.

This documentation commit is intentionally made after output regeneration so the final PR HEAD receives a fresh exact-head CI run without changing classification behavior.

## Acceptance gate

This implementation is review-ready only when final exact-head CI confirms:

1. profiler runs successfully against real `DATA/`;
2. generated outputs are deterministic and already current;
3. current classification `REVIEW_REQUIRED` count is zero;
4. all 54 tests pass;
5. mixed-sample worksheet discovery remains at safe base classifications;
6. the 2171-2172 exception resolves only the verified workbook/sheet;
7. no Python runtime artifacts are tracked;
8. `DATA/` is unchanged;
9. Windows output regeneration leaves no line-ending-only dirty files.

Cycle 2 remains unauthorized until owner approval of Classification Model v2.

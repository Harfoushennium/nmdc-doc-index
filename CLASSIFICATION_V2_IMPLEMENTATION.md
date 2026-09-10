# Classification Model v2 — Implementation Record

Collaboration-ID: `NMDC-DOC-INDEX-CLASS-V2-001`

This branch implements Classification Model v2 from the approved Cycle-1 evidence. ChatGPT acted as the implementer under explicit owner authorization.

## Scope

- resolve the 34 Cycle-1 `REVIEW_REQUIRED` classification rows by reproducible rules;
- preserve the existing taxonomy unless new evidence requires expansion;
- preserve original worksheet/section evidence;
- prevent title/document-number refinements from crossing established taxonomy branches;
- keep `DATA/` read-only;
- do not begin Cycle 2 extraction;
- do not build the final document index.

## Implemented controls

- whitespace-safe WORKSHEET/SECTION matching;
- Methods installation-procedure content fallback;
- bare Pipeline and Naval Marine aliases;
- safe General/Multidiscipline fallback for generic TECH `Documents`;
- narrowly path-qualified 3291 `CLIENT` duplicate-view exclusion;
- no global CLIENT exclusion;
- external `Requires_Discipline` / `Requires_Category` rule guards for context-sensitive refinements;
- Methods anchor/DP/setup title refinements restricted to `MARINE OPERATIONS -> DRAWING`;
- TECH report/analysis/calculation/procedure/specification/technical-note title refinements restricted to category `DOCUMENT`;
- technical-note document-number refinements restricted to category `DOCUMENT`;
- Methods installation title fallback restricted so it can resolve the verified nonstandard procedure case without overwriting a sketch/drawing/input classification;
- dedicated Classification v2 regression tests, including semantic cross-category leakage tests;
- CI checks for real-data profiling, deterministic outputs, zero current review-required rows, full regression suite, cache hygiene, and `DATA/` integrity.

## Semantic validation finding and correction

After the first v2 pass reached zero `REVIEW_REQUIRED` rows and passed the initial automated suite, an additional real-data semantic audit found that some rows were confidently classified into the wrong taxonomy branch. The common cause was an unguarded refinement rule applying to representative TITLE evidence after a different base category had already been established.

Observed examples included:

- Methods `Sketches` being refined from `OFFSHORE INSTALLATION -> SKETCH -> ENGINEERING SKETCH` to drawing-only subcategories such as `ANCHOR PATTERN` or `SETUP PLAN`;
- incoming `DOCUMENTS` / `DRAWINGS` section rows being overwritten by unrelated Methods anchor-pattern title text;
- the same rule-engine structure being capable of turning a TECH drawing/sketch into a DOCUMENT subtype from title or document-number evidence.

The fix keeps the taxonomy and rule table external/configurable, but adds explicit context gates. A refinement is now eligible only when any configured `Requires_Discipline` and `Requires_Category` values match the classification already established by earlier evidence.

The corrected real-data output now preserves examples such as:

```text
METHODS -> OFFSHORE INSTALLATION -> SKETCH -> ENGINEERING SKETCH
METHODS -> EXTERNAL / INPUT -> DOCUMENT -> INCOMING TECHNICAL DOCUMENT
METHODS -> EXTERNAL / INPUT -> DRAWING -> INCOMING DRAWING
```

while the verified 2171-2172 title-only procedure case still resolves to:

```text
METHODS -> OFFSHORE INSTALLATION -> PROCEDURE -> INSTALLATION PROCEDURE
```

This demonstrates why `REVIEW_REQUIRED = 0` is necessary but not sufficient; semantic branch-integrity regression tests are also required.

## Regression coverage

Classification v2 adds 25 focused tests on top of the 20 Cycle-1 profiler tests. The v2 tests cover:

- all original 34-gap resolution mechanisms;
- scoped 3291 CLIENT handling and unrelated CLIENT safety;
- unknown-future-input behavior;
- preservation of original worksheet evidence;
- backward compatibility of v1 sentinel mappings;
- Methods sketch/incoming-row protection against drawing-title leakage;
- Methods installation-title fallback protection;
- TECH drawing/sketch protection against document-title leakage;
- technical-note document-number context protection;
- loading of context guards from the external rule table.

Expected complete suite after this correction: **45 tests**.

## Acceptance gate

This implementation is review-ready only when CI passes on the exact final PR HEAD and confirms:

1. profiler runs successfully against real `DATA/`;
2. generated outputs are deterministic;
3. current `REVIEW_REQUIRED` classification count is zero;
4. all 45 legacy + v2 tests pass;
5. semantic cross-category guard tests pass;
6. no Python runtime artifacts are tracked;
7. `DATA/` is unchanged.

Cycle 2 remains unauthorized until owner approval of Classification Model v2.

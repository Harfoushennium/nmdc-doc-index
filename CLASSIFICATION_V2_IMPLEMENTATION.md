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
- strict worksheet-discovery versus per-document-refinement boundary;
- cross-platform LF policy for CSV, JSON and Markdown generated outputs;
- Linux and Windows CI jobs;
- real-data semantic sentinel validation in addition to unit regression tests;
- CI checks for real-data profiling, deterministic outputs, current classification gaps, cache hygiene, `DATA/` integrity, and Windows worktree cleanliness.

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

The first semantic correction attempted to preserve the known 2171-2172 procedure register using a HEADER fallback. Exact-head CI showed that the profiler's structural header extraction did not expose the phrase needed for that rule, leaving exactly one real-data `REVIEW_REQUIRED` row.

Rather than reintroduce sampled-title leakage, the fix uses a narrower structural exception:

- exact worksheet: `2171-2172`;
- exact verified workbook path: `DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx`;
- result: `OFFSHORE INSTALLATION -> PROCEDURE -> INSTALLATION PROCEDURE`.

This exception is intentionally path-qualified and regression-tested so a worksheet with the same name in another workbook does not inherit the rule.

## Cross-platform validation finding

After adding an actual `windows-latest` CI job, the Windows runner exposed a second line-ending defect that the earlier CSV-only test could not see: `workbook_profiles.json` and `source_selection_report.md` were written through `Path.write_text()`, which translated `\n` to CRLF on Windows.

The corrected writer now uses an explicit `newline="\n"` text path for JSON and Markdown as well as the existing CSV `lineterminator="\n"` policy. A regression test writes the real output set into a temporary directory and verifies that JSON and Markdown contain LF only.

The Windows CI job now also requires:

- real profiler execution against `DATA/`;
- zero Classification-v2 `REVIEW_REQUIRED` rows;
- real-data semantic sentinels to pass;
- the full unit suite to pass;
- `DATA/` to remain unchanged;
- `git diff -- outputs/cycle1` to be empty after regeneration;
- final `git status --porcelain` to be clean.

## Real-data semantic sentinel coverage

`tests/verify_real_data_semantics.py` validates the generated real-data discovery output, including:

- all broad Pipeline & Cable document worksheets remain `GENERAL TECHNICAL DOCUMENT` at discovery level;
- all broad Naval & Marine document worksheets remain `GENERAL TECHNICAL DOCUMENT` at discovery level;
- generic TECH `Documents` remains the safe General/Multidiscipline document base;
- METHODS Sketches remain Engineering Sketches;
- METHODS Setup Plans & Anchor Patterns remain `METHOD DRAWING` at worksheet discovery level;
- 2891 incoming DOCUMENTS and DRAWINGS retain their External/Input classifications;
- the 2171-2172 exception resolves through the narrow structural M002 rule;
- the 3291 CLIENT view is excluded only by scoped X017;
- no discovery row uses worksheet-wide TITLE/DOC_NUMBER refinement evidence.

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
- CSV LF output consistency;
- JSON and Markdown LF output consistency.

Complete unit suite: **55 tests** (32 Classification v2 + 20 Cycle-1 profiler + 3 cross-platform output tests), plus the explicit real-data semantic sentinel verifier.

## Current validation result

On code HEAD `d500933ec54a1723acfb89a9670a1cdf25f7cca2`, GitHub Actions run `34452275790` completed successfully on both platforms:

### Linux

- profiler: 56 workbooks / 211 classification rows;
- classification `REVIEW_REQUIRED`: 0;
- deterministic output comparison: PASS;
- real-data semantic sentinels: PASS;
- 55/55 unit tests: PASS;
- Python runtime-artifact hygiene: PASS;
- `DATA/` immutability: PASS.

### Windows

- runner: Microsoft Windows Server 2025;
- Python: 3.11.9;
- profiler: 56 workbooks / 211 classification rows;
- classification `REVIEW_REQUIRED`: 0;
- real-data semantic sentinels: PASS;
- 55/55 unit tests: PASS;
- Python runtime-artifact hygiene: PASS;
- `DATA/` immutability: PASS;
- regenerated `outputs/cycle1`: byte-clean versus committed outputs;
- final working tree: clean.

The semantic verifier checked all 211 discovery rows and specifically confirmed 18 Pipeline broad worksheets, 24 Naval broad worksheets, 2 generic TECH Documents worksheets, 21 METHODS Sketches worksheets, and 19 METHODS setup/anchor worksheets against the safe discovery taxonomy.

## Acceptance gate

This implementation is review-ready only when final exact-head CI confirms:

1. profiler runs successfully against real `DATA/`;
2. generated outputs are deterministic and already current;
3. current classification `REVIEW_REQUIRED` count is zero;
4. all 55 unit tests pass;
5. real-data semantic sentinel verification passes;
6. mixed-sample worksheet discovery remains at safe base classifications;
7. the 2171-2172 exception resolves only the verified workbook/sheet;
8. no Python runtime artifacts are tracked;
9. `DATA/` is unchanged;
10. Windows regeneration leaves `outputs/cycle1` and the worktree clean.

Cycle 2 remains unauthorized until owner approval of Classification Model v2.

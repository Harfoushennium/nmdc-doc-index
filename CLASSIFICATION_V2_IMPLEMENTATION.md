# Classification Model v2 — Implementation Record

Collaboration-ID: `NMDC-DOC-INDEX-CLASS-V2-001`

This branch implements Classification Model v2 from the approved Cycle-1 evidence. ChatGPT acted as the implementer under explicit owner authorization.

## Scope

- resolve the 34 Cycle-1 `REVIEW_REQUIRED` classification rows by reproducible rules;
- preserve the existing taxonomy unless new evidence requires expansion;
- preserve original worksheet/section evidence;
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
- dedicated Classification v2 regression tests;
- CI checks for real-data profiling, deterministic outputs, zero current review-required rows, full regression suite, cache hygiene, and `DATA/` integrity.

## Acceptance gate

This implementation is review-ready only when CI passes on the exact final PR HEAD and confirms:

1. profiler runs successfully against real `DATA/`;
2. generated outputs are deterministic;
3. current `REVIEW_REQUIRED` classification count is zero;
4. all legacy and v2 tests pass;
5. no Python runtime artifacts are tracked;
6. `DATA/` is unchanged.

Cycle 2 remains unauthorized until owner approval of Classification Model v2.

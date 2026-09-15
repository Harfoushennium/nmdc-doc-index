# NMDC Document Index — Release Acceptance Test Plan

**Written by:** ChatGPT  
**Role:** Implementer / Release Coordinator

## Purpose

This plan turns the owner’s reliability findings into hard release gates. The normal Excel product must extract the current known `DATA/` corpus without asking the owner to diagnose parser layouts and without reporting routine source lifecycle notifications as errors.

A **clean current-data Full Rescan** is accepted only when:

- all known readable/selected source workbooks are accounted for;
- data-bearing registers are extracted automatically;
- known empty register templates are recognized as empty rather than reported as errors;
- `Review Flags` contains **0 actionable extraction findings**;
- `Error Log` contains **0 extraction/refresh errors**;
- routine additions/changes remain visible in `Pending Update`/history instead of `Review Flags`;
- source-file and document hyperlinks point to the exact row-specific target;
- the generated workbook opens without the slow-workbook metadata warning, merge warning, or black console window;
- source `DATA/` remains unchanged.

Real future faults such as a locked/corrupt source, an actually ambiguous duplicate, or a non-overridable conflict must still fail safely and remain visible. The product must never hide a genuine data-integrity problem merely to reach zero findings.

## Gate A — Real-data extraction completeness

1. Profile every current workbook under `DATA/` using the same code path as the packaged engine.
2. Verify every selected source is either extracted or intentionally excluded by an explicit rule.
3. Verify the four previously unresolved populated layouts extract automatically:
   - `2035 - UMM Shaif Delivarables.xlsx` / `Setup Plans & Anchor Patterns`;
   - `2745-PP-GE-001-MDR Rev_2.xlsx` / `Anchor Pattern`;
   - `2745-PP-GE-001-MDR Rev_2.xlsx` / `DP`;
   - `7279-OS Document register.xlsx` / `NAVAL MARINE`.
4. Verify the eleven previously flagged but genuinely empty register sheets are classified internally as valid empty registers and do not create Review Flags.
5. Verify the standalone `2. Barges Sketch Deliverables.xlsx` is extracted without inventing a project number; known document `BRG-D4200-001` must be present.
6. Require no runtime parser/layout flags for the current corpus.

## Gate B — Staging and Review Flags

1. Run a brand-new Full Rescan against an empty runtime state.
2. Require staged `flags.json` to be empty for the current known corpus.
3. Require `review_flags = 0`, `blocking_flags = 0`, and staged status `STAGED`.
4. Confirm `SOURCE_NEW`, `SOURCE_CHANGED`, parser/config refresh notices and cache-rebuild notices never appear in Review Flags.
5. Confirm those lifecycle changes remain represented in `Pending Update`, source decisions, dashboard counters and Update History.

## Gate C — Hyperlink integrity

1. Source links must be native row-specific Excel hyperlinks, not ListObject calculated-column `HYPERLINK()` formulas.
2. Excel calculated-column autofill must be disabled while links are installed and restored afterward.
3. Displayed Source File text must remain the source path used for that row.
4. Relative source paths must resolve against the configured DATA folder.
5. Document Link must use the row’s own extracted target.
6. Regression tests must reject reintroduction of formula-array hyperlinks.

## Gate D — Workbook structure and performance

1. All required sheets and uniquely named tables must exist even when empty.
2. No refresh may create/delete temporary worksheets.
3. Home styling must be limited to the visible dashboard range; code must never format `ws.Cells` or otherwise style all 16,384 columns.
4. The source base workbook must not contain a formatted M:XFD tail on Home.
5. Generated workbook must not show Excel’s “99% unused formatting and metadata” warning in owner acceptance testing.
6. Refresh must use manual calculation/events/screen-update suppression and batch typed conversions.

## Gate E — Excel UX and data presentation

1. No merged-cell warning on open.
2. No visible CMD/console window during engine execution.
3. Status/progress feedback remains visible while extraction and table refresh run.
4. Table column guidance and input dropdowns remain present.
5. `Pending Update` is explicitly review-only; decisions belong in `Review Flags` only when a genuine actionable anomaly exists.
6. Empty Error Log/Review Flags tables remain valid Excel tables with headers and filters.
7. Quoted commas/newlines in CSV exchange data must not split one logical row into multiple Excel rows.

## Gate F — Safety and negative tests

1. A locked/unreadable source must produce a clear genuine error and must not silently disappear from the approved index.
2. A failed stage must not change the approved pointer/data.
3. Reset must never delete source DATA, configuration, or audit history.
4. Undo must refuse when no previous approved version exists.
5. Non-overridable conflicts must remain blocking.
6. Source `DATA/` must remain byte-for-byte untouched by profiling/extraction tests.

## Gate G — Determinism and packaging

1. Linux and Windows real-data extraction must be deterministic.
2. Cycle 1 profiler/classification, Cycle 2 sentinel, Cycle 3 full extraction, and Production Excel Package workflows must all pass on the final HEAD.
3. Windows engine build and packaged-engine smoke test must pass.
4. Production ZIP content must be complete and its SHA-256 recorded.
5. Final owner acceptance must use a fresh package and a freshly generated `.xlsm`; old generated workbooks are not valid release evidence.

## Owner acceptance sequence

After all automated gates are green, the owner should:

1. close all source Excel workbooks and wait for OneDrive sync to settle;
2. extract the final package into a new local non-OneDrive test folder;
3. run `Create_NMDC_Document_Index.vbs`;
4. open the fresh workbook and confirm no performance/merge warning appears;
5. select the real DATA folder and run **Full Rescan / Rebuild All**;
6. confirm Review Flags is empty and Error Log contains no extraction/refresh error;
7. sample Source File and Document Link hyperlinks across multiple projects and confirm every link opens the correct file;
8. verify Pending Update remains a review-only change preview;
9. test Reset All Records and rerun Full Rescan;
10. do **not** approve the staged update until these checks are accepted.

## Release rule

PR #8 stays **CHANGES** and unmerged until the owner confirms the acceptance sequence. Only the owner may authorize merge.

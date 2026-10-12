# Owner controls and performance: architecture review

**Scope:** survey implementation baseline PR #8 at `846f5ab4d80f41554fe5b609ca1c201da18633e0`. Measurements below describe that baseline. The owner subsequently selected candidate 1 and approved the five slices and its test seams on 2026-10-11. [Plan #21](https://github.com/Harfoushennium/nmdc-doc-index/issues/21) and [spec #18](https://github.com/Harfoushennium/nmdc-doc-index/issues/18) govern implementation. `GLOSSARY.md` now records the agreed review terminology; there were no ADRs when surveyed. Other architecture alternatives remain proposals.

## Confirmed control semantics

`Pending Update` source checkboxes select whole workbooks. `NMDC_SaveSourceSelections` persists the batch, and the next stage applies exclusions. Review Flags decisions are per flag, but `apply_review_decisions` only records a decision/comment/status in staged flags and history. `build_pending_rows` emits a review summary; `approve_stage` promotes the whole staged package by run ID. **There is no partial-row approval.** The “User Decisions” table is described as an audit trail. The existing code therefore supports file-level source selection and whole-proposal approval; it does not provide record-level inclusion or correction before approval.

A disposable COM probe confirmed a scope-loss defect: a locally unchecked source and its owner note reverted to checked/blank when `NMDC_RebuildPendingSourcePanel` ran (8.961 s); no approval occurred. The owner review draft is the strongest first candidate.

## Deepening candidates

1. **Protect owner review drafts — Strong.** Files: `excel/vba/modNMDC_Checkboxes.bas`, `excel/vba/modNMDC_Performance.bas`, `nmdc_profiler/source_selection.py`, `excel/vba/modNMDC_Actions.bas`. Refresh rebuilds the panel from engine-exported selection state, so unsaved checkbox and note edits can disappear; approval refreshes before the run check. Keep owner edits recoverable through refresh and make approval detect an unsaved scope change. This increases locality for draft behavior and gives Excel regression tests one review path to exercise. Regression seam: dirty checkbox/note → refresh/reopen/approve attempt; confirm draft retained or approval stopped until restaged. Partial-row approval remains an owner decision.

2. **Concentrate operation ordering and refresh work — Worth exploring.** Files: `modNMDC_Actions.bas`, `modNMDC_Performance.bas`, `modNMDC_Admin.bas`, `modNMDC_Refresh.bas`, `nmdc_profiler/runtime_admin.py`. Sync and async flows coexist; approve exports/refreshes before and after promotion; reset deletes state, then exports and refreshes. A shared operation module could own sequencing and scope table refreshes to what the current action needs. This improves locality and makes timing regressions testable. Regression seam: disposable Excel operation matrix for update, review, approval, reset, errors, and preserved approved pointer.

3. **Deepen live search — Worth exploring.** Files: `modNMDC_LiveFilter.bas`, `Cls_LiveFilter_Listener.cls`. Each textbox change calls the filter; multi-term/exclusion/exact paths inspect visible cells and union rows. Separate query matching from Excel visibility work, then batch table reads/writes. This concentrates query semantics and enables fast fixture-level tests; Excel tests still verify visibility. Regression seam: exact, AND, exclusion, empty/reset, target switching, and 40k-row timing.

4. **Indexed engine storage — Speculative.** Consider SQLite only if profiling shows repeated structured queries dominate. It might help indexed lookups, but cannot be assumed to speed Excel table import or source extraction; it adds schema, migration, and package behavior to own. Benchmark against the current JSON/CSV path before choosing. [SQLite query planning](https://www.sqlite.org/queryplanner.html) supports the indexing premise; applicability here is inference.

## Evidence and tracer slices

On disposable COM probes, full refresh was 57.204 s; a separate probe loaded the 40,494-row Pending Update table (parse/format/links) in 97.330 s, with CSV parse 6.659 s, workbook guidance 2.844 s, and source panel 7.446 s. A cProfile saved-stage export was 1.012 s. Omitting only source-link creation from the real loader in another disposable trial reduced load to 72.557 s. That probe violates the hyperlink requirement and is diagnostic only; one trial cannot quantify a reliable improvement. Reset and extraction durations remain unmeasured. These separate runs are not additive or directly comparable.

Direct filter calls took 0.063 s for a no-match query, 0.133 s for an exclusion matching no rows, and 0.041 s to clear. These are limited timing probes; they do not assert query correctness or establish typing latency, complex rejection/exact/AND behavior, or physical input acceptance. Sixteen existing update-engine tests passed in 0.515 s. Microsoft recommends block range reads/writes ([official guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)); applicability here is inference.

Owner-approved slices, published with native sub-issue and blocking relationships under #21:

1. **P0 #18 owner control:** retain checkbox choices/notes and flag inputs through refresh/reopen; stop approval until scope is saved/restaged and decisions saved. Retain old-run drafts for visible review. **Blocked by:** none. Related #10/#17; partial-record approval is outside this slice.
2. **P1 #19 refresh:** demonstrate faster repeat review/update with unchanged exports skipped, links preserved and phase timings. **Blocked by:** #18 before integrating new refresh paths.
3. **P1 #12 reset:** time reset/export/Excel clearing independently, improve the measured slow step, preserve DATA/configuration/audit and demonstrate a successful rescan. **Blocked by:** none for diagnosis.
4. **P2 #20 search:** prove AND/exact/exclusion/reset and both column selectors in Excel, then improve only reproduced delays. **Blocked by:** none for diagnosis. Related #16 and acceptance case 11.
5. **#15 Release verification:** complete all 23 cases on two fresh installations and retained owner settings. **Blocked by:** #18/#19/#12/#20 and preserved earlier #10/#11/#13/#14 gates, plus working desktop input. PR #8 stays CHANGES until acceptance.

Owner accepted the real Excel edit → refresh → save/reopen → approval attempt seam, with engine exclusion/approved-pointer checks. Full physical input, 23-case and two-installation acceptance remain pending. Current measurements do not establish owner acceptance.

## P0 implementation evidence (2026-10-12)

The working candidate now journals source checkbox choices, notes and Review Flag inputs by proposal and stable row identity. Refresh captures edits before rebuilding its panel; old proposal drafts remain visible and are not applied to replacement proposals. Approval is blocked while source changes await save/restage or Review Flag decisions await save. Workbook setup also keeps the native 15-worksheet / 15-Excel-table contract; data regions continue to use named `ListObject` tables.

The maintained disposable Excel verifier passed **28 checks** against package `outputs/current-native-package/20261011-211029` (workbook SHA-256 `6bf17136ea7e633e35927dcf44140e73a8abdb06614bb3d3137d25063fdf187c`). This includes refresh/reopen retention, row reordering and table growth, replacement proposals, waiting-restage guards, malformed draft rejection, approved-state preservation, and all 15 named worksheet tables. The actual approval action reached its guard before export/promotion; only its notification was adapted to a captured cell in the disposable test, so physical dialog interaction remains unproven. A separate native 40,494-row control probe passed 11 checks for clean/dirty/reverted/restaged source choices and the 15-sheet/15-table contract. It did not approve or write DATA.

The standard-library suite passes **269 tests** (`uv run python -m unittest discover -s tests -p 'test_*.py'`). A real-engine CLI scope integration test proves that excluding one source removes its records from the next proposal while the approved pointer/records and both source workbooks remain byte-identical. The 15-table user requirement is already explicit in `USER_PRODUCT_REQUIREMENTS.md` and `PROJECT_SPEC.md`, and is now included in this P0 verification contract.

These checks use the uncommitted working candidate and the PR #8 baseline engine for disposable UI setup. They are not yet GitHub CI evidence or owner acceptance. Physical Excel interaction, final-artifact verification, the other approved fixes, and the 23-case/two-installation release gates remain open.

## Reset clear timing and candidate fix (2026-10-12)

Two disposable 40,494-row Reset trials measured engine reset at 0.587/0.635 s and export at 0.532/0.578 s, while the Excel refresh/clear took 103.373/135.281 s. A phase probe attributed 125.453 s to `DataBodyRange.Hyperlinks.Delete` over the full Pending Update table; `ClearContents` took 0.070 s. Deleting hyperlinks from only the generated link columns still took 99.412 s because the source column held tens of thousands of native links.

The working fix uses Excel's bulk `Range.Clear` on generated hyperlink columns (`Document Link`, `Source File`, `relative_path`, `Relative Path`), then clears the remaining data body as before. A disposable COM probe confirmed the bulk clear removed the hyperlink objects in 0.148 s, preserved the ListObject and all 40,494 rows, and a subsequent refresh restored a source link matching its row. Two independent full Reset trials on the candidate measured Excel clear/refresh at 1.692 and 1.751 s; engine reset/export remained about 0.6 s each. Both retained 15 worksheets/15 named tables and zero nonblank Pending Update rows. These are disposable package/runtime trials, not the two-fresh-install owner acceptance gate. Verify table formatting, validation and source/document links again on the exact CI artifact before considering the performance slice complete.

## Live search rejection timing (2026-10-12)

A 40,494-row Excel probe confirmed visible row IDs for no-match (0.048 s), broad AND `OFO+SK` (0.067 s), exact case-sensitive (1.297 s), and clear (0.432 s). Rejection queries that hide 12,742 rows exposed the implementation cost: building a growing `Union` did not complete within about 3.5 minutes, so the owned disposable Excel process was stopped. Replacing the growing union with direct per-row visibility changes restored exact row IDs in 3.994 s for `OFO -SK` (1 row visible) and 4.725 s for exclusion-only `-SK` (27,750 visible). Eight controlled seven-row semantics checks also passed, including successive broader queries and clear after exact/exclusion. The direct hide path is a candidate correction; its 4–5 second worst-case still needs an owner latency target and the exact CI artifact retest. Physical textbox typing, caret focus, CHANGE COLUMN button and Ctrl+Shift+F remain unverified.

## Guided Matt workflow

Use `diagnosing-bugs` for a failing real-call loop; `improve-codebase-architecture` for candidate comparison; `grill-with-docs` for the selected decision; `to-spec` for behavior and test seams; `to-tickets` for approved vertical slices; `tdd` after agreeing the behavioral test seams; `code-review` for standards/spec checks at a fixed revision. Invoke by name, for example `Use the diagnosing-bugs skill for unsaved source-choice loss`. Reserve Luna for bounded implementation and keep one desktop operator. GitHub remains authoritative.

Evidence files are ignored local artifacts: `outputs/architecture-probes/controls-20261011-195112/report.json`, `timing-20261011-195221/report.json`, `links-off-20261011-195553/report.json` and `engine-export/report.json` under that same folder. No owner-root changes, restage, approval or DATA writes were performed.

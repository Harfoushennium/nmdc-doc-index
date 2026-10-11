# Owner controls and performance: architecture review

**Scope:** implementation baseline PR #8 at `846f5ab4d80f41554fe5b609ca1c201da18633e0`. Survey and disposable experiments; no product code or architecture changed. `GLOSSARY.md` and `docs/adr/` are absent. Candidate descriptions do not define interfaces or approve implementation.

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

Draft tickets, not published or approved:

1. **P0 owner control:** retain checkbox choices/notes through refresh; stop approval of an unsaved scope change until it is saved/restaged. **Blocked by:** none for existing file-level behavior; individual-record policy needs owner clarification. Covers #10/#17.
2. **P1 refresh:** demonstrate faster repeat review/update with unchanged exports skipped, links preserved and phase timings. **Blocked by:** 1 before integrating new refresh paths.
3. **P1 reset:** time reset/export/Excel clearing independently, improve the measured slow step, preserve DATA/configuration/audit and demonstrate a successful rescan. **Blocked by:** none for diagnosis. Covers #12.
4. **P2 search:** prove AND/exact/exclusion/reset and both column selectors in Excel, then improve only reproduced delays. **Blocked by:** none for diagnosis. Covers #16 and acceptance case 11.
5. **Release verification:** complete all 23 cases on two fresh installations and retained owner settings. **Blocked by:** 1–4 plus working desktop input. PR #8 stays CHANGES until acceptance.

Acceptance dependencies: clarify whether decisions should ever exclude individual records; restore an interactive desktop for physical input checks; complete the remaining 23-case and two-installation gates. Current UX and artifact measurements do not establish owner acceptance.

## Guided Matt workflow

Use `diagnosing-bugs` for a failing real-call loop; `improve-codebase-architecture` for candidate comparison; `grill-with-docs` for the selected decision; `to-spec` for behavior and test seams; `to-tickets` for approved vertical slices; `tdd` after agreeing the behavioral test seams; `code-review` for standards/spec checks at a fixed revision. Invoke by name, for example `Use the diagnosing-bugs skill for unsaved source-choice loss`. Reserve Luna for bounded implementation and keep one desktop operator. GitHub remains authoritative.

Evidence files are ignored local artifacts: `outputs/architecture-probes/controls-20261011-195112/report.json`, `timing-20261011-195221/report.json`, `links-off-20261011-195553/report.json` and `engine-export/report.json` under that same folder. No owner-root changes, restage, approval or DATA writes were performed.

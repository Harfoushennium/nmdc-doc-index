# NMDC Document Index — Release Acceptance Test Plan

> **Current owner stabilization override (13 Oct 2026):** PR #8 remains unaccepted. File Review is the primary affected-file acceptance worksheet; Change Details is optional. Unchecked files retain their approved data for this proposal and never change permanent exclusions. Native Excel simulation and the two-install gate remain mandatory.

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
- routine additions/changes remain visible in `Change Details`/history instead of `Review Flags`;
- source-file and document hyperlinks point to the exact row-specific target;
- the generated workbook opens without the slow-workbook metadata warning, merge warning, or black console window;
- source `DATA/` remains unchanged.

Real future faults such as a persistently locked/corrupt source, an actually ambiguous duplicate, or a non-overridable conflict must still fail safely and remain visible. The product must never hide a genuine data-integrity problem merely to reach zero findings.

## Gate A — Real-data extraction completeness

1. Profile every current workbook under `DATA/` using the same code path as the packaged engine.
2. Verify every selected source is either extracted or intentionally excluded by an explicit rule/owner source-scope decision.
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
5. Confirm those lifecycle changes remain represented in `Change Details`, source decisions, dashboard counters and Update History.

## Gate C — Hyperlink integrity

1. Source links must be native row-specific Excel hyperlinks, not ListObject calculated-column `HYPERLINK()` formulas.
2. Excel calculated-column autofill must be disabled while links are installed and restored afterward.
3. Displayed Source File text must remain the source path used for that row.
4. Relative source paths must resolve against the configured DATA folder.
5. Document Link must use the row’s own extracted target.
6. Regression tests must reject reintroduction of formula-array hyperlinks.

## Gate D — Workbook structure and performance

1. All required sheets and uniquely named tables must exist even when empty, including all user-facing review/data tables. A legacy/internal `SourceSelection` table may exist only for backward compatibility; there must be no separate owner-facing Source Selection worksheet.
2. No refresh may create/delete temporary worksheets.
3. Home styling must be limited to the visible dashboard range; code must never format `ws.Cells` or otherwise style all 16,384 columns.
4. The source base workbook must not contain a formatted M:XFD tail on Home.
5. Generated workbook must not show Excel’s “99% unused formatting and metadata” warning in owner acceptance testing.
6. Refresh must use manual calculation/events/screen-update suppression and batch typed conversions.
7. Active filters and sort fields must be cleared before table resize/repopulation so stale hidden/sorted row state cannot mix refreshed data.
8. Normal **Update Changed Files** must use the persistent local source cache and must not reread/reprocess unchanged OneDrive workbooks.
9. Runtime/state/cache must use the package-local `runtime` folder. NMDC runtime/support data must not be configured under AppData or another unrelated user folder.
10. Long scans must use the responsive/background VBA workflow and provide status feedback without a visible command window.

## Gate E — Excel UX and data presentation

1. No merged-cell warning on open.
2. No visible CMD/console window during engine execution.
3. Status/progress feedback remains visible while extraction and table refresh run.
4. Every action on Home has plain-English guidance describing when to use it and what it changes.
5. File Review is the normal affected-file review; Change Details is a separate optional record preview. Review Flags is reserved for actionable anomalies.
6. All worksheet data areas, including File Review and Change Details, are valid Excel Tables even when empty.
7. Quoted commas/newlines in CSV exchange data must not split one logical row into multiple Excel rows.
8. Body rows use professional default presentation: Aptos 10, automatic font color, no body fill, consistent alignment, sensible widths/heights, light borders and wrapping only where useful.
9. Small tables AutoFit row height within safe limits; large extraction tables use compact stable row height to preserve performance.
10. Business/review columns appear before technical keys according to the owner’s review logic.

## Gate F — Rules & Mappings non-coder UX

1. Default view must be understandable without knowledge of regex, parser internals or configuration file formats.
2. Normal user controls must include Add Simple Rule and Show/Hide Advanced.
3. Advanced technical columns are hidden by default.
4. Normal choices use dropdowns: Enabled, Source Family, Match Scope, Match Type, Include and Stop on Match.
5. Shipped/default rules use plain matching only: `CONTAINS`, `EXACT`, `STARTS_WITH`, `ENDS_WITH`; no shipped rule requires REGEX.
6. Add Simple Rule generates a safe Rule ID / Priority and sensible defaults.
7. Discipline, Category and Subcategory are clearly described as the classification outputs the rule assigns.
8. Invalid/duplicate rules must fail validation before staging; they must never be silently accepted.

## Gate G — Proposal file acceptance

1. File Review shows only affected files, with Project No., Source File, Change Type, added/modified/removed counts and Owner Note.
2. Native Accept Update? checkboxes are required; creation failure is a release-blocking UI error and must not silently downgrade to TRUE/FALSE text. Checked accepts the file update; unchecked retains previously approved data for this proposal.
3. Checkbox clicks never scan or approve. Save File Choices binds choices to the exact displayed run and approved baseline without restaging.
4. Approval is blocked until file choices and Review Flag decision/comment/status inputs are saved. Unresolved conflicts remain subject to existing conflict guards.
5. Mixed acceptance retains deferred file rows, manifest entries and caches. The next scan must still offer deferred changes.
6. Deferred new files contribute no records; deferred removed files retain their approved data. All-deferred approval leaves approved records unchanged.
7. Saving/reloading, sorting and save/reopen preserve choices and notes by file identity. Replacement runs preserve prior drafts without cross-applying them.
8. View Change Details opens a separate optional Table; file approval does not require viewing individual record changes.
9. Both review worksheets retain independent Excel Tables and native controls through refresh/table growth.
10. Existing permanent exclusions remain excluded; File Review choices do not edit config/source_exclusions.csv or DATA.
11. Engine regression tests prove mixed/deferred/stale-run safety. Real Excel acceptance proves typing, buttons, checkbox rendering, notes, save/reopen and actual dialogs.

## Gate G2 — Review Flags selection checkbox workflow

1. `Review Flags` must begin with a native Microsoft 365 **Select?** checkbox column.
2. The owner can tick one or multiple exact parser/extraction issues and use **Report Selected Parser Fix**.
3. **Select All** and **Clear Selection** must operate on the Review Flags selection column only.
4. A failure to create Review Flags native checkboxes is release-blocking and must show a visible error; TRUE/FALSE text is not an accepted production substitute.
5. Review Flags must display **Source File** and a separate explicit **Worksheet Name** column.
6. Reporting selected flags must create `PARSER_FIX_REPORTS` beside the XLSM, then allocate the next sequential subfolder (`0001`, `0002`, `0003`, ...).
7. Each numbered request folder must contain only `PARSER_FIX_REQUEST.md` and `PARSER_FIX_REQUEST.json` for that request. No parser-report files may be written loose in the project root.
8. Windows Explorer must open to the exact numbered request after creation; a later request must never overwrite or mix with an earlier request.

## Gate H — Dynamic Live Filter

1. Live Filter must be available on table-heavy sheets that benefit from rapid search: Master Documents, Revisions, Transactions, Change Details, Review Flags, User Decisions, Update History and Error Log.
2. The search input must be a **real editable Excel input**: the user can click inside it, see a caret, type normally, use Backspace normally, and never enter a global keyboard-capture mode.
3. Filtering updates on **every keystroke** without Enter/Tab.
4. The user explicitly chooses one target table column and the selected column is clearly displayed.
5. The implementation must first inspect and preserve the search semantics/help from the owner's original Dynamic Live Filter reference module.
6. Test the behavior actually present in the owner REV03 reference: normal partial matching, multiple required terms/AND behavior using spaces or `+`, excluded words using `-`, and quoted exact-match behavior. Do not add a different search grammar and call it equivalent.
7. Help/placeholder text must show the usable syntax with examples.
8. The search implementation must work in the owner's corporate Microsoft 365 environment and must not reproduce runtime error 40040.
8a. The generated XLSM must contain a valid **Microsoft Forms 2.0 Object Library (MSForms)** project reference before `Cls_LiveFilter_Listener` is compiled. A `User-defined type not defined` error on `MSForms.TextBox` is a release-blocking setup failure.
8b. Production setup runs Excel invisibly under COM automation and must **not create worksheet ActiveX controls during setup**. REV03 `TxtBox_Search` controls are created lazily only when a supported worksheet is activated in normal visible Excel. Any setup-time COM/RPC failure while initializing Live Filter is release-blocking.
9. The owner REV03 architecture is required: a real worksheet ActiveX `TxtBox_Search` with a listener-class `Change` event. `Application.OnKey` is permitted only for the `Ctrl+Shift+F` target-column shortcut; it must never capture normal typing.
10. Clearing/resetting Live Filter must never merge populated ranges, damage data, or disturb unrelated table filters.
11. Live Filter must survive refresh, close/reopen and table rebuilds without manual VBA repair.
12. Real Excel simulation is mandatory; static VBA string-contract tests are not sufficient evidence.

### Review Flag parser/mapping repair handoff

A genuine `UNRECOGNIZED_LAYOUT` or equivalent extraction flag must not end in a dead audit choice.

1. Selecting **NEEDS PARSER/MAPPING FIX** must preserve the affected source file, explicit Worksheet Name, flag code, project/document context and owner comment.
2. **Request Parser / Mapping Fix** must write the request into its own sequential `PARSER_FIX_REPORTS\NNNN` folder beside the workbook package; it must not write loose report files in the project root or hide them under runtime/AppData.
3. The workbook must state clearly that the installed package does not rewrite its own parser executable automatically.
4. Approved data and source `DATA/` must remain unchanged while the issue is unresolved.
5. After corrected code/configuration is installed, **Retry After Fix** must run a Full Rescan so unchanged flagged sources are actually reprocessed.
6. The flag is considered resolved only when the corrected extraction no longer reproduces it; saving a decision alone is not resolution.

## Gate I — User Custom Fields & Keyword Mappings

1. A dedicated `Custom Fields` workspace must exist with two valid Excel tables: `CustomFields` and `KeywordMappings`, even when no rules have been entered.
2. `CustomFields` defines user-derived Master Documents columns with: Enabled?, Field Name, Search In, Match Behavior, Separator and Notes.
3. `KeywordMappings` defines the owner-managed dictionary with: Enabled?, Field Name, Keyword / Pattern, Result, Match Type, Priority and Notes.
4. Enabled? is a true binary choice and therefore uses Excel checkboxes. Match Behavior / Match Type remain dropdowns because they have several meanings.
5. A user can create a field such as **Vessel Names** without changing Python or VBA.
6. Search In can name one or several Master Documents fields (for example `Document Title;Source File`) or `ALL TEXT`.
7. `CONTAINS` is the simplest/default mapping method.
8. `ALL TERMS` supports the same `+AND / -EXCLUDE` idea as Live Filter.
9. `EXACT` supports whole-text matching.
10. `WILDCARD` provides an advanced adaptation of the owner’s keyword tool: `?` fixed-width extraction and `*` variable-width extraction where applicable.
11. `FIRST` returns the first enabled mapping by priority; `ALL UNIQUE` collects every unique match in priority order with the configured separator.
12. Core NMDC fields (Project No., Document No., Document Title, classifications, links, source identities, keys, etc.) are reserved and cannot be overwritten by Custom Fields.
13. Any internal field beginning `__NMDC_` is also reserved.
14. Custom fields are a presentation/enrichment layer only; they must not mutate canonical approved/staged engine records or source workbooks.
15. The custom definitions/mappings must be backed up to the local runtime as `user_custom_fields.csv` and `user_keyword_mappings.csv` so a fresh generated workbook can restore the owner’s definitions.
16. After a core CSV refresh trims Master Documents back to canonical columns, enabled custom columns must be restored automatically when Master Documents is next activated/used.
17. A user may either use **Add Custom Field** or add a non-reserved column in the Master Documents table manually; a matching enabled CustomFields definition must be able to populate it.
18. Automated regression tests must cover the workspace contract, checkbox policy, match modes, core-field protection, runtime backup names, bootstrap packaging and post-refresh restoration guard.

## Gate J — Source access resilience

1. Valid Excel/OneDrive source workbooks that return a transient Windows sharing/permission error must be retried automatically before the scan fails.
2. Retry handling must cover both profiling and content hashing.
3. If direct access continues to fail, the engine should attempt a temporary read-only snapshot before declaring the source unavailable.
4. A persistent access failure must become one clear `SOURCE_HASH_ERROR`/blocking source-access finding instead of crashing the packaged engine with exit code 2.
5. A persistently unavailable source must never be treated as removed from the approved index.
6. Parser/cache compatibility must be versioned. When parser behavior changes in a way that can alter extraction or Review Flags, the shared parser version must change so unchanged sources cannot silently reuse stale extraction caches.
6. The recommended action must tell the user to close the source workbook if open and wait for OneDrive synchronization before retrying.
7. Regression tests must verify transient PermissionError recovery and the real Project 2369 source remains readable in the repository baseline.

## Gate K — Safety and negative tests

1. A locked/unreadable source must produce a clear genuine error and must not silently disappear from the approved index.
2. A failed stage must not change the approved pointer/data.
3. Reset must never delete source DATA, configuration, or audit history.
4. Undo must refuse when no previous approved version exists.
5. Non-overridable conflicts must remain blocking.
6. Source `DATA/` must remain byte-for-byte untouched by profiling/extraction tests.

## Gate L — Determinism and packaging

1. Linux and Windows real-data extraction must be deterministic.
2. Cycle 1 profiler/classification, Cycle 2 sentinel, Cycle 3 full extraction, and Production Excel Package workflows must all pass on the final production-code HEAD.
3. Windows engine build and packaged-engine smoke test must pass.
4. Production ZIP must contain the complete VBA set, including `modNMDC_Checkboxes.bas`, `modNMDC_LiveFilter.bas`, `modNMDC_CustomFields.bas` and `modNMDC_CustomFieldsSetup.bas`, plus source-selection configuration and exchange support.
5. Production ZIP content must be complete and its SHA-256 recorded.
6. Final owner acceptance must use a fresh package and a freshly generated `.xlsm`; old generated workbooks are not valid release evidence.

## Owner acceptance sequence

After all automated gates are green, the owner should:

1. close all source Excel workbooks and wait for OneDrive sync to settle;
2. extract the final package into a new test folder;
3. run `Create_NMDC_Document_Index.vbs`;
4. open the fresh workbook and confirm no performance/merge warning appears;
5. select the real DATA folder and run **Full Rescan / Rebuild All**;
6. confirm Review Flags is empty and Error Log contains no extraction/refresh error;
7. sample Source File and Document Link hyperlinks across multiple projects and confirm every link opens the correct file;
8. verify Change Details remains a review-only change preview;
9. open File Review; uncheck an affected disposable fixture file, add a note, save choices and reopen; verify both inputs. Before approval verify approved data is unchanged. In the disposable approval fixture, verify the unchecked file keeps its approved contents.
10. stage another proposal and confirm deferred file changes are still offered; verify prior drafts remain reviewable without cross-application.
11. test Check All and cancel/confirm Uncheck All in File Review. Verify that unsaved choices and unsaved review decisions block approval.
12. on Master Documents, press Ctrl+Shift+F and select the Document Title table header; confirm the placeholder names Document Title; then type directly in `TxtBox_Search` using one partial term, two required terms (space/`+`), one `-EXCLUDE` term and a quoted exact value; use **RESET SEARCH** afterward;
13. open **Custom Fields & Keywords**, add a temporary `Vessel Names` field searching `Document Title;Source File`, add at least two vessel keyword mappings, and apply them; confirm the new column is populated only where the configured keywords match;
14. test both `FIRST` and `ALL UNIQUE`; test one `ALL TERMS` expression with `+`/space AND and `-EXCLUDE`; test wildcard mode only with a harmless sample rule where the expected result is obvious;
15. run an update/refresh and return to Master Documents; confirm `Vessel Names` is restored/repopulated and the Live Filter remains usable without manual VBA repair;
16. disable one keyword mapping using its checkbox and re-apply; confirm that disabled mapping no longer contributes to output;
17. confirm a Custom Field named like a core field (for example `Document Title`) is rejected and the core data remains unchanged;
18. apply a normal table filter/sort, refresh, and confirm the refreshed table resets cleanly without mixed rows;
19. confirm Rules & Mappings is synchronized from the packaged current `classification_rules.csv`; specifically verify the 2171-2172 Methods rule and 3291 CLIENT exclusion use plain path qualifiers rather than legacy regex path text;
20. test Reset All Records and rerun Full Rescan;
21. **do not approve the staged update** until these checks are accepted.

## Release rule

PR #8 stays **CHANGES** and unmerged until the owner confirms the acceptance sequence. Only the owner may authorize merge.

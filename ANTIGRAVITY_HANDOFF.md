# NMDC Document Index — Antigravity IDE Full Stabilization & Simulation Handoff

**Written by:** ChatGPT
**Role:** Repository Cleanup / Handoff Coordinator
**Date:** 19 Sep 2026

## 1. Mission

Take over PR #8 and continue until the NMDC Document Index is proven to be a **working Microsoft Excel desktop product**.

Do not stop at planning, code review, unit tests, or green GitHub Actions. The final gate is a complete simulated owner workflow in real Microsoft Excel, repeated on a second clean installation.

Required loop:

**inspect → reproduce → root-cause → implement → simulate in real Excel → fix → regress → package → independently inspect → deliver**

Do not return only a plan.

---

## 2. Repository / authority

- Repository: `Harfoushennium/nmdc-doc-index`
- Active PR: `#8 — [08][CHANGES] Excel User Interface & Engine Integration`
- Branch: `08-excel-user-interface`
- Repository was cleaned for this handoff on commit `d98ee4c4bb28953f4f59dfe78c8907af1ab67c31`.
- Always live-read PR #8 and use the actual current HEAD after pulling.
- Owner is the **only final merge authority**.

Mandatory governance:

- DO NOT merge PR #8.
- DO NOT approve PR #8.
- DO NOT modify `master` directly.
- DO NOT force-push.
- DO NOT reset/rebase away existing work.
- DO NOT delete unrelated current project files.
- Keep PR #8 in `CHANGES` until owner acceptance.
- GitHub is the source of truth.
- Important comments must include:
  - `Written by: Antigravity IDE`
  - `Role: Implementer / Test Engineer`
- Update collaboration dashboard comment `5637853813` **in place**.

---

## 3. Clean-repository rules

The branch has already been cleaned of obsolete historical cycle plans, duplicated owner-review workbooks, old pilot validation packages, and obsolete one-off workflows.

Do **not** restore those deleted artifacts just because they exist in git history.

Canonical project areas to keep/use:

- `DATA/` — authoritative read-only real-data regression corpus.
- `config/` — editable deterministic configuration.
- `excel/` — workbook contract and VBA source.
- `nmdc_profiler/` — profiler/extractor/update-engine implementation.
- `packaging/` — Windows production package builder.
- `tests/` — active regression tests; add real-Excel E2E tests here.
- `tools/` — only current diagnostic/verification tools.
- `outputs/` — deterministic validation outputs/baselines used by CI.
- `.github/workflows/` — active profiler, Cycle 2, Cycle 3 and production-package validation.
- `README.md`, `PROJECT_SPEC.md`, `USER_PRODUCT_REQUIREMENTS.md`, `CLASSIFICATION_MODEL.md`, `RELEASE_ACCEPTANCE_TEST_PLAN.md`, `AGENTS.md` and this handoff.

Use a **local non-OneDrive workspace** for development and repeated simulation.

The owner's normal project folder is:

`C:\0\OneDrive - NMDC Group\Desktop\NPCC\AI PROJECTS\NMDC DOCUMENTS INDEX`

Recommended scratch/test location:

`C:\NMDC_Index_Antigravity_Test`

or `%LOCALAPPDATA%\NMDC Document Index\Antigravity Test`.

`DATA/` is READ ONLY. Never modify, repair, resave, rename, delete or normalize source workbooks.

---

## 4. Required reading order

Before changing implementation:

1. `README.md`
2. `ANTIGRAVITY_HANDOFF.md`
3. `PROJECT_SPEC.md`
4. `USER_PRODUCT_REQUIREMENTS.md`
5. `CLASSIFICATION_MODEL.md`
6. `RELEASE_ACCEPTANCE_TEST_PLAN.md`
7. `excel/WORKBOOK_UI_SPEC.md`
8. current PR #8 description and complete discussion/dashboard
9. actual current repository diff/tree

Where older implementation detail conflicts with the owner's latest requirements in this handoff, this handoff and the live PR discussion govern.

---

## 5. Forensic owner workbook

Primary failing workbook supplied by owner:

`NMDC_Document_Index(5).xlsm`

If available locally, preserve it unchanged and use copies for diagnosis.

Known owner screenshots/failures from this workbook are release blockers.

---

## 6. Stop the patch loop

Previous cycles repeatedly produced this pattern:

- patch one defect;
- lose or break another feature;
- static/CI tests pass;
- real Excel still fails.

Before implementing, create a **root-cause / regression map** with:

| Feature | Last owner-confirmed working revision | Current behaviour | Regression/root cause | Fix strategy | Regression test |
|---|---|---|---|---|---|

At minimum cover:

- setup/startup;
- Full Rescan;
- Update Changed Files;
- Reset;
- Undo;
- Home buttons;
- Pending Update;
- source include/exclude;
- Review Flags;
- hyperlinks;
- Live Filter;
- Rules & Mappings;
- Custom Fields;
- formatting;
- close/reopen;
- OneDrive/runtime behaviour.

Restore stable architecture first, then layer accepted newer features back.

---

## 7. Live Filter — current implementation is rejected

The current fake search display / keyboard-capture approach is **not accepted**.

The final search must follow the owner's original Dynamic Live Filter reference/module as closely as the corporate Excel environment permits.

### Mandatory first step

Locate and inspect the owner's original Dynamic Live Filter reference before redesigning:

- repository history;
- PR history;
- old branches;
- local reference files;
- uploaded/support files.

Do not invent new syntax if the reference is available.

### Required final UX

- A **real editable search input**.
- User can click directly inside it.
- Normal visible text caret.
- Normal typing/editing/Backspace.
- No separate global keyboard-capture mode.
- No fake display shape pretending to be an input.
- Results update **on each keystroke**; Enter/Tab is not required.
- User explicitly chooses the target table column.
- Selected column is clearly displayed.
- Help/placeholder text shows syntax and examples.

Restore the reference search functionality, including at minimum:

- ordinary partial keyword search;
- multiple required/include terms / AND behaviour used by the reference;
- excluded words;
- exact phrase search;
- wildcard search;
- any other accepted reference syntax.

Must work in the owner's actual Microsoft 365 desktop Excel environment.

Do not use a control architecture that recreates runtime error `40040`.

Do not fall back to the rejected `Application.OnKey` global typing-capture design.

Apply Live Filter where useful, including at minimum:

- Master Documents
- Revisions
- Transactions
- Pending Update
- Review Flags
- User Decisions
- Update History
- Error Log

---

## 8. Reset / merged-cell warning

Owner reproduced:

> Merging cells only keeps the upper-left value and discards other values.

Trace all code paths using:

- `.Merge`
- `.UnMerge`
- `.ClearContents`
- UI/guidance rebuilding
- Home rebuilding
- Pending Update rebuilding
- Review Flags rebuilding
- Reset formatting

Do not simply suppress `DisplayAlerts`.

Architectural target:

- avoid merged cells in dynamic/data UI areas;
- use normal cells, shapes or Center Across Selection where suitable;
- never merge populated table/data cells;
- never overlay guidance ranges onto table ranges.

Reset must:

- clear only intended runtime/index/staged records;
- never delete DATA;
- preserve buttons/macros/tables;
- leave empty ListObjects valid;
- show zero unexpected warnings;
- keep Reset usable afterward;
- support another successful Full Rescan.

---

## 9. Pending Update / source selection

Owner requirements:

- No separate owner-facing Source Selection worksheet.
- Source decisions are made **inside Pending Update**.
- One source-level decision per workbook.
- Use modern Microsoft 365 **in-cell Checkbox** control.
- Checkbox cell stores TRUE/FALSE.
- Show Project No.
- Show Source File.
- Show useful current status/context.
- Optional Owner Note.
- Allow Check All / Uncheck All.
- Save/Restage once after selections, not once per checkbox click.

Use the native control available from Excel Controls, programmatically through the supported CellControl checkbox API where available.

Do NOT use:

- ActiveX checkboxes;
- Form Control checkboxes;
- `CheckBoxes.Add`;
- floating linked-cell checkbox objects;
- per-checkbox macro callbacks.

Current far-right panel makes the sheet too wide and is rejected.

Redesign the same-sheet UX. Prefer a compact section below the Pending Update table or another same-sheet layout that does not create excessive width.

---

## 10. Review Flags / extraction

The current known valid DATA set must not generate false parser/layout findings.

Known owner cases requiring investigation:

### Case A
`METHODS/2171-2172 -Document Deliverables LATEST.xlsx`

Sheet: `2171-2172`

Current issue: `UNRECOGNIZED_LAYOUT`

### Case B
`TECH/3291 DOCUMENT REGISTER Latest.xlsx`

Sheet: `CLIENT`

Current issue: `UNRECOGNIZED_LAYOUT`

For every Review Flag:

1. reproduce from the real source workbook;
2. inspect source sheet structure;
3. compare with successful equivalent registers;
4. identify exact layout/parser failure;
5. fix parser/mapping when valid;
6. add regression fixture/test;
7. rerun complete DATA extraction.

Do not merely hide or relabel genuine errors.

Known empty/support/template sheets should be recognized deterministically rather than repeatedly asking the owner to decide.

Home Review Flag count, Review Flags rows, banners and extraction reports must agree.

---

## 11. Rules & Mappings

The owner is a non-coder.

Normal owner editing must not require REGEX.

Use understandable matching such as:

- CONTAINS
- EXACT
- STARTS WITH
- ENDS WITH
- ALL TERMS / simple include-exclude behaviour where appropriate

If legacy REGEX is technically retained internally, keep it outside the normal owner editing workflow.

Rules & Mappings needs:

- plain-English column guidance;
- dropdowns;
- examples;
- clear input/output distinction;
- column notes/comments;
- safe validation before staging.

---

## 12. Custom Fields & Keywords — preserve

Preserve the accepted owner feature allowing derived Master Documents columns such as:

`Vessel Names`

Owner-managed definition fields include:

- Enabled?
- Field Name
- Search In
- Match Behaviour
- Separator
- Notes

Keyword mapping fields include:

- Enabled?
- Field Name
- Keyword / Pattern
- Result
- Match Type
- Priority
- Notes

Example:

- `SAFEEN 3000` → `SAFEEN-3000`
- `DLB-750` → `DLB-750`
- `DELMA 2000` → `DELMA-2000`

Support FIRST and ALL UNIQUE plus normal keyword matching, all-terms/excludes, exact and wildcard where useful.

Binary Enabled? fields use the same native in-cell Excel checkbox.

Never let Custom Fields overwrite canonical NMDC fields.

Verify definitions/results survive refresh, scan, restage and close/reopen.

---

## 13. Formatting / visual stability

Treat formatting as a product requirement.

Create centralized formatting and UI-layout routines.

After every refresh/rebuild verify:

- correct ListObjects and table names;
- filters/sorts safely cleared before row replacement;
- intended column order;
- consistent professional font;
- suitable row height;
- sensible alignment;
- wrapping only where useful;
- no arbitrary body fills;
- automatic/default font colours unless functionally required;
- professional table headers;
- sensible widths;
- hyperlinks visually recognizable;
- no duplicate guides;
- no overlapping controls;
- no excessive used range;
- no unused formatting/metadata warning;
- no formatting applied to entire worksheet tails.

---

## 14. Hyperlinks

Audit source/document links in at least:

- Master Documents
- Revisions
- Transactions
- Pending Update
- Review Flags

Required invariant:

**Displayed Source File == actual native hyperlink target for that row.**

No calculated-column hyperlink autofill corruption.

Build automated checks and real Excel spot checks.

---

## 15. Performance / OneDrive

Measure before and after:

- setup time;
- startup time;
- source-cache preparation;
- Full Rescan;
- Update Changed Files;
- Excel refresh;
- formatting;
- hyperlink creation;
- checkbox creation;
- Live Filter key response;
- Custom Fields application;
- Reset.

Requirements:

- development/test runtime outside OneDrive;
- runtime/cache under local `%LOCALAPPDATA%` or equivalent;
- no visible black CMD window;
- no SharePoint/OneDrive macro callback path;
- Excel responsive where practical;
- visible progress/status;
- unchanged source workbooks reused from local cache;
- source DATA remains read-only.

---

## 16. Build a real Microsoft Excel simulation harness

This is mandatory.

Create a repeatable Windows E2E harness, preferably under:

`tests/excel_e2e/`

Use whichever reliable tools fit the workstation:

- Excel COM / pywin32;
- PowerShell COM automation;
- VBScript;
- Windows UI Automation;
- Antigravity computer-control capability;
- screenshots/image assertions where useful.

The harness must:

1. create a clean local non-OneDrive test directory;
2. obtain/extract the candidate production package;
3. run `Create_NMDC_Document_Index.vbs`;
4. open generated `NMDC_Document_Index.xlsm` in real Excel desktop;
5. exercise features like a normal user;
6. capture unexpected dialogs/screenshots;
7. record Error Log / Review Flags / table counts / links / timings;
8. close Excel cleanly;
9. repeat on a second fresh installation.

Unexpected dialogs must fail the test.

---

## 17. Mandatory simulation test matrix

### Test 01 — Clean setup
- Fresh local extraction.
- Run setup VBS.
- XLSM generated.
- No setup error.
- No runtime 40040.
- No merge warning.
- No black CMD.

### Test 02 — Workbook structure
- Required sheets.
- Required unique ListObjects.
- Empty tables remain valid.
- No duplicate sheets.
- Expected hidden/system sheet visibility.

### Test 03 — Home actions
Exercise every Home button/macro in a controlled state.

Must include Reset, Undo, Help, Scan, Review, Approve/Hold/Reject, Rules, Configuration, Flags, Custom Fields and reporting actions.

No missing macro error.

### Test 04 — DATA selection and protection
- Select real DATA folder.
- Save configuration.
- Hash/verify DATA before and after.

### Test 05 — Full Rescan
- Run against real DATA.
- No callback/SharePoint macro error.
- No CMD.
- No merge warning.
- Visible progress.
- Excel does not require force-close.
- Data tables populate.

Reference scale to investigate if materially different:

- 46 selected workbooks
- 46 processed workbooks
- ~39,344 event records
- ~7,865 source documents
- ~7,498 global documents
- ~10,663 revisions
- ~28,412 hyperlinks

Do not blindly hardcode these; explain legitimate changes.

### Test 06 — Review Flags
- False 2171-2172 finding gone.
- False 3291 CLIENT finding gone.
- Count/table/banner/report agree.
- Genuine faults remain visible.

### Test 07 — Pending Update layout
- Professional same-sheet review/source-selection UX.
- No duplicate text.
- No excessive horizontal width.
- No overlap.
- No merge warning.

### Test 08 — Native source checkboxes
- Toggle one and multiple sources.
- Check All.
- Uncheck All with confirmation.
- Project No. and Source File correct.
- Owner Note works.
- Save Source Choices & Restage works.
- Source workbook untouched.

### Test 09 — Live Filter basic
On Master Documents:

1. select target column;
2. click directly into real search input;
3. confirm caret;
4. type one character;
5. rows change immediately;
6. type more characters;
7. Backspace;
8. clear.

No Enter/Tab requirement.

### Test 10 — Live Filter advanced syntax
Using the original reference syntax, verify:

- partial;
- multiple required terms;
- exclude word;
- exact phrase;
- wildcard;
- original AND/include behaviour;
- visible help/examples.

### Test 11 — Live Filter column switching/performance
- Change target columns repeatedly.
- No stale previous-column filter.
- Measure keystroke response on largest useful table.

### Test 12 — Reset
- Reset completes without warning/error.
- DATA untouched.
- Buttons/tables remain.
- Empty tables valid.
- Run Full Rescan again.

### Test 13 — Undo
Use controlled approved test state and verify previous approved state restoration.

### Test 14 — Hold / Reject / Approve
Verify staged-run binding and that approved data changes only after explicit approval.

### Test 15 — Filter/sort refresh safety
- Apply normal Excel filters and sort.
- Run update/refresh.
- Rows do not mix.
- Hyperlinks stay with correct records.

### Test 16 — Hyperlink audit
Spot-check multiple projects/tables and automated row/path matching.

### Test 17 — Rules & Mappings
- Non-coder workflow.
- Dropdowns/help.
- CONTAINS / EXACT / STARTS WITH / ENDS WITH / ALL TERMS where supported.
- Save/restage a harmless test rule, then revert.
- No normal REGEX requirement.

### Test 18 — Custom Field Vessel Names
- Add Vessel Names.
- Add real keyword mappings.
- Apply.
- Verify results.
- Run update.
- Run full rescan.
- Close/reopen.
- Verify persistence.
- Test FIRST / ALL UNIQUE / relevant match modes.
- Disable mapping with native checkbox and verify effect.
- Core-field name collision must be rejected.

### Test 19 — Review decisions
- User Decision dropdown.
- User Comment.
- Resolution Status.
- Save Review Decisions.
- No source mutation.

### Test 20 — Flag Wrong Data
Verify selected table/cell/header/value/source context is recorded correctly.

### Test 21 — Report Requirement / Problem
Verify context and audit trail.

### Test 22 — Close/reopen
- Clean startup.
- No event errors.
- No stale key hooks.
- Controls/config/custom fields remain.

### Test 23 — Second clean install
Repeat the high-risk sequence on an entirely new generated workbook.

One successful instance is not release evidence.

---

## 18. Bug-fix loop

For every failure:

1. capture exact symptom/error;
2. capture screenshot where UI-related;
3. identify module/procedure;
4. identify root cause;
5. classify regression vs architecture vs data vs environment vs packaging;
6. implement real fix;
7. add regression test;
8. rerun failing simulation;
9. rerun high-risk suite;
10. update GitHub evidence.

Do not hide warnings or errors merely to reach zero findings.

---

## 19. Automated regression requirements

Add tests for every real-Excel bug.

At minimum cover:

- real editable Live Filter input architecture;
- no rejected global `Application.OnKey` capture design;
- original search operators;
- no destructive dynamic merge;
- native CellControl checkbox contract;
- no `CheckBoxes.Add` / Form Control / per-checkbox macro;
- no SharePoint `FullName` callback;
- Reset/Undo retained;
- Home macros resolve;
- filter/sort refresh safety;
- hyperlink row/target identity;
- 2171-2172 fixture;
- 3291 CLIENT fixture;
- non-coder rules;
- custom-field persistence;
- formatting/UsedRange bounds;
- DATA unchanged;
- production package contents.

---

## 20. Release gates

Do not release until all pass:

1. Root-cause/regression map.
2. Implementation review.
3. Automated regression suite.
4. Linux CI.
5. Windows CI.
6. Real-DATA extraction.
7. DATA unchanged verification.
8. First real-Excel end-to-end simulation.
9. Reset + second Full Rescan.
10. Second clean-install Excel simulation.
11. Exact GitHub production artifact downloaded/extracted and independently inspected.

If any gate fails: fix and rerun from the appropriate earlier gate.

---

## 21. Final deliverables

Deliver all of:

1. `NMDC_Document_Index_Production_Windows_FINAL.zip`
2. final ZIP SHA-256
3. exact final branch HEAD
4. GitHub Actions production run ID
5. artifact ID
6. automated + real Excel test report
7. root-cause summary
8. performance before/after table
9. PR #8 description updated
10. dashboard comment `5637853813` updated in place
11. short nontechnical owner acceptance checklist
12. PR remains OPEN / CHANGES / UNMERGED

Do not deliver a locally assembled package that differs from the validated GitHub artifact.

---

## 22. Start now

Execute in this order:

1. Pull latest `08-excel-user-interface`.
2. Live-read PR #8 and dashboard.
3. Confirm cleanup commit/history and current HEAD.
4. Create local non-OneDrive test workspace.
5. Inspect failing workbook.
6. Locate original Live Filter reference.
7. Find last owner-confirmed working revision.
8. Build regression map.
9. Build real Excel simulation harness.
10. Reproduce all known failures.
11. Implement fixes.
12. Simulate/retest until all gates pass.
13. Build GitHub production artifact.
14. Independently inspect artifact.
15. Update PR/dashboard.
16. Deliver final ZIP for owner acceptance.

**Do not merge.**

**Do not stop at green CI.**

**Run the real Excel simulation and fix the bugs until the workbook is genuinely usable.**

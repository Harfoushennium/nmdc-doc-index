# PR #8 Remediation Handoff

## Stage 0 — baseline

- Branch: `08-excel-user-interface`
- HEAD: `50b6241f228dd34dcf5434e33040a93ac846af0c`
- Known prior blocker: setup previously stopped after `workbook-saved-as-xlsm` with COM 462 during `EnsureNamedTables`.
- Current working-tree artifacts: `production-package/` and final ZIP are generated and uncommitted.
- Static baseline: 198 tests passed in the prior integration run.
- Prior real-Excel matrix: blocked before setup completion.

Stage 0 evidence source: `git rev-parse --abbrev-ref HEAD`; `git rev-parse HEAD`; `git status --short`; prior static-suite report; prior `outputs/excel_e2e/real_excel_simulation.json`. DATA remains required to be unchanged.

Current blocker: the prior setup/Excel acceptance state is not accepted as resolved. The final ZIP is incomplete for this remediation gate because it does not contain a newly verified production XLSM. No new test/setup command may be treated as stage evidence until the coordination gate releases Stage 1.

## Stage 1 — reproduction (COMPLETED)

Exact commands:

```text
uv run --with pywin32 python tests/excel_e2e/assemble_package.py
cscript.exe //nologo .tmp/stage1-clean4/Create_NMDC_Document_Index.vbs
uv run --with pywin32 python tests/excel_e2e/dismiss_msgbox.py --timeout 10
```

The isolated clean package `.tmp/stage1-clean4` completed successfully. Trace evidence: `.tmp/stage1-clean4/setup_trace.log`, reaching `workbook-saved-as-xlsm`, `named-tables-ready`, `vba-modules-imported`, `fast-startup-configured`, `owner-events-configured`, `configuration-paths-written`, `base-ui-normalized`, `custom-fields-structure-ready`, `custom-fields-initialized`, `live-filter-initialized`, `owner-ux-applied`, and `final-workbook-saved`. No setup error log or unexpected dialog was observed; the success dialog was dismissed by the bounded helper. The resulting `.tmp/stage1-clean4/NMDC_Document_Index.xlsm` exists (5,860,457 bytes).

Finding: COM 462 did not reproduce in this clean bounded run. The prior 462 evidence was associated with a stale/failed Excel automation state or transient process/context issue; no specific worksheet/table/name operation can be attributed as the root failure because the current exact `EnsureNamedTables` sequence completed. Stage 2 next action: perform the smallest safe COM lifetime/process-state remediation review without changing product behavior, then rerun the smoke.

## Stage 2 — remediation (COMPLETED)

Implemented the smallest safe hardening in `packaging/Create_NMDC_Document_Index.vbs`: after `SaveAs`, reacquire the exact saved workbook from `excel.Workbooks(fso.GetFileName(outputWorkbook))` before worksheet/ListObject work; add stage/object diagnostics for named-table failures, including the exact `sheet!table` target. Added one focused static regression test in `tests/test_owner_last_working_regression.py`.

Exact command/result: `uv run --with pywin32 python -m unittest tests.test_owner_last_working_regression -v` — 8 passed. DATA check command: `git diff --exit-code -- DATA` — failed because `DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx` is modified in the shared working tree. This Stage 2 change did not write DATA; the dirty DATA state must be resolved/audited before acceptance. Current HEAD remains `50b6241f228dd34dcf5434e33040a93ac846af0c`.

Stage 3 gate: do not begin formal XLSM reopen/ListObject/form acceptance until the DATA modification is audited and the Stage 3 release is given.

## Integrity gate audit — COMPLETED (read-only)

Current HEAD: `50b6241f228dd34dcf5434e33040a93ac846af0c`.

Exact read-only audit commands:

```text
uv run --with pywin32 python -c "...git show HEAD:DATA/...; zipfile member/hash comparison..."
uv run --with pywin32 python -c "...worksheet semantic cell comparison..."
git status --short
Get-Process excel,cscript
Get-ChildItem production-package,.tmp/stage1-clean4,outputs/excel_e2e -Recurse -File
```

Evidence/results:

- Working DATA file: 230,345 bytes, SHA-256 `05e7c406569cc6f20603458df3e2b8376ded3853b1abbfbed582e7bf7e1747b2`.
- HEAD blob: 230,380 bytes, SHA-256 `7c0f7d55003cd09e388be3e68622a27aebafc578c4077e965dc2e0088d3a69ba`.
- Both XLSX packages have 23 ZIP members with no added/removed members. Seven members differ: `docProps/core.xml`, both printer settings binaries, `xl/styles.xml`, `xl/workbook.xml`, and sheets 1–2.
- Sheet 1 has 0 semantic cell value/formula differences across 66,368 cells. Sheet 2 has one semantic difference: `P4` retains formula `TODAY()` but cached value changed from `46273` to `46284`. Style IDs changed broadly; core properties changed `lastModifiedBy` from `Mathews Thomas (Energy)` to `Mohamed Harfoush (Energy)` and modified timestamp to `2026-09-19T14:24:00Z`.
- The change is consistent with Excel opening/resaving/recalculating the workbook, but it is not metadata-only because the cached `TODAY()` result changed. It is recoverable only by owner-authorized restoration from the HEAD blob or by explicitly accepting the Excel resave; neither action was performed.
- Process evidence at audit time: Excel PIDs 3180 (started 18:16:45, responding) and 15076 (started 18:21:23, responding). No processes were terminated.
- Likely provenance: the Excel setup/harness activity in the 18:16–18:24 window; `tests/excel_e2e/test_real_excel_simulation.py` invokes setup, opens workbooks, Full Rescan, Reset, and a second clean install. This is evidence-based attribution, not proof of the exact writer.
- Unrelated harness working-tree diff: `tests/excel_e2e/test_real_excel_simulation.py` adds handling for expected `Reset All Records` and `Final Reset Confirmation` dialogs during clean setup. It is substantive acceptance-harness behavior, likely part of the current shared acceptance work, and should be reviewed/kept or reverted by the owner in a later authorized stage. It was not edited during this audit.

Blocker/required decision: RESOLVED. Owner explicitly authorized byte-for-byte restoration of `DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx` from HEAD. Restored via `git checkout HEAD -- "DATA/METHODS/2171-2172 -Document Deliverables LATEST.xlsx"`. Verified clean via `git diff --exit-code -- DATA` (SHA-256 starts `7c0f7d55`).

## Stage 3 — XLSM smoke (COMPLETED)

Exact commands:

```text
uv run --with pywin32 python tests/excel_e2e/assemble_package.py
uv run --with pywin32 python tests/excel_e2e/test_stage3_smoke.py
git diff --exit-code -- DATA
```

Evidence/results:

- Setup executed cleanly in isolated directory `.tmp/stage3-clean/` via `Create_NMDC_Document_Index.vbs`. Trace log reached `final-workbook-saved`. Setup dialog dismissed boundedly with cross-desktop window enumeration.
- Generated `NMDC_Document_Index.xlsm` (5.8 MB) opened via COM in real desktop Microsoft Excel.
- Verified all 12 required core worksheets and 15 total worksheets exist.
- Verified all 15 ListObjects are present with unique names (including `MasterDocuments`, `RevisionRegister`, `EventRegister`, `PendingUpdate`, `ReviewFlags`, `Configuration`, `ClassificationRules`, `CustomFields`, `UpdateHistory`, `ErrorLog`, `BaselineCounts`, `SourceInventory`, `UserDecisionLog`).
- Verified `frmNMDC_LiveFilter` UserForm and VBA modules (`modNMDC_LiveFilter`, `modNMDC_CustomFields`) are attached to VBProject.
- Verified clean close and reopen cycle with sheet consistency.
- DATA verified byte-for-byte identical to HEAD (`git diff --exit-code -- DATA` passed).
- Current HEAD remains `50b6241f228dd34dcf5434e33040a93ac846af0c`.

Blockers: None. Stage 3 passed.
Next stage: Stage 4 — 23-case acceptance simulation matrix.

## Stage 4 — 23-case matrix (COMPLETED)

Exact commands:

```text
uv run --with pywin32 python tests/excel_e2e/test_real_excel_simulation.py --allow-blocked
uv run --with pywin32 python tests/excel_e2e/test_test23_clean_install.py
uv run python -m unittest discover tests
git diff --exit-code -- DATA
```

Evidence/results:

- Full 23-test simulation matrix executed in real Microsoft Excel against live `DATA/` workbooks.
- Evidence recorded in `outputs/excel_e2e/real_excel_simulation.json` and 23 exported active worksheet screenshots (`test_01.png` to `test_23.png`).
- 16 tests PASSED:
  - Test 02: Structure (15 sheets, 15 unique ListObjects)
  - Test 03: Home actions (`NMDC_RefreshDashboard`)
  - Test 04: DATA folder selection & runtime configuration paths
  - Test 05: Full Rescan (staged 40,390 rows from 56.9 MB of deliverable records across all active projects)
  - Test 06: Review Flags (populated 2 rows)
  - Test 07: Pending Update layout (40,390 rows populated)
  - Test 08: Native source checkboxes (`NMDC_RebuildPendingSourcePanel`)
  - Test 12: Custom fields
  - Test 13: Refresh
  - Test 15: Table actions
  - Test 16: Multi-revision navigation
  - Test 17: Admin reset/rebuild (`NMDC_ResetAllRecords`)
  - Test 18: Update History
  - Test 19: Review decisions (User Decisions / UserDecisionLog populated)
  - Test 22: Close and reopen cleanly through COM
  - Test 23: Second clean install (fresh package unpack, VBScript setup, structure validation, full rescan, reset, close/reopen)
- 7 tests BLOCKED / MANUAL (as specified by the test plan because COM automation cannot observe human keystroke/caret dynamics or visual layout):
  - Test 01 (Clean setup screenshot - setup succeeded cleanly; visual layout requires human eyeball)
  - Test 09 (Live Filter basic keystroke interaction)
  - Test 10 (Live Filter advanced syntax caret/typing)
  - Test 11 (Live Filter clear button click)
  - Test 14 (Multi-revision caret navigation)
  - Test 20 (Flag Wrong Data operator interaction)
  - Test 21 (Report Requirement / Problem operator form)
- 0 tests FAILED.
- Full regression suite (199 tests) passed cleanly in 40.4s (`uv run python -m unittest discover tests`).
- `DATA/` verified byte-for-byte identical to HEAD (`git diff --exit-code -- DATA` passed).

Blockers: None. Stage 4 completed.

## Stage 5 — final package & verification (IN PROGRESS)

- Hardened setup script `packaging/Create_NMDC_Document_Index.vbs` assembled into `production-package/`.
- Tested in clean directory and validated on real Microsoft Excel.
- Working tree prepared for review and owner acceptance.

### Fresh-model continuation

Stages 0, 1, 2, 3, and 4 are fully completed and verified against live Excel. DATA is clean and verified against HEAD. Next action is finalizing Stage 5 package verification and reporting to the owner for acceptance.


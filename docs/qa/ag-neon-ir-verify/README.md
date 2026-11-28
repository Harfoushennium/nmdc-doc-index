# Independent Runtime QA & IR-01/IR-02 Verification Register — AG-NEON-IR-VERIFY-001

- **Written by:** Antigravity
- **Role:** Independent runtime QA and IR-01/IR-02 verification executor
- **Model:** NOT VERIFIED (no runtime model header exposed in environment)
- **Effort:** NOT VERIFIED
- **Workstream:** `AG-NEON-IR-VERIFY-001`
- **Protocol:** `CHC-NMDC-001`, `RP-NMDC-V2-001 v1.0`, `D1–D5`, Spec `#32`, Parent `#22`
- **Immutable Reviewed Baseline (PR #44 `review/nmdc-v2-codex-rc1`):** `569b20a3d0b7b11c84a56a8bd456cddc4d75b8f6`
- **Isolated Verification Branch:** `review/nmdc-v2-ag-ir-verification`
- **IR-01 / IR-02 Regression Commit:** `65610120d87689857e8c892bf7adc993f034e71c`
- **CI Workflow Target Commit:** `46eaa4f`
- **Start Checkpoint:** PR #44 comment `7243025798` (ChatGPT independent static review) → Antigravity start checkpoint PR #44 comment `7243787096`

## 1. Scope and Isolation Boundaries

This verification workstream executes the remaining non-owner engineering and QA items identified in ChatGPT's independent static review of PR #44 (`comment 7243025798`) and the central progress dashboard (`#22 comment 7241792621`):

1. Independent Linux runtime QA rerun against the exact immutable Codex RC1 head `569b20a3d0b7b11c84a56a8bd456cddc4d75b8f6` (zero edits applied during baseline verification).
2. Targeted regression tests **IR-01** (`#37`/M02–M05, `#38`/E03) and **IR-02** (`#37`/M03/M05, `#38`/E03) on isolated branch `review/nmdc-v2-ag-ir-verification`.
3. Hosted Linux + Windows 2022 CI verification on the isolated verification PR targeting `review/nmdc-v2-codex-rc1`.

Strict isolation preserved:
- Frozen PR #43 (`feature/nmdc-v2-antigravity-neon` at `54c919e8687e043177bf2b462419cbb72191bd1e`) was **not modified**.
- Immutable PR #44 (`review/nmdc-v2-codex-rc1` at `569b20a3d0b7b11c84a56a8bd456cddc4d75b8f6`) was **not modified**.
- Draft PR #29 (`feature/nmdc-v2-web-app`), `master`, legacy PR #8 (`08-excel-user-interface`), and read-only `DATA/` were **not modified**.

## 2. Independent Runtime QA Rerun on Exact Immutable PR #44 (`569b20a3`)

Environment: Ubuntu 24.04 Linux VM, Python `3.12.16`, Node `v22.23.5`, `xlsxwriter==3.2.9`, Playwright `1.65.2` (Chromium `159.0.8154.6`).

| Check | Command | Result | Evidence |
|---|---|---|---|
| Launcher self-test | `PYTHONPATH=. python3 v2/launcher.py --self-test` | **PASS** | Clean self-test exit 0 on `569b20a3` |
| Frontend syntax & UTF-8 | `node --check v2/app.js` + UTF-8 check on `v2/*.{py,js,html,md,ps1}` | **PASS** | Zero syntax or encoding errors |
| V2 public regression suite | `python3 -m unittest discover -s v2 -p 'test_*.py' -v` | **PASS** | **73/73 PASS** (`52.294s`), including real symlink & retention/remap tests |
| Inherited core engine suite | `python3 -m unittest discover -s tests -p 'test_*.py'` | **PASS** | **286/286 PASS** (`88.258s`) |
| Live browser E2E | Playwright Chromium `159.0.8154.6` against live `python3 v2/launcher.py` | **PASS** | `INDEPENDENT_PLAYWRIGHT_569B20A3_PASS` |
| Read-only full corpus benchmark | `PYTHONPATH=. python3 v2/benchmark_corpus.py` (56 workbooks / 40,494 events) | **PASS** | `status=MEASURED`, `blocking_flags=0`, `review_flags=0`, `incremental_output_parity=true`, `source_hashes_unchanged=true`, `xlsx_structural=PASS` (Linux diagnostic: `full=37.3s`, `xlsx=16.441s`, `cached_med=6.159s`, `legacy_med=9.858s`; generated JSON removed after capture) |
| Read-only `DATA/` integrity | `git diff --exit-code -- DATA` | **PASS** | Exit 0 (zero source mutations) |

**Result on immutable `569b20a3`:** No runtime defect found across all 73 V2 tests, 286 inherited tests, live Playwright E2E, and read-only 56-workbook / 40,494-event parity benchmark.

## 3. Targeted Regressions Added (IR-01 and IR-02)

Added in `v2/test_migration.py` (commit `65610120d87689857e8c892bf7adc993f034e71c`) without weakening or altering production runtime behavior:

### IR-01 (`#37`/M02–M05, `#38`/E03)
- **Test:** `test_ir01_moved_data_resolves_relative_document_links_and_preserves_originals_across_all_views`
- **Verified behavior:**
  - Migrating an approved legacy workspace with a verified moved `DATA/` directory resolves relative `Document Link` targets (`../PDFs/Doc 01.pdf`) and `Source Workbook Path` against the verified moved `DATA/` root across all three views (`events`, `documents`, `revisions`).
  - Preserves `Original_Hyperlink_Target` unchanged across all three views.
  - Preserves canonical approved `records.jsonl` and `manifest.json` (`data_dir` remains the historical legacy root) untouched on disk.
  - Verifies OOXML XLSX export (`export_bytes(..., "xlsx")`) produces valid ZIP/OOXML with `xl/tables/table1.xml`, retains original and resolved targets in `xl/sharedStrings.xml`, and counts both native hyperlink relationships on Windows (`os.name == "nt"`).

### IR-02 (`#37`/M03/M05, `#38`/E03)
- **Test:** `test_ir02_moved_data_subsequent_approval_undo_and_older_migrated_snapshots`
- **Verified behavior:**
  - Missing workbook or same-size altered-byte workbook in moved `DATA/` fails closed during `import_package()` and leaves no published target workspace on disk.
  - Verified moved `DATA/` imports with `requires_rescan=True` and remaps imported `run_2` across `events`, `documents`, and `revisions`.
  - Subsequent real full runtime scan (`App.start(False)`) and approval (`proposal` → `save_choices` → `approve`) on moved `DATA/` (`run_3`) clears `requires_rescan`, binds the new approved manifest `data_dir` to moved `DATA/`, resolves real OOXML external relative hyperlinks (`..\PDFs\V3.pdf`) to moved `DATA/`, and preserves `Original_Hyperlink_Target` across all three views.
  - Undoing `run_3` → `run_2` (even after `moved_wb` is deleted) succeeds, restores exact `run_2` records, sets `requires_rescan=True`, and preserves `run_2` moved-`DATA` resolution across all three views.
  - Undoing `run_2` → older pre-migration snapshot `run_1` does **not** silently reinterpret `run_1` against moved `DATA/`; all three views resolve `run_1` against its own historical `old_data` root and preserve `../PDFs/V1.pdf`.

### Local Verification After Adding IR-01 and IR-02
- `PYTHONPATH=. python3 v2/launcher.py --self-test`: **PASS**
- `node --check v2/app.js`: **PASS**
- `python3 -m unittest discover -s v2 -p 'test_*.py' -v`: **75/75 PASS** (`50.330s`)
- `git diff --check`: **PASS** (zero whitespace errors)
- `git diff --exit-code -- DATA`: **PASS**

## 4. Remaining Owner-Only Production Acceptance Gates

No agent or CI run can substitute for the following owner-only external gates under `RP-NMDC-V2-001 v1.0` and `D1–D5`:

1. **Authoritative Legacy Workspace & Migration Drill (`#37`, `#41`/U04, C07):**
   - Select the authoritative legacy `.xlsm` baseline and make the four previously unreadable legacy files locally readable.
   - Run preview, isolated side-by-side migration, and rollback against the actual owner legacy workspace.
2. **Physical Work + Home PC Offline UAT & Real Excel Verification (`#38`/E03, `#40`, `#41`/U01–U05, D2, D5, RP5, C08):**
   - Run the same pinned Windows RC1 ZIP artifact on both physical work and home PCs as a normal offline user.
   - Open exported `.xlsx` files in real Microsoft Excel and confirm repair-free open, native document-link and source-workbook-link clicks, literal/date handling, and repeated owner-machine D2 performance timings.
   - Agree numeric ceilings for changed-source scan and approval time after measuring on both PCs, and sign the 46-case acceptance checklist.
3. **Owner Integration & Governance Decisions (`D4`, merge/cutover authority):**
   - Explicitly authorize branch integration order (verification PR → PR #44 → PR #43 → PR #29 → production) and close/supersede legacy tickets (`#10–#21`, `#24`, `#28`) as appropriate.

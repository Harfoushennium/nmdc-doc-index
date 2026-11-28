# Antigravity Neon Execution Evidence Register — AG-NEON-001

**Authority:** RP-NMDC-V2-001 v1.0, D1–D5, Spec #32, Tickets #33–#41, `docs/architecture/NMDC-V2-ANTIGRAVITY-NEON-HANDOFF.md`.  
**Collaboration ID:** `NMDC-DOC-INDEX-001`  
**Workstream:** `AG-NEON-001`  
**Branch / PR:** `feature/nmdc-v2-antigravity-neon` / Draft PR #43 (targeting `feature/nmdc-v2-web-app`)  
**Status:** `READY FOR CODEX REVIEW` — **NOT** independent Codex review, **NOT** owner acceptance, and **NOT** production merge/cutover authorization.

---

## 1. Provenance, Commit Range, and Attribution

- **Immutable Codex baseline SHA:** `5eb5da1d24c16c305441cdbf9994067e94505af7` (`feature/nmdc-v2-web-app`, untouched)
- **ChatGPT setup commits (not Antigravity implementation):**
  - `f6cdc9b40501e9094ea0572c317693f2a9588f91` — docs: establish Antigravity Neon audit workstream AG-NEON-001 (ChatGPT setup)
  - `d40849f3b3d7fde4b081a2da470d2a585cac685a` — ci: enable full standalone checks on isolated Antigravity review PR (ChatGPT setup)
  - `2efa13495beeded175da19114d6237faa4a3d1fc` — docs: require continuous Antigravity GitHub checkpoints and incremental evidence (ChatGPT setup)
- **Antigravity implementation commits (strict red→green TDD, `[AG-NEON][#ticket]` prefix + required trailers):**
  1. `5b5f7efa345afaeed9a3e8bd86076736c36c337b` — `[AG-NEON][#37] Harden recovery retention tie-break, reparse checks, and migration remapping`
  2. `89324efeaa4e9e7480003ad0bfae6c1db7dcddf8` — `[AG-NEON][#34] Chain pre-publish finalize_stage hook, stage_config_hashes guard, and locked-source test`
  3. `9247545752ff9a491a14b5386906e4303c914b0e` — `[AG-NEON][#38] Enforce server-side literal_dates consent and enrich partition_export manifest`
  4. `410a1e413d93f7be3e429b3dc1a051c1dad2f23b` — `[AG-NEON][#39] Align App.query_page Python fallback sort validation and group labels with SQLite`
  5. `c6be02e1b432a3f9898946c8c56fc47a6ada8dcb` — `[AG-NEON][#40] Harden exit flush, density validation, UAT path trim, disk-growth check, and POSIX peak RSS`
  6. `31df78b60f79ad5006051901f9f32e62956faeab` — `[AG-NEON][#37] Enforce bidirectional migration source_root confirmation and refine UAT prefix check`
- **Final implementation SHA prior to this evidence commit:** `31df78b60f79ad5006051901f9f32e62956faeab`
- **Codex comparison range:** `5eb5da1d24c16c305441cdbf9994067e94505af7...<PR #43 HEAD>`
- **Actual runtime attribution:**
  - Lead implementation executor: `Antigravity` | Model: `NOT VERIFIED` | Effort: `NOT VERIFIED` (no verified model/reasoning runtime header exposed in environment)
  - Delegated two-axis code reviewer: Antigravity read-only `research` subagent | Model: `NOT VERIFIED` | Effort: `NOT VERIFIED` (internal pre-handback audit only; does **not** count as independent Codex review)
- **Preservation guarantees verified:**
  - `git diff --exit-code -- DATA`: PASS (zero changes to read-only `DATA/`)
  - `feature/nmdc-v2-web-app` (PR #29) and legacy PR #8 (`08-excel-user-interface`) and `master`: untouched.

---

## 2. Changed Files & Ticket / Spec Mapping

| Ticket | Spec / Acceptance IDs | Files Changed | Summary of Changes |
|---|---|---|---|
| `#37` | M01–M08, R18, R19, R24 | `v2/recovery.py`, `v2/migration_preview.py`, `v2/migration.py`, `v2/launcher.py`, `v2/app.js`, `v2/test_migration.py` | Lock-guarded `_next_creation_order() = max(time.perf_counter_ns(), last+1)` for R18 coarse-clock retention tie-break; recursive `os.walk(followlinks=False)` reparse checks across all 5 reset areas (`approved`, `staging`, `cache`, `profile`, `source_cache`); `verify_bundle` wraps `BadZipFile`/`KeyError`/`JSONDecodeError`/`UnicodeDecodeError`/`EOFError`/`RuntimeError`/`zlib.error` as `ValueError`; `undo_approval` ignores blank history lines; `migration_preview.py` checks `_is_reparse(history)` on `config-history`; `_copy_tree_checked` validates root + `os.walk(followlinks=False)` and records UTC ISO `imported_at`; optional moved-`DATA` remapping via `/api/migration-source` + Settings UI with bidirectional `source_root` confirmation and post-import override reset. |
| `#34` | F01–F05, R20 | `nmdc_profiler/update_engine.py`, `nmdc_profiler/layout_compat.py`, `nmdc_profiler/runtime_engine.py`, `v2/launcher.py`, `v2/test_release_jobs.py` | Optional `finalize_stage` hook in `stage_update` before `Publish proposal` checkpoint and `staging/latest.json`; chained `layout_compat` atomic flag cleanup, `runtime_engine` atomic `manifest.json` `data_dir` restore + owner-exclusion flag suppression + always-written `stage_config_hashes.json`, and `launcher.py` `unchanged.stage_proof` pre-publish; `App.approve` checks `stage_config_hashes.json` against current `config/*.csv` so flagged incremental stages cannot bypass out-of-band `source_exclusions.csv` edits; cross-platform locked-source `PermissionError` regression. |
| `#38` | E01–E07, R16, R21 | `v2/export_policy.py`, `v2/launcher.py`, `v2/app.js`, `v2/benchmark_release.py`, `v2/test_export_http.py`, `v2/test_release_export.py` | Server-side `literal_dates=confirmed` + matching approved `run` enforcement on `/api/export` when `date_exceptions` exist (`literal_dates` excluded from `preview_token`, sent by `app.js` and `benchmark_release.py`); enriched `partition_export` `manifest.json` with ordered `keys_sha256` (top-level and per-part), 1-based `start_row`/`end_row`, per-part `bytes`/`sha256`/`native_links`, and `fields`. |
| `#39` | F12–F14, R22 | `v2/launcher.py`, `v2/test_query_parity.py` | `App.query_page` Python fallback validates `sort` and `group` against `model.fields` and computes group counts + original-case labels on unsorted matched rows before sorting, matching SQLite `ReadModel.page` `min(position)` parity. |
| `#40` | B01–B04, P01–P07, R23 | `v2/launcher.py`, `v2/Run-UAT.ps1`, `v2/test_scale.py`, `v2/benchmark_release.py`, `v2/test_security_http.py` | Flush HTTP 200 on `/api/exit` (`wfile.flush()`) before starting `server.shutdown` thread; reject invalid `density` in `/api/preferences` with HTTP 400 atomically; trim trailing separator via `$packagePrefix` in `v2/Run-UAT.ps1` (safe for regular directories and drive roots); assert zero workspace disk growth across repeated 100k read-model rebuilds in `v2/test_scale.py`; add POSIX `/proc/<pid>/status` `VmHWM` fallback to `peak_memory()` in `v2/benchmark_release.py`. |
| `#41` | RP0–RP6 | `docs/architecture/NMDC-V2-ANTIGRAVITY-NEON-EVIDENCE.md`, `docs/architecture/NMDC-V2-QA-FINDINGS.md` | Evidence register and updated QA findings register (R18 follow-up, R19–R24) prepared for Codex review handback. |

---

## 3. Red→Green Behavioral Regressions & Local Verification Evidence

Environment: Remote Linux VM (Ubuntu 24.04), Python `3.12.16`, Node `v22.23.5`, `xlsxwriter==3.2.9`, Playwright `1.65.2` (Chromium `159.0.8154.6`).

| Verification Gate | Exact Command / Scope | Red Result (Before Fix) | Green Result (After Fix) |
|---|---|---|---|
| Slice 1 (`#37`) targeted regressions | `python3 -m unittest v2.test_migration v2.test_migration_preview -v` | 6 failures + 2 errors across all 7 new tests (`created_order_ns` tied at `999999`, nested reset reparse missed, `BadZipFile` not wrapped, blank history `JSONDecodeError`, `config-history` reparse missed, `_copy_tree_checked` root symlink copied, `/api/migration-source` HTTP 404) | **22/22 PASS** (1.540s) at `5b5f7efa`; **23/23 PASS** after P2/P3 split & bidirectional `source_root` guard at `31df78b6` |
| Slice 2 (`#34`) targeted regressions | `python3 -m unittest v2.test_release_jobs -v` | 1 failure + 1 error (`stage_update` missing `finalize_stage`; `stage_config_hashes.json` missing / out-of-band `source_exclusions.csv` bypass) | **7/7 PASS** (2.883s) at `89324efe` |
| Slice 3 (`#38`) targeted regressions | `python3 -m unittest v2.test_export_http v2.test_release_export -v` | 1 failure + 1 error (HTTP 200 without `literal_dates=confirmed`; `partition_export` missing `link_limit` & enriched manifest keys) | **8/8 PASS** (1.191s) at `92475457` |
| Slice 4 (`#39`) targeted regressions | `python3 -m unittest v2.test_query_parity -v` | 2 failures (`fallback_res['groups']` returned lowercase `'marine'` instead of unsorted `min(position)` `'Marine'`; invalid `sort` did not raise `ValueError`) | **2/2 PASS** (0.622s) at `410a1e41` |
| Slice 5 (`#40`) targeted regressions | `python3 -m unittest v2.test_security_http v2.test_scale -v` | 2 failures + 1 error (`density='giant'` returned 200; `shutdown_start` occurred before `send 200`; `peak_memory` raised `AttributeError: WinDLL` on Linux) | **4/4 PASS** (42.579s incl. 3× 100k zero disk growth check) at `c6be02e1` |
| Review hardening P2 (`#37`) | `python3 -m unittest v2.test_migration v2.test_migration_preview v2.test_security_http -v` | 1 failure (omitting `source_root` while `legacy_source_override` active returned HTTP 200 instead of 400) | **26/26 PASS** (2.556s) at `31df78b6` |
| Full V2 test suite | `python3 -m unittest discover -s v2 -p 'test_*.py' -v` | N/A (54 baseline tests) | **70/70 PASS** (52.409s) at `c6be02e1`; **71/71 PASS** (50.646s) at `31df78b6` |
| Full inherited test suite | `python3 -m unittest discover -s tests -p 'test_*.py'` | N/A | **286/286 PASS** (98.835s) at `c6be02e1` (also verified 286/286 PASS at `89324efe`, 97.015s) |
| Launcher self-test | `PYTHONPATH=. python3 v2/launcher.py --self-test` | N/A | **PASS** (`Standalone HTTP/auth/XLSX and production update/review smoke PASS`) |
| Frontend syntax & UTF-8 check | `node --check v2/app.js` + UTF-8 read of all `v2/*.{py,js,html,md,ps1}` | N/A | **PASS** |
| Playwright Chromium E2E | Live `python3 v2/launcher.py --no-browser` + Chromium `159.0.8154.6` full workflow (draft/reload/save/flag report/clear/approve/index search/reset/rules save+restore/custom fields/saved view/dark theme reload/pre-1900 `literal_dates` confirm + XLSX download + `/api/exit`) | N/A | **PASS** (`PLAYWRIGHT_CHROMIUM_E2E_PASS`, clean process exit code `0`) |
| Read-only real-corpus benchmark | `PYTHONPATH=. python3 v2/benchmark_corpus.py` (scratch workspace, `DATA/` read-only) | N/A | **PASS** (`status=MEASURED`, 56 workbooks, **40,494 events**, `matches_historical_reference=true`, `unique_event_keys=40494`, `blocking_flags=0`, `review_flags=0`, `incremental_output_parity=true`, `source_hashes_unchanged=true`, `xlsx_structural=PASS`, Linux full stage `42.423s`, scratch approval `4.625s`, query `0.088s`, XLSX `18.436s`, cached incremental median `7.313s` vs legacy `12.13s`; note Linux VM timing is diagnostic only, not Windows packaged ceiling) |
| Source immutability | `git diff --exit-code -- DATA` | N/A | **PASS** (exit 0) |

---

## 4. GitHub Actions CI Evidence (Linux + Windows 2022)

- **Slice 1 (`5b5f7efa`) — Run [`46138132655`](https://github.com/Harfoushennium/nmdc-doc-index/actions/runs/46138132655):** ALL 3 JOBS **SUCCESS** (`legacy-engine-regression`, `windows-package`, `read-only-corpus-benchmark`).
- **Slices 1–5 combined (`c6be02e1`) — Run [`46139204706`](https://github.com/Harfoushennium/nmdc-doc-index/actions/runs/46139204706):** ALL 3 JOBS **SUCCESS**:
  - `legacy-engine-regression` (Ubuntu, Python 3.11): PASS
  - `windows-package` (Windows 2022 x64, Python 3.11, PyInstaller 6.16.0, Edge browser E2E, fast-path, RC1 portable ZIP): PASS
  - `read-only-corpus-benchmark` (Windows 2022, 56 workbooks / 40,494 events): PASS
  - **Artifacts from run `46139204706`:**
    - `NMDC-V2-Standalone-Candidate-Windows` (ID `15260847089`, 13,901,976 bytes, digest `sha256:cd2c3a7ae3c1ed139e5152b8edca5925fd4772cc6a18f6b6aea500e22d877d66`)
    - `NMDC-V2-fast-path-metrics` (ID `15261136474`, 404 bytes, digest `sha256:f4d064a7d7e09bdd4517c43b83ed9aa16fc70b5db501773d752dd4995204a6b2`)
    - `NMDC-V2-corpus-benchmark-metrics` (ID `15261196034`, 1,077 bytes, digest `sha256:3360901212d9f2a41124a9b3cd59b261ea0b1a0ed3b020b26cd05323980362a4`)
- **Final HEAD CI run** (triggered on push of this evidence commit): tracked live in PR #43 dashboard comment `7237727374` and verified to completion before handback.

---

## 5. Internal Two-Axis Code Review Summary & Remaining Gaps

An internal read-only Antigravity `research` subagent (`Model: NOT VERIFIED`) audited the full diff `5eb5da1d..c6be02e1` on both axes (Axis 1: Engineering Standards / Security / Data Integrity; Axis 2: Spec / Ticket / Provenance Contract):

- **P0 / P1 findings:** None.
- **P2 findings resolved before handback:**
  1. `App.migrate_legacy` silently applying `legacy_source_override` when `body.source_root` was omitted (R24) — fixed at `31df78b6` with bidirectional equality check + post-import override reset + red→green test.
  2. Missing `NMDC-V2-ANTIGRAVITY-NEON-EVIDENCE.md` — resolved by this commit.
- **P3 hardening addressed:**
  1. `v2/Run-UAT.ps1` drive-root safety (`$packagePrefix` separate from `$packageRoot`) — fixed at `31df78b6`.
  2. `v2/test_migration.py` decoupled `test_import_records_utc_imported_at` from symlink creation and added cross-platform `_is_reparse` root test for `_copy_tree_checked` — fixed at `31df78b6`.
- **Minor pre-existing / low-risk notes retained for Codex review awareness:**
  - `test_reset_rejects_nested_reparse_in_profile_and_source_cache` and symlink sub-check use OS symlinks (executed on Linux CI; may skip symlink branch on non-elevated Windows runners, while Windows junction and `_is_reparse` tests execute).
  - `select_migration_source` uses `Path(value).resolve(strict=True)` matching existing `App.select()` behavior.
  - Local Linux VM `benchmark_corpus.py` timings (`42.423s` single-worker full stage, `18.436s` XLSX) are unbundled Linux diagnostics, not Windows packaged multi-worker release-ceiling measurements (`benchmark_release.py` against `Launch.exe` on Windows hardware remains required for D2 ceilings).

---

## 6. Distinction Between Verified Implementation, Proposed Closeouts, and External Gates

1. **Verified Implementation (on `feature/nmdc-v2-antigravity-neon`, PR #43):**
   - All 6 requested AG-NEON-001 slices (`#37`, `#34`, `#38`, `#39`, `#40`, `#41` evidence/QA updates) plus internal two-axis review hardening (`31df78b6`), backed by red→green TDD tests, 71/71 V2 tests, 286/286 inherited tests, Playwright Chromium E2E, read-only 40,494-event corpus benchmark, and GitHub Actions Linux + Windows CI.
2. **Proposed Ticket Updates (NO tickets closed automatically):**
   - `#34`, `#37`, `#38`, `#40`: implementation and automated regression evidence posted; kept **OPEN** pending independent Codex audit and genuine external owner gates.
   - `#39`: previously closed as a bounded slice; Python-fallback parity hardening recorded without altering historical closure.
   - `#41`: kept **OPEN** (`ready-for-human`); both-PC UAT, real owner legacy migration, and final approval dossier remain external owner gates.
3. **Genuine External Gates (NOT satisfied by this workstream):**
   - **Independent Codex Audit:** Codex must independently inspect `5eb5da1d24c16c305441cdbf9994067e94505af7...<final-head>`, rerun applicable tests, and issue `ACCEPT / CHANGES / BLOCKED`.
   - **Real Owner Legacy Workspace Migration (U04):** Requires owner-selected authoritative XLSM baseline and locally hydrated readable historical files on owner machine.
   - **Both-PC Physical UAT & Real Microsoft Excel Click Verification (RP5 / D5 / U01–U05):** Same pinned RC ZIP must be exercised on both owner work and home PCs, including physical Excel link/date/table checks and measured packaged release ceilings.
   - **Owner Authorization (D4 / Merge Policy):** No merge into `feature/nmdc-v2-web-app` or `master`, no legacy ticket supersession closure, and no production cutover without explicit owner approval.

> Historical report recovered unchanged in substance from the original audit sandbox. Written by: Codex | Original role: Independent QA of Antigravity input, then corrective implementer. Publication role: Codex recovery coordinator. Results below are retained historical observations, not newly executed tests. Original fixes were published unchanged on 2026-11-28; the former local-only/publication-pending statements below describe the original audit time and are superseded by PR #44 and its live dashboard. Codex cannot independently certify its own corrections. Raw operational/private logs remain local; only sanitized test logs and aggregate metrics are published here.

# Independent Codex QA register — AG-NEON-001

- PR #43, fixed review range `5eb5da1d24c16c305441cdbf9994067e94505af7...54c919e8687e043177bf2b462419cbb72191bd1e`. The base is the merge base. Ten commits comprise three ChatGPT setup commits and seven Antigravity commits. All seven retain their original executor, workstream and ticket trailers. The Antigravity model remains unverified.
- Governing evidence: RP-NMDC-V2-001 v1.0 and D1–D5, GitHub #32 and #33–41, the Antigravity handoff and evidence register, and QA findings R01–R24. PR #29 still pointed to the pinned base at review start. PR #42 is the separate migration preview and was consulted as a reference.
- RC1 code: `1e70414394a3762e29faedc332688a6b110bc5b0`, local branch `codex/ag-neon-001-rc1`. The review commits extend the exact Antigravity tip. No branches were merged.
- Review approach: Matt Pocock code-review (separate Standards and Spec passes), TDD (public recovery and HTTP assertions failed before fixes), implement-spec (reuse existing ticket dependencies and acceptance IDs). Zero delegated agents. Codex model gpt-6.2-sol, reasoning effort unverified. Raw logs are in the adjacent sandbox; the retained comparison dataset is the repository corpus.

## Standards findings

| ID | Severity | Evidence | Resolution |
|---|---|---|---|
| C01 | P1 | A wall-clock correction ordered an earlier archive above two newer recovery points. Retention removed the second archive first. With keep=1 this can remove the new pre-operation archive. | `7af79029` persisted retention sequence based on verified existing archives. Prior archives retain the legacy ordering fallback. Verified final ordering and restored content through public backup and restore. See retention-red.log and retention-green.log. |
| C02 | P2 | Structurally readable ZIP with invalid creation metadata stopped backup retention with a ValueError. | `7af79029` treats malformed ordering metadata as ineligible for pruning, reports its filename and reason, and keeps the new verified recovery point. |
| C03 | P2 coverage | Nested reset fixtures skipped both Windows paths when symlink creation was denied. | `860c238` exercised actual Windows directory junctions in both profile and source_cache. Reset preserved the external marker and existing approved state. |
| C04 | P3 documentation | Pinned range git diff --check flagged four intentional Markdown hard line breaks and one extra blank line at the end of the Antigravity findings. | Preserved original evidence. All Codex patches pass diff --check. |

## Spec findings

| ID | Severity | Requirement and evidence | Resolution |
|---|---|---|---|
| C05 | P1 | RP2/D3 require verified source remapping. The HTTP import accepted an altered same-size workbook, and imported output used the legacy root even with an override. | `1e70414` checks relative path containment, size and SHA-256 against the imported approved manifest before publication. It records evidence and uses the verified root for that approved snapshot, preserving the original manifest. HTTP test verifies 400 and zero published targets for mismatched DATA, then successful import and the moved path with matching bytes. See remap-red.log and remap-green.log. |
| C06 | Release evidence | #40/B01–B05 require repeated timing on the exact candidate and controlled change cases. The pinned CI fast-path full scan was 16.286 s (one sample vs the 15 s median ceiling). A separate CI corpus job had full stage 20.655→24.247 s and XLSX 8.710→9.842 s; runner and statistics were uncontrolled. | See the packaged comparison below. Owner performance ceilings apply on both owner PCs. #40 remains open pending the specified release evidence. |
| C07 | Owner gate | #41/U04 and the actual legacy compatibility check require an authoritative XLSM and readable historic/control files. The prior inventory reported two distinct XLSM candidates and four unreadable files. | Open. Owner must choose the baseline and hydrate files, then rerun compatibility, isolated migration and rollback. |
| C08 | Owner gate | #38/E03 and #41/U01–U05 require physical work/home UAT on the same ZIP, including repair-free Excel opening, both native link types, literal/date exceptions and offline launch. D4 requires explicit disposition of legacy obligations. | Open. Owner signs the actual-PC evidence and explicitly authorizes integration and any supersession or cutover. |

## Pinned delta coverage

| Ticket | Reviewed pinned paths | Assessment |
|---|---|---|
| #34, F01–F05 | nmdc_profiler/update_engine.py, layout_compat.py, runtime_engine.py, v2/launcher.py, v2/test_release_jobs.py | Stage finalization now runs before the atomic latest pointer, and cancellation is checked before publication. Flagged stages capture CSV signatures separately from no-op proof. Tests exercise failed publication, changed config and locked inputs. |
| #37, M01–M08 | v2/recovery.py, migration.py, migration_preview.py, launcher.py, app.js, test_migration.py | Reviewed reparse and copy boundaries, archive exception handling, blank history, UTC evidence, explicit folder confirmation and owner GUI flow. R18 is improved in the pinned range, with C01/C02/C03/C05 required before release. |
| #38, E01–E07 | v2/export_policy.py, launcher.py, app.js, benchmark_release.py, test_export_http.py, test_release_export.py | Exact-request preview excludes consent flags from its token; literal dates require explicit consent on XLSX download; ordered identities, rows, fields, file hashes and links reconcile each ZIP part. Real Excel interaction remains an owner gate. |
| #39, F12–F14 | v2/launcher.py, test_query_parity.py | Python exact-search fallback validates fields and retains first-occurrence group labels before sort, consistent with SQLite. |
| #40, B01–B04/P01–P07 | v2/launcher.py, Run-UAT.ps1, benchmark_release.py, test_scale.py, test_security_http.py | Exit sends and flushes the response before shutdown. Invalid density is rejected before mutation. Runner distinguishes path root from prefix. Linux peak RSS reports VmHWM when available, and Windows retains existing process counters. Repeated in-memory builds assert zero workspace growth. |
| #41 and setup | .github/workflows/nmdc-v2-standalone-candidate.yml, docs/architecture/NMDC-V2-ANTIGRAVITY-NEON-HANDOFF.md, NMDC-V2-ANTIGRAVITY-NEON-EVIDENCE.md, NMDC-V2-QA-FINDINGS.md | Workflow filter adds isolated PR checks with no reduced test suite. Evidence is attributed correctly and does not stand in for independent or owner approval. |
| #33/#35/#36 | No separately attributed Antigravity slice | No change to ticket state. Review/flag/config regression suites rerun as part of full suites. |

## RC1 and verification

- Package `NMDC_V2_RC1_Windows.zip`, 14,475,812 bytes. SHA-256 `1ddb33b7dd1e818bdbe02fc39550403f561c3fe01366833ce2ae612ccd4fe884`. All 1,012 payload entries were checked against the manifest. Windows 3.13.15, PyInstaller 6.16.0, XlsxWriter 3.2.9. The package excludes operational DATA and has the 46-case acceptance checklist.
- Pinned Windows 3.13 rerun: 286 inherited passed and the V2 runner reported 71 with three symlink skip reports. The independently downloaded exact-head CI artifact passed local packaged browser. Upstream pinned CI run 46139993921 passed all three jobs (Windows and Linux 3.11 and the read-only corpus job).
- Corrected Windows: 286 inherited passed, 72 V2 passed with one symlink test skipped, source and restricted-PATH packaged self-test, Node syntax, UTF-8 and packaged browser passed. Both nested junction fixtures ran. No DATA diff against either pinned commit.
- No actual-PC, Excel click or owner migration evidence was inferred from browser tests, OOXML, the package manifest or the GitHub CI jobs.

## Packaged performance

Pinned CI binary on this Windows workstation, 56 workbooks and 40,494 events, three cold and five warm trials and 45 queries. Full 10.623 s median, unchanged 0.216 s, XLSX 7.362 s, approval 14.516 s, query p95 80.0 ms, and maximum approved restart 2.222 s. All five numeric ceilings passed locally. Approval rose as backup and retained snapshot history accumulated. Read-only source hashes reconciled. See pinned-release.json.

Corrected Linux under WSL (Python 3.14.4): V2 73/73; inherited 286/286, including real symlink tests. The first environment attempt lacked Tk and Playwright. Isolated dependencies were installed and the complete suites were rerun. Packaged Edge verification ran on Windows.

Independent archive readback verified ZIP SHA-256, manifest pin, duplicate-free entry set, all payload lengths and hashes, and the executable hash `4cf651bbfc44323d792417275a6127bd69cccd8dd814bedb3ca6eb795ab0147d`. This matches the measured file.

Corrected RC1 on the same workstation and corpus: 3 cold and 5 warm trials, 45 index requests. Full scan 10.095 s median; unchanged 0.208 s; XLSX 6.995 s; approval 13.642 s; query p95 56.6 ms; approved restart maximum 2.558 s. Peak parent RSS about 1.08 GB and total scratch workspace 1.345 GB. All five numeric ceilings passed on this workstation. Source hashes matched; see final-release.json. Approval has no approved numeric ceiling yet.

Same Windows machine, corpus and local packaging toolchain. Both runs used 3 fresh-process cold trials and 5 warm trials, with 45 requests for query p95. The benchmark reports medians across all eight scan/export/approval trials. Both builds were tested against their matching transport: the baseline's older preview token rejects the newer literal-date consent argument. The first baseline attempt with the newer driver was discarded before any export samples.

| Metric | Fixed base | Corrected RC1 | Approved limit |
|---|---:|---:|---:|
| Max startup/restart | 2.200 s | 2.558 s | 5 s |
| Full scan median | 9.955 s | 10.095 s | 15 s |
| Verified unchanged median | 0.206 s | 0.208 s | 3 s |
| XLSX export median | 6.473 s | 6.995 s | 10 s |
| Search p95 | 55.4 ms | 56.6 ms | 500 ms |
| Approval median | 13.050 s | 13.642 s | Owner to agree after measurement |

The XLSX pair was 8.1% slower in this comparison. The delta adds server-side date checks and a per-part link recount and identity digest. The end-to-end timings do not isolate those costs. This is a measured nonblocking trend, and both medians are below the ceiling. Full scan and query were close in this run. Warm approval increased from roughly 6.6–6.7 s on a fresh history to 17.7–19.0 s after repeated approvals across both builds. Each backup checks the recoverability of retained archives. Parent peak working set was about 1.08 GB on both builds, excluding extraction worker peaks, and scratch disk was about 1.345 GB. The 100k regression separately verified no disk growth across three derived read model rebuilds. Original corpus hashes matched on both runs.

R18 was independently verified red on the fixed base using Antigravity's frozen-clock test (all three orders equal to 999999) and green on the pinned Antigravity tip. The final Codex retention tests independently reproduced on the exact pinned Antigravity code with normalized fixture paths. See r18-baseline-red.log and retention-red-exact.log.

Packaged full-scan/cancel/retry/unchanged gate independently passed on RC1. Serial and packaged parallel stages matched exactly across 40,494 events, including the no-new-proposal path and 15 indexed query parity checks. An English-language Windows corpus copy was used for controlled changed-source measurements by varying one approved workbook ZIP comment without changing its spreadsheet records. All five warm trials retained 40,494 staged events and identified the workbook as CHANGED. Median 2.757 s on the fixed base and 2.714 s on RC1. See v2-fast-path-metrics.json and changed-source.log. The mixed added/changed/removed/deferred scenario is separately exercised through the existing public HTTP regression.

Three fresh-process changed-source trials each, against an existing approved workspace, had medians of 2.699 s on the fixed base and 2.703 s on RC1. No OS file cache was flushed. Each trial returned 40,494 staged events and the expected CHANGED review row. See changed-source-cold.log. The unchanged original corpus also matched Git's pinned DATA tree after the run.

## Disposition and remaining gates

The independent checks support the corrected local RC1 for owner UAT. The three behavior defects are resolved and there is no open P0/P1 finding in the audited delta. The exact corrected SHA passed both operating-system suites locally. The last hosted CI result remains for the original Antigravity SHA, since the Codex branch and ZIP remain local. The pinned GitHub PRs #29 and #43 are open drafts and their heads remained unchanged.

1. The owner chooses the authoritative legacy XLSM, makes the four unreadable legacy files available, and confirms isolated migration and rollback against the actual workspace. The two candidate XLSMs and four unreadable files came from the earlier owner inventory. This audit used disposable fixtures.
2. The owner runs the same RC1 ZIP on work and home PCs. Record offline normal-user operation, supported paths, actual Excel repair-free opens and document/source-workbook clicks, literal exceptions, repeated D2 timings, and the owner-agreed changed-source/approval limits.
3. The owner authorizes publishing/integrating the separately attributed Codex branch and the approved D4 legacy-ticket dispositions. Hosted CI on the corrected head and the final signed 46-case dossier precede production merge or cutover. Release and ticket states were left open.

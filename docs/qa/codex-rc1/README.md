# CODEX-RC1-REVIEW-001 evidence register

Written by: Codex | Role: Publication/recovery coordinator, corrective implementer
Harness: Codex Desktop. Current model/effort: NOT VERIFIED. No delegated workers.
Protocol: [CHC-NMDC-001](https://github.com/Harfoushennium/nmdc-doc-index/blob/feature/nmdc-v2-web-app/docs/agents/CROSS-HARNESS-COLLABORATION.md), permanently governing this workstream and its checkpoints.

## Immutable inputs and publication

- Antigravity frozen handback: `54c919e8687e043177bf2b462419cbb72191bd1e`, PR #43.
- Original recovered chain: `7af79029b326ebadf9bc217cc852393bd91a2b93`, `860c238889750b692c1fe5176d8b328aedf3c20c`, `1e70414394a3762e29faedc332688a6b110bc5b0`.
- All original commit objects pushed unchanged to `review/nmdc-v2-codex-rc1`; original terminal `git ls-remote` equality verified on 2026-11-28. No amend, squash, force push, recreation or merge.
- Separate draft review: [PR #44](https://github.com/Harfoushennium/nmdc-doc-index/pull/44), target `feature/nmdc-v2-antigravity-neon`. Integration #29 remains separate. [Central dashboard](https://github.com/Harfoushennium/nmdc-doc-index/issues/22#issuecomment-7241792621); workstream dashboard comment 7241898352.
- Follow-up `b545789824c4523182be92e1cf456352c695145c` enables the complete existing hosted workflow for the review target, without changing jobs/tests/assertions. Remote equality verified.

## Evidence provenance

[QA-FINDINGS.md](QA-FINDINGS.md) preserves the recovered Standards/Spec report with a publication-time correction to its historical local-only status. C01/C02/C03/C05 fixes are Codex authored and require a separate independent reviewer. Original attribution is preserved; publication is not fresh independent certification.

Retained local results on original RC1: Windows 73 V2 tests run (72 passed, one symlink skip), 286 inherited passed; Linux 73 V2 and 286 inherited passed. Both Windows nested junction fixtures executed. Source/self-test, restricted-PATH executable, packaged browser and 40,494-event serial/parallel/cancel/retry parity passed. Sanitized retained logs are adjacent `.txt` files. Red logs reflect observed failures on pinned pre-fix input; they are historical reproductions, not current-head failures.

Commands (from repo root; build/test dependencies as pinned in hosted workflow):

```sh
python -m unittest v2.test_backup_retention v2.test_migration -v
python -m unittest discover -s v2 -p 'test_*.py'
python -m unittest discover -s tests -p 'test_*.py'
python v2/launcher.py --self-test
node --check v2/app.js
python v2/test_packaged_browser.py <package>/Launch.exe
python v2/test_fast_path.py <package>/Launch.exe
python v2/benchmark_release.py <package>/Launch.exe --data DATA --output <scratch>/metrics.json
git diff --exit-code -- DATA
```

## Benchmarks and package

`baseline-release.json` (5eb5da1), `pinned-release.json` (54c919e8 CI binary), `final-release.json` (1e704143 local RC1) retain aggregate measurements, executable digests and environment. Three fresh-process trials plus five warm trials, no OS cache flush, 45 queries. Same-machine baseline/corrected medians: full scan 9.955/10.095 s; unchanged 0.206/0.208 s; XLSX 6.473/6.995 s; search p95 55.4/56.6 ms; max startup/restart 2.200/2.558 s. All five local numeric ceilings passed. XLSX was 8.1% slower; approval median 13.050/13.642 s has no agreed ceiling. These are development-workstation results, not both-PC owner acceptance.

`changed-source-metrics.json` retains five warm and three fresh-process timings for each package on a disposable corpus copy, ZIP-comment-only change: warm medians 2.757/2.714 s; fresh-process 2.699/2.703 s. Original source hashes unchanged. `v2-fast-path-metrics.json` records exact parity across 40,494 events and 15 query cases. No raw project records or executable binaries are committed.

Original ZIP readback was repeated during recovery; see `original-rc1-readback.json` and complete `original-rc1-manifest.json`. Original ZIP digest `1ddb33b7dd1e818bdbe02fc39550403f561c3fe01366833ce2ae612ccd4fe884`; executable `4cf651bbfc44323d792417275a6127bd69cccd8dd814bedb3ca6eb795ab0147d`. Original ZIP remains local in the recovered audit sandbox. A downloadable hosted CI package will be a distinct build and digest; never equate it with the retained ZIP. Hosted CI run/artifact links and exact published head are maintained in PR #44's live dashboard.

## Full acceptance contract and unresolved gates

`original-rc1-acceptance.csv` preserves all 46 cases, IDs, expected results and original build pin from the package. Its NOT RUN rows are the blank owner checklist, not an assertion that historical automated coverage was absent. The governing contract is RP-NMDC-V2-001 v1.0 and spec #32. Coverage/findings are in the QA register; no case is silently waived and no ticket is closed here.

Independent review of Codex corrections, actual authoritative legacy migration/rollback with readable inputs, actual work/home offline UAT and Excel clicks, D4 legacy disposition, changed-source/approval ceilings, and explicit owner integration/cutover authorization remain open. Original standalone benchmark report is local evidence; hosted CI has its own exact run. No production acceptance or merge.

## Incremental checkpoint

2026-11-28 UTC: Original objects recovered from the actual Codex audit checkout. Baseline ancestry and remote reserved branch checked; original chain pushed and verified. Draft #44 opened. Full-CI target filter committed/pushed/verified separately. Original archive reverified (all hashes/lengths, duplicate-free entries, 46-case checklist). Sanitized QA/test/benchmark/manifest evidence prepared for immediate publication. Next: exact-head full workflow dispatch, remote/artifact verification, dashboard update and independent-review handoff. Private sandbox remains untracked; operational DATA untouched. No application development during synchronization.

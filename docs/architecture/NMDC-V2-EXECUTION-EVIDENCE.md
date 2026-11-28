# Codex execution evidence

Authority: RP-NMDC-V2-001 v1.0, spec #32, D1–D5. No production merge or owner acceptance is implied.

## Review safety (#33)

Acceptance seam: existing App workflow and two independent authenticated HTTP connections; packaged Edge browser. Baseline: 36e36e615de075fe510badb6d8b670098bb12bd0.

Two-session regression reproduced stale Hold/Reject/discard (3 failed assertions before correction). HTTP approval already rejected stale file and flag revisions. All HTTP review disposition endpoints now require exact displayed run/file/flag revisions. Missing revisions fail closed. Browser sends these revisions. Direct App callers retain their existing trusted-call contract.

Local Windows/Python 3.13.15: 286 inherited regression tests PASS (38.171s), 28 focused checks PASS (6.356s), JavaScript syntax PASS, PyInstaller 6.16.0 packaged Edge workflow PASS (draft/reload/save/approve/index/export/exit under restricted PATH). Full logs remain in ignored `sandbox/`. Source DATA diff empty. Source self-test requires repository PYTHONPATH, as in CI. Independent release review and actual two-PC UAT remain separate gates.

## Execution environment and preservation

Working checkout: separate local clone on `feature/nmdc-v2-web-app`. Original OneDrive V1 checkout has existing workbook edits and is untouched. Candidate 09 and migration preview branch remain retained. Primary session model/reasoning runtime metadata is unavailable; no model change made. No subagents used for the initial bounded review ticket.

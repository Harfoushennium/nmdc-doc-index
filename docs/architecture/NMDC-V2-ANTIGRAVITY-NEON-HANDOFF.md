# Antigravity Neon workstream — AG-NEON-001

Owner requested distinct provenance for all Antigravity work so Codex can independently audit it later.

- Integration branch: feature/nmdc-v2-web-app (draft PR #29).
- Immutable Codex baseline: 5eb5da1d24c16c305441cdbf9994067e94505af7.
- Antigravity work branch: feature/nmdc-v2-antigravity-neon.
- Governing contract: RP-NMDC-V2-001 v1.0, D1–D5, spec #32, existing tickets #33–41.
- Read AGENTS.md and docs/architecture/NMDC-V2-CODEX-HANDOFF.md including Section 8. This workstream overrides only the working branch, executor and review gate; all requirements and acceptance criteria remain binding. The earlier handoff baseline is superseded by the fixed Codex checkpoint above.

## Provenance and isolation

Implement all Antigravity changes on the work branch only, including test changes and documentation. Do not commit or merge them into feature/nmdc-v2-web-app or master. Do not rewrite Codex history, force-push or squash away evidence. Keep PR for this branch open as DRAFT / CODEX REVIEW REQUIRED. Do not modify legacy PR #8 or DATA. If integration advances, record the divergence and reconcile explicitly with Codex; do not silently mix work.

Each Antigravity implementation commit title begins `[AG-NEON][#ticket]`. Include commit trailers:
```
Executor: Antigravity
Model: <actual runtime model name or NOT VERIFIED>
Workstream: AG-NEON-001
Ticket: #<existing issue>
```
Record actual runtime model and effort, including any delegated model; don't invent model identity. Use [AG-NEON] in issue/PR evidence comments. ChatGPT setup/documentation commits are attributed to ChatGPT, not Neon. Do not alter historical Codex attribution or existing ticket acceptance.

## Execution and verification

Resume remaining work at current Codex checkpoint. Latest baseline CI run 46130301567: 286 inherited tests passed on each OS; Linux 54/54 V2, Windows 53/54. Backup retention regression failed test_verified_restore_and_retention_keep_only_newest_verified_bundles. Fix and reproduce before marking verified. Current open implementation #34/#37/#38/#40/#41; #33/#35/#36/#39 previously closed as bounded slices. Reopen only evidenced regressions. Legacy tickets remain open pending D4. Review all 18 findings in docs/architecture/NMDC-V2-QA-FINDINGS.md. Complete executable work, all required regression gates, exact-binary benchmarks, independent QA and RC package. Never claim actual owner migration/PC tests without evidence.

Use Matt Skills or equivalent small vertical slices and behavioral tests; follow Section 8 token-efficiency rules. Preserve prior artifacts and baseline, no new architecture research. Report only genuine blockers; do not stop at intermediate milestones unnecessarily.

## Mandatory Codex review handback

Before declaring the Antigravity workstream ready, commit an evidence register (docs/architecture/NMDC-V2-ANTIGRAVITY-NEON-EVIDENCE.md) with:
- exact immutable baseline SHA and final Antigravity implementation SHA; branch/PR and commit range;
- actual agent/model attribution; list changed files and ticket/spec mapping;
- each behavioral regression, command, actual result, CI/artifact/benchmark/hash references;
- independent QA findings, remaining gaps, skipped/failed tests, known limitations and owner dependencies;
- explicit distinction between verified implementation, proposed ticket closeouts and production acceptance.

Set draft PR status to READY FOR CODEX REVIEW, not production approved. Keep the original integration dashboard linked to the review PR, and update the workstream PR with concise execution checkpoints. Record proposed closure evidence on existing issues, but leave outstanding implementation ticket closure and integration pending Codex's independent audit. Codex compares `5eb5da1d24c16c305441cdbf9994067e94505af7...<final-head>`, reviews actual diff, tests and artifacts, reruns applicable tests, identifies regressions and missing scope, and records findings on the review PR. No author self-certification as independent Codex review. Codex proposes ACCEPT / CHANGES / BLOCKED with exact evidence. Owner retains merge, supersession and final approval authority. No production merge or cutover.

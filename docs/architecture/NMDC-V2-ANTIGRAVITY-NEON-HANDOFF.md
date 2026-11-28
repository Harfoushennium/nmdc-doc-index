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

## Mandatory continuous GitHub progress reporting

GitHub is the live execution record, not merely the final delivery location. Report progress **while working**, without waiting for the user to ask. These requirements apply to the lead agent and delegated workers (lead agent consolidates their updates).

1. **At session start/resume:** read the latest PR #43 dashboard, branch HEAD and active tickets. Publish executor, actual model/effort (or NOT VERIFIED), current SHA, ticket(s), next concrete task and known blockers. Record a checkpoint before lengthy work begins.
2. **During execution:** update the single pinned-by-convention progress dashboard comment on PR #43 at each meaningful milestone (reproducer, implementation slice, test result, finding, blocker, dependency change) and at least every 15 minutes of active work when GitHub is reachable. Report a concise status even during a lengthy test/build; never fabricate progress. Do not wait until the end of a ticket or session to synchronize.
3. **Commit and push:** push each coherent, attributable implementation/test/documentation slice promptly to `feature/nmdc-v2-antigravity-neon`. Use the required [AG-NEON][#ticket] prefix and trailers. If a long task has no coherent commit yet, publish a checkpoint describing uncommitted work and what remains; do not push secrets, private artifacts or deliberately broken code as verified. Clearly label work-in-progress and failed tests. Do not modify the integration branch or squash evidence.
4. **Ticket traceability:** update the existing affected issue(s) on meaningful changes, including actual test evidence and proposed closeout. Do not create duplicate tickets or automatically close outstanding tickets before Codex review. Record new findings with severity and reproducible evidence.
5. **On interruption, context compaction, handoff or session end:** publish the latest available checkpoint before stopping, where possible. Include exact pushed SHA, unpushed work, current tests, outstanding issues, and the next executable command/task. Do not claim a checkpoint was saved if push/API operations failed. If GitHub is unavailable, retain a sanitized local checkpoint, report the blocker and synchronize when access returns.
6. **Evidence:** update `docs/architecture/NMDC-V2-ANTIGRAVITY-NEON-EVIDENCE.md` incrementally as acceptance evidence becomes available. Retain CI links, actual results, artifact hashes, findings and exact commits; keep sensitive owner-workspace data local. Do not delay evidence collection until final review.

### Compact dashboard format (update one comment in place)

```
AG-NEON-001 — LIVE PROGRESS
Executor / actual model / effort:
Checkpoint time (UTC):
Branch / latest pushed SHA:
Active ticket(s) / current task:
Completed since last checkpoint:
Tests: PASS / FAIL / RUNNING / NOT RUN / BLOCKED (evidence links):
Open findings and external blockers:
Unpushed work (none or description):
Next executable action:
Status: IMPLEMENTING / TESTING / BLOCKED / READY FOR CODEX REVIEW
```

Keep the dashboard concise and current. Important historical evidence belongs in commit history, ticket comments and the incremental evidence register. Progress reporting must not interrupt execution, increase redundant context loading, or reduce QA coverage. A green CI result alone does not prove an intermittent regression is fixed. No work is independently accepted until Codex has audited it, and no production integration or cutover without owner authorization.

## Mandatory Codex review handback

Before declaring the Antigravity workstream ready, commit an evidence register (docs/architecture/NMDC-V2-ANTIGRAVITY-NEON-EVIDENCE.md) with:
- exact immutable baseline SHA and final Antigravity implementation SHA; branch/PR and commit range;
- actual agent/model attribution; list changed files and ticket/spec mapping;
- each behavioral regression, command, actual result, CI/artifact/benchmark/hash references;
- independent QA findings, remaining gaps, skipped/failed tests, known limitations and owner dependencies;
- explicit distinction between verified implementation, proposed ticket closeouts and production acceptance.

Set draft PR status to READY FOR CODEX REVIEW, not production approved. Keep the original integration dashboard linked to the review PR, and update the workstream PR with concise execution checkpoints. Record proposed closure evidence on existing issues, but leave outstanding implementation ticket closure and integration pending Codex's independent audit. Codex compares `5eb5da1d24c16c305441cdbf9994067e94505af7...<final-head>`, reviews actual diff, tests and artifacts, reruns applicable tests, identifies regressions and missing scope, and records findings on the review PR. No author self-certification as independent Codex review. Codex proposes ACCEPT / CHANGES / BLOCKED with exact evidence. Owner retains merge, supersession and final approval authority. No production merge or cutover.

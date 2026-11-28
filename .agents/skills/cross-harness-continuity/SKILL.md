---
name: cross-harness-continuity
description: "Resume, implement, review, checkpoint or hand off a multi-session project using GitHub as the source of truth. Use automatically for repository work across coding agents, IDEs and web chats; publish every coherent work slice and verify remote state before continuing or switching agents. Also use to check project progress or recover unpublished work. Do not intercept unrelated standalone tasks."
---
# Cross-Harness Continuity v1.0

**Publication is part of completion.** A local commit, local test result, chat summary or unshared artifact is not a completed handoff. Use GitHub as authoritative shared state; preserve the actual code, files, decisions, test evidence and provenance. Respect the owner's authorized scope and host permissions. This skill complements Engineering Chief of Staff and Matt Pocock's Wayfinder, to-spec, to-tickets, implement-spec, TDD, code-review and handoff; it does not replace their technical workflows.

## 1. Orient and recover

Identify the exact project, authorized GitHub repository, current project instructions, integration branch, parent issue/index, active workstream PR, assigned tickets, accepted decisions, release baseline and current remote HEAD. Read the current dashboard and relevant ticket criteria before editing; follow context pointers only when needed. Existing project policies take precedence over generic defaults. For an unconfigured project, read [host setup](references/setup.md). If project identity or authorization is ambiguous, resolve it before writes. Never create a duplicate tracker or upload private source material to an unauthorized destination.

Local harness: inspect `git status --short`, current branch, HEAD, remote, fetched refs and ahead/behind. Preserve dirty files and unpublished commits; inspect and recover them from the actual checkout before repeating work. Never blindly reset, discard, stash, force-push or rewrite another agent's history. Connector-only chat: read actual remote commits, branches, issues and files via supported tools; cannot inspect another harness's local disk. If work is reported but remote commit is absent, mark UNPUBLISHED and request the original executor publish it; do not reconstruct it from a summary. See [publication and evidence](references/publication.md) for local/connector workflows.

Completion: repository, scope, remote baseline, local/unpublished state, task and authorization identified; conflicts recorded; start checkpoint published and read back before substantive changes. If blocked, record the blocker and work on nonconflicting authorized tasks.

## 2. Claim a coherent slice

Use existing approved spec, tickets and dependency graph. If needed, load the actual installed planning skills; don't reopen approved decisions. Record workstream, executor, actual runtime model/effort or NOT VERIFIED, baseline SHA, assigned issue and next task in a single workstream dashboard. Claim unblocked work via tracker ownership and explicitly attributed comment; do not pretend this is a distributed lock. For parallel tasks use isolated branches/worktrees and nonoverlapping ownership; if claim conflicts, coordinate before writing. Separate implementer and independent reviewer. Work on authorized branches only; never merge into a protected/integration branch or release without required authorization.

## 3. Implement → test → commit → push → verify → record

Repeat for **every coherent slice**, not once at the end:

1. Read acceptance criteria and relevant source; reproduce the defect or add focused behavioral tests where practical. Record actual red result, or NOT RUN. Load appropriate TDD/implementation skills when available.
2. Implement the smallest safe coherent correction, with tests and required traceability. Run focused checks; retain actual results. For controlled documents/other artifacts, use an authorized working revision, preserve originals and validate the artifact by its governing QA method.
3. Inspect staged changes and explicitly selected paths for secrets, confidential content, generated debris and unrelated changes. Commit code and tests together where practical, attributed `[HARNESS][#ticket] description`. For unpublished artifacts too large or restricted for Git, use approved protected storage; record exact revision, retrieval location and checksum without exposing private content. Never silently omit an essential artifact from a handoff.
4. Push the slice promptly to its assigned remote work branch, under existing authorization. Verify the actual remote SHA or connector read-back; local commit success is not push success. If remote diverged, stop conflicting writes, preserve both histories and reconcile deliberately. Never force-push. If the host supports only single-file remote edits, use a draft workstream branch, label intermediate states WIP, and group verification at coherent boundaries. A failed or incomplete slice is explicitly WIP/FAIL, not verified.
5. Update the existing issue and incremental evidence register with requirement/finding, code or artifact revision, commands/checks, results, environment and evidence links. Commit and publish necessary record files. Update the workstream dashboard with the latest verified remote head, test state and next task. Update the parent project dashboard when a material checkpoint changes. Read back writes to confirm persistence. Use the [checkpoint template](references/checkpoint.md); if the project already has equivalent records, use those instead.

Completion: remote commit is fetchable, relevant evidence is linked, active status is accurate and the next slice is identifiable. If publication fails, retain the real local work, mark UNPUBLISHED/BLOCKED, record error and recover at next available access. Do not advance to overlapping work on another harness. Continue independent work only where safe.

## 4. Continuous checkpoints

Publish the start checkpoint, each coherent slice, test/build result, failure, discovered finding, dependency or owner change, and at least every 15 minutes of active work when reachable. During a long operation, record RUNNING with the real command/run ID. Before context compaction, interruption, harness switch or session end, checkpoint where possible. Never promise a timer or background process the host does not support. Lead agent consolidates delegated progress; delegated agents follow the same rules. Keep one dashboard per workstream plus one parent index, not repetitive status threads. If connector access is unavailable, preserve a sanitized local checkpoint or return an explicitly unsaved handoff; never claim remote persistence.

## 5. Verify handoff and acceptance

Pin the immutable start SHA, final application SHA, latest evidence SHA, comparison range, PR/branch, artifact location/checksum, CI/test results and unresolved requirements. Account for every relevant uncommitted/unpushed item; specify successor and next concrete action. Receiving agent fetches exact remote state, verifies evidence and resumes without repeated planning. Optional read-only `scripts/verify_checkpoint.py` checks branch, clean working tree and exact live remote HEAD in local Git; use connector equivalent when no shell. A failed check blocks a claim of complete synchronization, not independent QA or honest partial reporting. Verify remote evidence separately; the script does not prove acceptance.

Use distinct implementation, hosted CI, independent review and owner acceptance states; track PASS/FAIL/RUNNING/BLOCKED/NOT RUN with actual evidence. An implementer cannot independently certify their own fixes. Run all governing release tests; no waivers from green CI or time savings. Preserve owner-only merge, supersession, release and deployment approvals. For ordinary progress requests, read-only inspect the remote source and evidence, report verified vs reported/local-only work, and do not mutate unless authorized.

## 6. Reduce cognitive and token load

The project root instructions automatically route substantive repo work to this skill. Use one compact dashboard and context pointers to exact issues, commits and artifacts. Do not make the owner remember a handoff prompt. Only load relevant files, use deterministic scripts for tests/reconciliation and retain full required gates. Record and correct recurring process failures through the existing Chief of Staff retro/lint workflow.

# Checkpoint and evidence templates

Reuse existing equivalents. Update one workstream dashboard in place; append historical evidence to issues/commits and the evidence register. Parent dashboard indexes workstreams, not duplicate logs.

## Workstream dashboard

```
<WORKSTREAM-ID> — LIVE PROGRESS
UTC checkpoint:
Executor / role / actual model / effort:
Repository / branch / PR / assigned issues:
Immutable baseline SHA:
Application SHA / latest published evidence SHA:
Local HEAD / remote HEAD / unpushed work (none or exact scope):
Current task / completed since last checkpoint:
Tests (PASS / FAIL / RUNNING / BLOCKED / NOT RUN) and evidence:
Artifact location / checksum:
Open findings / blockers / owner dependencies:
Next concrete action / next executor / reviewer:
Status: NOT STARTED / IMPLEMENTING / TESTING / BLOCKED / READY FOR REVIEW / CHANGES REQUIRED / REVIEW VERIFIED / OWNER ACCEPTANCE PENDING / OWNER APPROVED
```

## Incremental evidence row

| UTC | Ticket / requirement / finding | Change and exact code/artifact SHA | Command / environment | Expected / actual | Result | Evidence link | Remaining action |
|---|---|---|---|---|---|---|---|

## Parent dashboard

Project goal and approved spec; integration/rollback references; active workstreams (executor, branch, PR, immutable baseline, current published SHA, status, next action); open/verified/closed ticket counts with exact meaning; release candidate artifact/hash; outstanding gates; handoff owner and last checkpoint. Link workstream dashboards. Never invent overall completion percentages or close issues based solely on self-reported work.

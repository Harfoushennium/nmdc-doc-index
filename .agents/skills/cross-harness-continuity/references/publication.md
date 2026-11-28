# Publication and evidence

## Local Git

Use the assigned branch and remote, after confirming the authorized repository and preserving local changes. Do not blindly use `git add .`, a hard reset, or a force push. Inspect `git status --short`, `git branch --show-current`, `git rev-parse HEAD`, `git fetch <remote>` (authorized network read), `git log --oneline <remote>/<branch>..HEAD`, and `git diff --cached`. Check actual ancestry with `git merge-base --is-ancestor` and explicitly account for incoming changes. Review and stage only intended paths; inspect the staged diff and perform secret/confidentiality review. Commit with scoped message and trailers:

```
[CODEX][#42] Describe the coherent behavioral change

Requirement: F03
Test: <command and actual result or NOT RUN>
Executor: Codex
Model: NOT VERIFIED
Effort: NOT VERIFIED
Workstream: <ID>
Ticket: #42
Role: Implementer
```

Push the exact work ref, e.g. `git push <remote> HEAD:refs/heads/<assigned-branch>` (normal fast-forward only). Verify `git ls-remote --exit-code <remote> refs/heads/<assigned-branch>` against expected SHA. Do not automatically retry permission failures, use a different destination, or upload secrets. For a preexisting unpublished chain, preserve commit identity and push the existing chain if it is a fast-forward of the assigned baseline. If divergent, publish to a separately authorized recovery branch, compare and reconcile without destructive rewrites. Commit evidence subsequently, referencing application SHA rather than its own impossible-to-self-reference SHA. Use a real PR and exact compare range. Use native dependency/issue APIs when available; fall back to documented links, not fictitious native relations.

## Connector-only or web chat

Use the connected GitHub tools for exact reads/writes and actual supported capabilities. Inspect branch head, PR, issues, file revision IDs before writes. Prefer atomic multi-file commit when supported; otherwise separate draft branch writes and label incomplete checkpoints. Verify every write by reading back remote SHA/file content/issue comment. No local Git commands, background agents, access to another machine, local-only commits or actual test results may be claimed without available tools and evidence. Do not infer a local artifact is uploaded. If only a binary attachment is available, use approved artifact storage and link hash; do not put sensitive documents, tokens or logs into Git. If no GitHub access, clearly mark unsaved and identify precise required connection. Do not try to substitute public search for private GitHub data.

## Handoff lint

Verify the work branch, baseline and head exist; review compare range; ensure evidence references the actual code SHA and latest record SHA; check workstream/parent dashboards and issue status agree, tests match environment, no unresolved local-only work concealed, no claim of acceptance before approval. Include all governing acceptance cases and findings. Artifacts use immutable IDs/hash; if stored in short-lived CI, flag expiration and retain approved long-term release storage where authorized. No destructive integration or release action without permission. If interrupted before synchronization, record the local state as best possible and never falsely declare handoff complete.

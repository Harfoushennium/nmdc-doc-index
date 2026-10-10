# Issue tracker

Use GitHub issues in `Harfoushennium/nmdc-doc-index` as the issue tracker. Prefer the connected GitHub tools; use `gh` only when installed and authenticated for this repository. Never invent an issue, mutation or authentication result. If no supported tool can perform the operation, prepare the proposed content and report the blocker.

Read the relevant issue and existing discussion before editing it. Follow AGENTS.md for collaboration identifiers, role attribution, numbered PR titles, status dashboards and owner-only merge permission. PR #8 remains the active product stabilization discussion; do not create a duplicate dashboard or overwrite its status as part of an unrelated task.

The triage surface is issues by default. Include pull requests only when the user requests PR triage. Use the canonical mapping in `triage-labels.md`, checking actual remote labels before applying one.

For ticket breakdowns, put the plan in a parent issue and implementation tickets in child issues. Use supported native sub-issue and dependency operations where available. GitHub native relations may require database IDs rather than issue numbers; obtain the correct IDs first. If unsupported, use clearly marked issue links and dependency checklists in Markdown. Do not claim those are native relations. Keep every ticket tied to the approved scope and observable acceptance criteria.

For ordinary GitHub CLI operations, use `gh issue view`, `gh issue list`, `gh issue create` and `gh issue edit` with `--repo Harfoushennium/nmdc-doc-index`. Use structured tool bodies or body files to preserve multiline text. Task-specific authorization and existing repository policy govern external writes.

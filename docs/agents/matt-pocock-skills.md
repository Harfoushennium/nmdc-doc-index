# Matt Pocock skills in Codex Cloud

The 27 published skills from Matt Pocock's package version 1.3.1 are installed in `.agents/skills/`. Each folder contains its original SKILL.md, Codex invocation metadata and supporting files. The upstream MIT license and pinned source/file hashes are in `.agents/mattpocock/`.

## Use in the browser

1. Open Codex Cloud and select the GitHub repository `Harfoushennium/nmdc-doc-index` in your environment.
2. Start a new task using branch `codex/matt-pocock-skills` while this installation is unmerged. After merging, use the branch that contains the installation.
3. Invoke a skill explicitly: `Use the ask-matt skill to help me choose the next workflow`, or `Use the grill-me skill to interview me about this plan`.
4. Supply the task, issue or specification when the workflow asks for it.

Explicit skill selection syntax depends on the client. If `/name`, `$name` or `@name` is passed to the agent as text, the repository instructions route it to the same skill. If the browser intercepts a command or does not offer a picker entry, use `Use the <name> skill ...`. This installation does not add browser slash-command autocomplete or globally install anything in ordinary ChatGPT conversations.

All published skills are present. Several are normally called by another skill; their original invocation policies are preserved. Workflow execution depends on tools available in the selected Cloud environment: authenticated GitHub access, shell/test tools and subagent support where required. Missing tools must be reported rather than simulated. Installation has been verified locally; account-specific Cloud discovery and complete workflow execution require a task in that environment.

## Available skill names

| Skill | Purpose |
| --- | --- |
| `ask-matt` | Find the right skill or workflow |
| `diagnosing-bugs` | Diagnose hard bugs and regressions |
| `grill-with-docs` | Grill a design and write its docs |
| `triage` | Move issues through triage roles |
| `improve-codebase-architecture` | Find and grill architecture improvements |
| `setup-matt-pocock-skills` | Configure a repo for the skills |
| `tdd` | Test-driven red-green-refactor |
| `to-spec` | Turn a conversation into a spec |
| `to-tickets` | Split a plan into tracer-bullet tickets |
| `wayfinder` | Map a large effort as decision tickets |
| `implement` | Build work from a spec or tickets |
| `implement-spec` | Implement the result of /to-spec and /to-tickets in code. |
| `prototype` | Prototype to answer a design question |
| `research` | Research from high-trust sources |
| `domain-modeling` | Build and sharpen a domain model |
| `codebase-design` | Vocabulary for deep-module design |
| `code-review` | Review a diff on standards and spec |
| `pr` | Write a PR body that's fast to review |
| `retro` | Conduct a retrospective on a coding session. |
| `wizard` | Generate an interactive setup wizard |
| `grill-me` | Sharpen a plan through interview |
| `grilling` | Stress-test thinking a round of questions at a time |
| `handoff` | Compact a conversation into a handoff |
| `teach` | Learn a concept in a guided workspace |
| `to-questionnaire` | Front-load questions into a doc for someone to answer |
| `wait-what` | Re-pitch that: simpler, with the context I'm missing |
| `writing-for-agents` | Write documents agents consume |

## NMDC integration

Project configuration lives in `issue-tracker.md`, `triage-labels.md` and `domain.md` in this folder. Existing NMDC scope, read-only DATA rules, PR reporting and owner-only merge requirements in AGENTS.md govern skill execution. Installing a workflow grants no independent permission to implement future product phases, approve or merge PR #8, or provision remote labels.

## Update

The collection is vendored and does not automatically update. To update, select a reviewed upstream commit, copy the full folders listed by its plugin manifest, retain the license, regenerate provenance hashes and verify all references and metadata before opening a PR. Keep project configuration separate from the unmodified upstream skill files.

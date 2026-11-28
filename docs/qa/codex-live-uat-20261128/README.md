# Local Windows UAT continuation — 2026-11-28

Written by: Codex / ChatGPT  
Role: Local Windows acceptance test executor  
Protocol: CHC-NMDC-001 / RP-NMDC-V2-001 / issue #41  
Model and effort: NOT VERIFIED

Status: **Package verification PASS; desktop UAT BLOCKED; application NOT RUN.**

Continues the live assignment in issue #41 comment 7244723642 and the remote
Antigravity capability blocker in comment 7244788785. Start checkpoint:
https://github.com/Harfoushennium/nmdc-doc-index/issues/41#issuecomment-7244854231

## Immutable input and actual verification

| Item | Verified value |
|---|---|
| Candidate PR/head | #45 / `63a523f4001727ecfd8fd34a9bf6edde48666440` |
| Hosted run/artifact | `46176644052` / `15275913015` |
| Outer artifact SHA256 | `d0c0e1fc242b848bdd50abbca6477e16f130c836b8c022be3189a4f0e45ab46c` |
| Inner ZIP SHA256 | `27e8a536009f946c8b5c3d6a0a58acec23c8bd6bbadb6957de790e9e14da748e` |
| Launch.exe SHA256 | `12cfd82e6daac053dcd67ab8574b23a92f03f5e7c12fc330d53cbff43e427048` |
| Manifest build | `e65762f48794046f15944eba2d212a03b8ceeeec` |
| Payload verification | All 1,005 manifest paths, sizes and SHA256 values matched; exact file membership matched |
| Acceptance checklist | 46 unique case IDs verified |
| Local OS | Windows 11, build 26300 |
| PC role | Awaiting owner identification as work or home |

The manifest build is GitHub's temporary PR merge commit, with parents
`569b20a3d0b7b11c84a56a8bd456cddc4d75b8f6` and the candidate head above.
`git diff --quiet 63a523f4001727ecfd8fd34a9bf6edde48666440 e65762f48794046f15944eba2d212a03b8ceeeec`
returned 0: the complete trees match. The workflow uses default actions/checkout
for pull requests. Preserve both identities; do not relabel the manifest build.
Hosted run status was read back as completed/success; hosted tests were not
rerun or represented as local desktop acceptance.

## Desktop blocker CUA-01

Installed Computer Use skill and guidance were read. The supported `@oai/sky`
runtime returned app/window inventories. An attempt to open pre-existing Windows
Explorer through `sky.launch_app` failed with:

```text
GetCursorPos failed: Access is denied. (0x80070005)
```

Fresh window discovery and one retry produced the same error. No targetable
Explorer window appeared. No input was sent to unrelated applications. No
application launch, browser interaction, actual Excel execution, hyperlink click,
migration, approval, reset or benchmark was performed. No screenshot or UI PASS
is claimed. No headless browser/API/COM substitute was used.

Local evidence contains environment.json, results.csv and isolated source hashes.
U01/U02/U03 are capability-blocked, U04 awaits authoritative legacy input;
remaining desktop cases are NOT RUN. Package verification alone does not pass
the complete P04 release case. Source paths and operational evidence remain local.

## Isolation and next action

An isolated non-OneDrive package and source worktree were prepared. The isolated
Git corpus has 56 files; SHA256 mapping fingerprint:
`38e9aec68cb733dbf7acd246a289c8bf9a0228492c6c2311b2e03a023ee48c86`.
No source workbook was opened in Excel or written. The original checkout's two
pre-existing modified DATA files, untracked Copy workbook and skill directory
were preserved. Product branches and immutable candidates remain unchanged.

Resume in an unlocked interactive Windows session with Computer Use input
available; identify work/home role, reverify this same artifact and manifest,
then follow the existing live assignment with an isolated workspace. Do not
start the default operational workspace. Obtain any action-time confirmation
required by the installed Computer Use skill before running newly downloaded
software through the UI. Authoritative legacy selection, other-PC access,
physical Excel, repeated performance measurements and owner sign-off remain open.

Fetch succeeded but automatic Git maintenance reported mmap/repack and stale
worktree cleanup errors. Subsequent object reads, exact tree comparison and
working-tree status succeeded. No manual pruning, history rewrite or deletion
was attempted. Verify publication separately at handoff.

No owner acceptance, issue closure, production merge or cutover is authorized
or claimed by this evidence slice.

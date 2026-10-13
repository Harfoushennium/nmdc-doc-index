# File review and optional change details

Owner confirmations: 13 October 2026. This replaces the previous same-sheet source-scope review requirement.

The owner reviews affected files on File Review and may open Change Details when record context is useful. The record preview remains available, but is not a prerequisite for normal acceptance. Both areas are Excel Tables.

Each affected file has Accept Update?, Project No., Source File, Change Type, Added, Modified, Removed and Owner Note. Unchanged files are omitted. Checked accepts the proposed update; unchecked keeps that file's approved contents for this proposal. An unchecked new file contributes no records. An unchecked removed file retains its approved records. This is distinct from a permanent source exclusion.

Save File Choices binds all file choices and notes to the exact displayed run and approved baseline. It does not rescan. Approve Update on File Review requires saved choices; unsaved review decisions and existing conflict guards still block promotion. Home opens file review. View Change Details opens the optional record table.

Partial approval combines accepted staged records with deferred approved records, manifest entries and caches. Deferred files retain their parser/configuration metadata so later scans cannot mistake skipped work for a completed upgrade. Duplicate keys in the combined result block approval before the atomic approved pointer changes. Saved choices from replaced or altered proposals are rejected. The version records the file choices for audit.

Workbook drafts use a separate FILES section keyed by run and source path. Historical SOURCE and STATE journal rows remain readable for review but are not replayed as file acceptance. Saving reconciles the live worksheet against the new saved exchange baseline; it does not clear a whole draft section and lose edits made while the engine was running. Replacement-run drafts remain available without cross-application.

## Regression map

| Area | Problem or requirement | Change | Verification |
|---|---|---|---|
| Review layout | Record preview obscures the owner's file decisions | Separate File Review and optional Change Details | Installer/VBA contracts; real Excel pending |
| Checkbox meaning | Existing checkbox persists source exclusions | Proposal-bound save-file-review command | Mixed/all-deferred/new/removed engine fixtures |
| Subsequent scans | Partial parser/config approval could mark skipped files current | Preserve per-file processing metadata | Parser/config upgrade regression fixtures |
| Promotion | Retained and newly accepted records can collide | Validate combined record keys before promotion | Duplicate-key negative fixture |
| Drafts | Save or replacement must not discard edits | Run-bound FILES journal and baseline reconciliation | Workbook contract checks; updated native fixture harness not executed |
| Source safety | Existing excluded records must remain absent | Preserve persistent scope engine APIs and stage contents during save | Disposable 40,494-row CLI export/save/reload; 47 affected files, 9 excluded sources absent |
| Real Excel | Current screen capture is black | Desktop input paused; unlock requested | No native acceptance claimed for this change |

The fresh candidate still requires real Excel compilation, controls, typing/caret, column selection, shortcut behavior, saved/reopened choices and notes, unsaved decision guards, prior draft review, and disposable approval checks. The blanket worksheet-data-in-Tables gate and the complete 23-case/two-install gate remain open. PR #8 remains CHANGES and unmerged.

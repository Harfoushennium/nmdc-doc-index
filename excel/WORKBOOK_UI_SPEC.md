# NMDC Document Index — Workbook UI Specification

## Product rule

The normal user operates the product from Excel only. Technical runtime details stay hidden unless the user opens the Error / Debug Log or exports a support package.

## Sheet order

1. `Home`
2. `Master Documents`
3. `Revisions`
4. `Transactions`
5. `File Review`
6. `Change Details`
7. `Review Flags`
8. `User Decisions`
9. `Configuration`
10. `Rules & Mappings`
11. `Custom Fields`
12. `Update History`
13. `Error Log`
14. `Help`
15. `System Data` (hidden)

## Home

The Home page is the control panel.

### Status cards

- Approved index status
- Current approved version / run ID
- Last successful update
- Current data folder
- Total approved documents
- Total revisions
- Total transactions/events
- Pending update status
- Review flags count
- Conflict flags count

### Main buttons

- **Update Changed Files**
- **Full Rescan / Rebuild All**
- **Review File Updates**
- **Approve Update**
- **Hold Update**
- **Reject Update**
- **Select Data Folder**
- **Refresh Dashboard**
- **Review Flags**
- **Flag Wrong Data**
- **Configuration**
- **Rules & Mappings**
- **View Log**
- **Report Requirement / Problem**
- **Help**

## Master Documents

One row per document. Preferred columns:

- Flag Level
- Project No.
- Source Family
- Discipline
- Category
- Subcategory
- Document No.
- Document Title
- Company Document No.
- Latest Revision
- Latest Event Date
- Latest Event Status
- Document Link
- Source File
- Source Sheet
- Source Row
- Source Cell
- Global Document Key

The user should be able to filter/sort without exposing internal formulas.

## Revisions

One row per document revision.

Preferred columns:

- Flag Level
- Project No.
- Document No.
- Document Title
- Revision
- Is Latest Revision
- Revision Key
- Source File
- Source Sheet
- Source Row
- Source Cell

## Transactions

Lossless event/history view. One row per extracted event/transaction.

Preferred columns:

- Flag Level
- Project No.
- Document No.
- Document Title
- Revision
- Event Type
- Event Date
- Event Reference
- Event Status
- Event Values
- Is Latest Event
- Document Link
- Source File
- Source Sheet
- Source Row
- Source Cell
- Event Key

## File Review

Primary owner review on the `FileReview` Excel Table. Show affected files only, one row per Source File: Accept Update?, Project No., Source File, Change Type, Added, Modified, Removed, Owner Note.

Checked accepts a file update; unchecked retains its previously approved data for this proposal, including removed files. Deferred new files add no records. This choice does not update permanent source exclusions. Save File Choices is bound to the exact displayed run and approved baseline and does not rescan. Approval is on this sheet and requires saved file choices, saved review decisions and the existing conflict checks. View Change Details opens the optional record preview. Review Saved Drafts exposes old-run edits without applying them to the new run.

## Change Details

Optional read-only record review of the staged update. Normal owner review and acceptance are on File Review.

**A staged update never replaces the approved master automatically.** Only the explicit **Approve Update** action may promote a staged dataset to the approved index.

Columns:

- Change Type (`ADDED`, `MODIFIED`, `REMOVED`, `UNCHANGED`)
- Record Identity
- Project No.
- Document No.
- Revision
- Event Type
- Source File
- Plain-English Summary
- Review Required

The sheet must show a top summary of Added / Modified / Removed / Unchanged counts.

## Review Flags

Columns:

- Flag Level (`OK`, `REVIEW`, `CONFLICT`)
- Flag Code
- Plain-English Problem
- Recommended User Action
- Project No.
- Document No.
- Revision
- Source File
- Source Sheet
- Source Row
- Source Cell
- User Decision
- User Comment
- Resolution Status
- Event Key

Buttons:

- **Flag Wrong Data**
- **Create Support Request**
- **Mark Reviewed**

Source-level flags may legitimately leave document/revision fields blank. Record-level flags must carry all available record context so the affected row can be identified and corrected without guessing.

## User Decisions

Auditable owner/user decisions and overrides.

Columns:

- Decision ID
- Date/Time
- User
- Decision Type
- Source File
- Project / Document / Revision
- Previous Value
- Approved Value
- User Note
- Status
- Configuration Version

## Configuration

Editable cells are clearly marked. System/formula cells are protected.

Required configuration:

- Data Folder
- Runtime Folder
- Engine Executable Path
- Classification Rules File
- Project Identity Overrides File
- Parser Version
- Configuration Version
- Optional hyperlink scan root (future)
- UI preferences

The system should favor editable configuration tables over hard-coded business rules.

## Rules & Mappings

This sheet is the non-programmer control surface for business rules that should not require a code change.

The workbook should present user-editable tables for items such as:

- classification keyword / match rule;
- discipline;
- category;
- subcategory;
- status mappings;
- project identity overrides;
- permitted flag-severity overrides;
- source-selection rules where safely configurable.

Example:

| Match / Condition | Discipline | Category | Subcategory | Enabled |
|---|---|---|---|---|
| Anchor Pattern | MARINE OPERATIONS | DRAWING | ANCHOR PATTERN | YES |
| Installation Procedure | OFFSHORE INSTALLATION | PROCEDURE | INSTALLATION PROCEDURE | YES |
| Sketch | OFFSHORE INSTALLATION | SKETCH | ENGINEERING SKETCH | YES |

Rules are validated before they are exported to the engine configuration. Invalid or contradictory rules are flagged instead of being silently accepted.

When an approved rule/mapping changes:

1. increment the configuration version/fingerprint;
2. record who/when changed it in Update History / User Decisions;
3. require the next update to reprocess affected sources (or all sources when the affected scope cannot be safely narrowed);
4. create a staged update for user review before the approved index changes.

A requirement that cannot be expressed safely through these editable tables must be reported through **Report Requirement / Problem** and treated as a parser/code change rather than approximated.

## Update History

One row per run / decision:

- Run ID
- Date/Time
- Mode (`INCREMENTAL` / `FULL_RESCAN`)
- Decision (`STAGED` / `APPROVED` / `HELD` / `REJECTED`)
- New Sources
- Changed Sources
- Removed Sources
- Added Records
- Modified Records
- Removed Records
- Review Flags
- Conflict Flags
- Note

## Error Log

Plain-English first, technical detail second.

Columns:

- Date/Time
- Severity
- Action
- Plain-English Error
- Recommended Action
- Technical Detail
- Run ID
- Source File
- Worksheet
- Source Row/Cell

Local Excel/VBA errors and engine/runtime errors must both be preserved. Refreshing engine-exported errors must not erase locally captured Excel errors.

## Help

Short non-technical instructions:

1. Select the data folder.
2. Press Update Changed Files.
3. Review File Updates and Review Flags.
4. Resolve/hold conflicts.
5. Press Approve Update when satisfied.
6. Use Full Rescan when rules/parser change or when you want a complete rebuild.
7. Use Rules & Mappings for supported business-rule changes that do not require programming.
8. Use Report Requirement / Problem to create a support package for GPT/developer follow-up.

## Button-to-engine commands

| Excel button | Engine command |
|---|---|
| Update Changed Files | `stage --mode incremental` |
| Full Rescan / Rebuild All | `stage --mode full` |
| Refresh Dashboard | `export-excel` |
| Approve Update | `approve` |
| Hold Update | `hold` |
| Reject Update | `reject` |
| Flag Wrong Data | `user-flag` |
| Report Requirement / Problem | `support-request` |

`Select Data Folder` is handled locally by Excel/VBA and saved into workbook configuration; it is not an engine command.
`Configuration`, `Rules & Mappings`, and `Help` are local Excel navigation actions and do not invoke the engine.

The production approval call must bind to the exact pending run displayed and reviewed in Excel: `approve --run-id <displayed run ID> --require-file-review`. Before approval, Excel refreshes the exchange data and stops if the current pending run ID differs from the reviewed run ID.

When the displayed run contains conflict flags, Excel requires a separate explicit warning/confirmation and may pass `--allow-conflicts`. The engine must still refuse non-overridable parser, source-hash, and duplicate-record-key conflicts. This permits an owner to intentionally accept an overridable business event such as a confirmed source removal without weakening technical data-integrity blocks.

## Exchange-data integrity

- Excel imports exchange CSV columns as text by default so identifiers and revisions are preserved exactly. Values such as revision `00`, document numbers with leading zeroes, and event references such as `1-2` must not be converted by Excel.
- The engine is responsible for intentional date normalization before export. The Excel refresh must not reinterpret identifier-like values as dates or numbers.
- If `export-excel` fails, Excel must stop the action, open the Error Log, and must not announce or open a stale File Review as if it were current.
- Review File Updates and Review Flags refresh the exchange data before opening their sheets.
- An approval is allowed only for the exact pending run ID that remains current after a successful refresh.

## Wrong-data context

When **Flag Wrong Data** is used on a selected record, Excel passes all context available on that row, including source file/sheet/row/cell, document number, revision, event identity, flag code, selected field/current value, expected value, user note, and user name. The engine adds the approved/pending run IDs plus parser and configuration versions to the stored support record.

## Visual behavior

- Green = approved / OK.
- Amber = review required.
- Red = conflict / blocked.
- Blue = navigation/action.
- Grey = system/read-only.
- User-editable cells are visually distinct.
- Freeze panes on all long tables.
- Use Excel Tables and filters.
- Keep technical identifiers available but place them to the right, after user-facing columns.

## Runtime error behavior

If the engine executable is missing, blocked, or returns a failure code:

1. Do not modify approved workbook data.
2. Show a plain-English message box.
3. Write an Error Log entry.
4. Preserve the technical error detail for support.
5. Offer the user the Report Requirement / Problem workflow.
6. Never continue to an older review table or approve a run that was not the one displayed and reviewed.

# NMDC Document Index — User Product Requirements

## Product principle

The final product is for a non-programmer. The user should work from Excel only. The user should not need to open Python, GitHub, PowerShell, Command Prompt, or any other technical interface for normal operation.

## Final user interface

The main product shall be an Excel workbook with clearly labelled buttons for all routine actions. At minimum the workbook should provide:

- **Update Changed Files** — compare source metadata/hashes, re-process only new/changed source files, rebuild the master dataset from cached unchanged data, and create a staged update for review.
- **Full Rescan / Rebuild All** — ignore the cache and re-read all eligible source workbooks from scratch.
- **Review Pending Update** — show additions, changes, removals, warnings, conflicts, and affected projects/documents before they enter the approved master index.
- **Approve Update** — only after user review, promote the staged dataset into the approved master dataset.
- **Reject / Hold Update** — keep the current approved master unchanged and retain the proposed update for review/correction.
- **Select Data Folder** — allow the user to choose a source-data folder located anywhere accessible on the PC/network.
- **Review Flags** — open/filter all outliers and conflicts with plain-English explanations.
- **Flag Wrong Data** — let the user mark a record as wrong, add a comment, and route it to the review/fix queue.
- **Report a Requirement / Problem** — allow the user to type a plain-English message for GPT/developer follow-up without editing code.
- **View Log** — show update history, errors, bugs, warnings, user decisions, and technical diagnostics in a readable form.

## Update approval workflow

Routine updates must be staged before they replace the approved master index.

1. User presses **Update Changed Files** or **Full Rescan / Rebuild All**.
2. System scans the selected source folder.
3. System creates a proposed/staged update.
4. System compares the staged result with the currently approved master dataset.
5. Excel shows a human-readable review screen: added / modified / removed / flagged / conflicted records.
6. User either **Approves**, **Rejects**, or **Holds** the update.
7. Only an approved update becomes the new master dataset.
8. The approval action, date/time, update ID, affected sources, and summary are written to the audit log.

No staged update may silently overwrite the approved master index.

## Incremental update logic

The system should maintain a source manifest containing at minimum:

- source path
- file name
- file size
- SHA-256 content hash
- modified date/time
- source-selection status
- parser/configuration version
- last successfully processed run
- user-approved project identity override, if any

Normal update behaviour:

- **NEW** source -> parse it.
- **CHANGED** source (content hash changed) -> re-parse the complete workbook.
- **UNCHANGED** source -> reuse the previous canonical extraction.
- **REMOVED** source -> stage retirement/removal of records previously produced from that source; do not delete approved data until the user approves the update.
- After processing changed/new sources, rebuild the staged master dataset from the complete canonical cache so the master remains deterministic.

The **Full Rescan / Rebuild All** button must re-parse every eligible source file regardless of metadata/cache state.

## Flag and conflict system

Every record can carry a simple user-facing flag level:

- 🟢 **OK** — no known issue.
- 🟡 **REVIEW** — unusual/incomplete data that may still be usable.
- 🔴 **CONFLICT** — data must not be accepted automatically.

Detailed machine codes can exist underneath, but Excel must show a plain-English description and the recommended user action.

Examples include:

- PROJECT_CONFLICT
- DUPLICATE_DOCUMENT
- REVISION_SEQUENCE_GAP
- FUTURE_DATE
- MISSING_DOCUMENT_NUMBER
- MULTIPLE_PROJECT_NUMBERS
- UNRECOGNIZED_LAYOUT
- BROKEN_HYPERLINK
- SOURCE_CHANGED_AFTER_LAST_RUN
- UNCLASSIFIED_DOCUMENT
- USER_FLAGGED_WRONG_DATA
- APPROVED_OVERRIDE_CONTRADICTED
- PARSER_ERROR

The Flags/Review sheet must allow filtering by flag level, project, source file, document number, flag type, and status.

## User correction and feedback

The user must be able to flag a wrong record directly in Excel and add a note. The system should distinguish:

- **Data correction** — can be resolved by an owner-approved override/configuration rule without code change.
- **Configuration change** — can be resolved by editing a user-editable rule/configuration table.
- **Code change required** — parser or logic does not support the source correctly.

When code change is required, Excel should prepare a simple support package/report containing:

- update/run ID
- source file
- worksheet
- source row/cell when available
- document/revision/event identity
- error/flag code
- plain-English description
- user comment
- relevant configuration version
- parser version

This package should be easy to copy/upload to GPT so the issue can be reproduced and fixed.

## Dynamic configuration

Business rules must be editable without modifying source code wherever practical. Configuration should be exposed in protected/user-editable Excel sheets and/or simple CSV tables that the Excel UI manages.

Examples:

- classification rules
- source-selection rules
- project identity overrides
- status mappings
- discipline/category/subcategory mappings
- flag thresholds and severities
- data-folder path
- hyperlink root paths
- folders excluded from scanning

Changing configuration should increment a configuration version and be recorded in the log.

## Data-folder selection

The source-data directory must not be hard-coded. The Excel UI must allow the user to select a local, OneDrive, shared-drive, or network folder accessible to Windows. The current selected path must be visible in Excel and stored as configuration.

## Future feature — folder/file scan for hyperlink enrichment

After the main index system is stable, add a feature that can scan a user-selected project/document directory recursively and create a file/folder inventory (CSV/canonical table). The system will use that inventory to match procedures, drawings, sketches, reports, etc. to their actual file/folder locations and populate/repair hyperlinks in the index.

Required characteristics:

- user selects the root folder from Excel
- recursive scan of subfolders
- capture path, file/folder name, extension/type, size, modified date, and available metadata
- generate a reviewable scan table
- proposed hyperlink matches must be reviewable before approval
- ambiguous matches must be flagged, not guessed
- this feature is **future/pinned** and should be implemented after the core index/update workflow is stable

## Excel usability

The final workbook should include at least:

- Home / Dashboard
- Master Documents
- Revisions
- Transactions / Events
- Pending Update
- Review / Flags
- User Decisions / Overrides
- Configuration
- Update History
- Error / Debug Log
- Help

Expected usability features:

- frozen headers
- filters and tables
- clear buttons
- simple status colours/icons
- hyperlinks
- protected formula/system areas
- editable user-input areas clearly distinguished
- plain-English messages
- no requirement to edit code for normal use

## Runtime / prerequisites target

Preferred deployment target:

- Windows PC
- Microsoft Excel desktop with macros/VBA allowed
- no Python installation required for the user
- no Git installation required for the user
- no command-line use required for the user

The preferred technical approach is an Excel `.xlsm` front end that calls a packaged local engine/executable silently in the background. The engine should be distributed with the workbook and should not require a separate Python installation. If corporate policy prevents macros or local executables, an alternative deployment must be agreed before final packaging.

## Governance / GitHub naming

Canonical development branch names must start with a sequence number, e.g. `01-...`, `02-...`, `03-...`.

PR titles must always include the final/current status and sequence, for example:

- `[DRAFT][06] ...`
- `[READY][06] ...`
- `[CHANGES-REQUESTED][06] ...`
- `[MERGED][06] ...`
- `[CLOSED-NOT-MERGED][06] ...`

When a PR reaches a final state, add a final status comment that clearly states:

- final status
- merge/close result
- exact final head SHA
- merge commit SHA when applicable
- tests/validation result
- what phase comes next

## Owner decisions on project identity — 2026-09-11

The owner reviewed the previously flagged conflict files and approved these project identities:

- `DATA/TECH/2035 DOCUMENT REGISTER.xlsx` -> **2035**. File/project document-number evidence is authoritative; prior inferred 2136 evidence is rejected as false/incorrect.
- `DATA/TECH/2035 Offshore Construction Engineering Register (Pipeline & Cables).xlsx` -> **2035**. File/project document-number evidence is authoritative; prior inferred 2136 evidence is rejected as false/incorrect.
- `DATA/TECH/2705 -DOCUMENT REGISTER.xlsx` -> **2705**. Filename/project context is authoritative; internal project number shown in the workbook is considered wrong.
- `DATA/TECH/2745-PP-GE-001-MDR Rev_2.xlsx` -> **2745**. Filename/project context is authoritative; conflicting internal evidence is considered wrong.

These are explicit owner-approved overrides. Future automatic inference must not silently replace them. If future file content strongly contradicts an approved override, raise `APPROVED_OVERRIDE_CONTRADICTED` for user review rather than changing project identity automatically.

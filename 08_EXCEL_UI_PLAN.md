# 08 — Excel User Interface & Engine Integration

## Owner goal

The normal user must operate the NMDC Document Index from Excel only. No routine action should require Python, Git, PowerShell, Command Prompt, GitHub, or direct editing of runtime files.

## Product boundary for this PR

This PR builds the Excel-facing contract and source components for the future macro-enabled production workbook. It includes:

1. workbook sheet / table / button specification;
2. VBA source modules for button actions and navigation;
3. a stable Excel-to-engine command contract;
4. a Python bridge that converts runtime state into Excel-friendly CSV exchange files;
5. tests for the bridge and safety contract;
6. a user-reviewable Excel UI prototype generated outside the repo binary path.

The final signed/packaged `.xlsm` binary and packaged Windows executable are still separate packaging/deployment work because the repository needs auditable text sources and CI cannot execute desktop Excel macros.

## Required workbook sheets

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

## Required buttons

- Update Changed Files
- Full Rescan / Rebuild All
- Review Pending Update
- Approve Update
- Hold Update
- Reject Update
- Select Data Folder
- Refresh Dashboard
- Review Flags
- Flag Wrong Data
- Report Requirement / Problem
- View Log

## User safety rules

- A staged update never replaces the approved master automatically.
- Approve requires an explicit user action.
- Hold / Reject leaves the approved master unchanged.
- Technical non-overridable conflicts remain blocked.
- Source workbooks stay read-only.
- Buttons call the local engine silently; no command window is part of the user workflow.
- If the engine executable is missing or blocked by IT, Excel shows a plain-English error and records it in the Error Log.

## Excel / engine exchange contract

Excel will call the packaged engine with commands and an exchange folder. The engine writes deterministic CSV files that Excel can load without a JSON parser dependency.

Expected exchange files:

- `dashboard.csv`
- `master_documents.csv`
- `revisions.csv`
- `events.csv`
- `pending_update.csv`
- `flags.csv`
- `history.csv`
- `errors.csv`

## Dynamic configuration

The workbook will expose user-editable configuration for:

- data folder path;
- classification rules reference/version;
- project identity overrides;
- flag severity overrides where permitted;
- status mappings;
- hyperlink roots / scan roots (future feature);
- UI preferences.

Changes that require parser/code changes must not be silently approximated. The user can create a support request from Excel for GPT/developer follow-up.

## Future pinned feature

After the core workbook/runtime is stable, add the project-folder/file scanner that inventories a user-selected root folder and proposes hyperlinks to procedures, drawings, sketches, reports, etc. Ambiguous matches stay in REVIEW rather than being guessed.

## Acceptance criteria

- Workbook contract documented.
- VBA button source is auditable and no shell window is required for normal operation.
- Excel bridge produces deterministic CSV exchange files from approved/staged runtime state.
- Bridge supports no-approved-state, approved-state, staged-state, flags, history and user support requests.
- Tests pass on Linux and Windows through the existing repository regression workflows.
- `DATA/` remains unchanged.
- No production `.xlsm` is claimed complete until macro injection/packaging is actually validated on Windows Excel.
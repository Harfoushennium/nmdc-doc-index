> Owner-confirmed workflow (13 Oct 2026): File Review is the primary affected-file Table; Change Details is optional. Unchecked files keep their approved contents and do not change persistent source scope. Save File Choices does not rescan. Approval is on File Review and requires saved choices and review decisions. Real Excel acceptance is still required.

# NMDC Document Index production package

## User setup

1. Extract the complete ZIP to a normal Windows folder.
2. Double-click `Create_NMDC_Document_Index.vbs` once.
3. Open the created `NMDC_Document_Index.xlsm` in Microsoft Excel desktop.
4. Enable macros when Excel asks.
5. Press **Select Data Folder** and choose the folder containing `METHODS` and `TECH`.
6. Press **Update Changed Files** or **Full Rescan / Rebuild All**.
7. Review **File Review** and **Review Flags** before approving.

The one-time setup uses Microsoft Excel to attach all production VBA modules and workbook events. If Excel blocks that attachment, setup displays the exact Trust Center setting needed and stops without changing the source workbook.

All NMDC-created state stays inside this extracted package folder. The workbook uses relative paths:

- engine: `engine\nmdc_index_engine.exe`
- runtime/cache/staging: `runtime`
- configuration: `config`
- parser identity: `cycle3-extractor-v2`

The production workbook does **not** configure AppData or another unrelated user folder for NMDC runtime/support data.

## Dynamic Live Filter — owner REV03 design

The Live Filter follows the owner's REV03 reference architecture.

- Each supported table sheet uses a real worksheet ActiveX textbox named `TxtBox_Search`.
- `Cls_LiveFilter_Listener` handles the textbox `MouseDown`, `Change`, and `KeyDown` events.
- Normal typing goes directly into the textbox with a visible caret.
- Filtering updates on every textbox `Change` event.
- Press `Ctrl+Shift+F` only to choose the target **table header**.
- The selected target is stored in the hidden workbook name `LiveFilter_Anchor`.
- Placeholder/help text is:
  `Search <Column>... (+AND / -EXCLUDE) (Ctrl+Shift+F: Change Column)`
- Space or `+` means required/AND terms.
- `-word` excludes rows containing that term.
- Text wrapped in double quotes uses the reference exact-match behavior.
- **RESET SEARCH** clears the table filter and restores the placeholder.

Setup registers Microsoft Forms 2.0 before importing the listener class. ActiveX search boxes are created lazily only after the workbook is opened in normal visible Excel; setup does not create them while Excel is hidden under VBS automation.

## File acceptance on File Review

**File Review** contains affected files; **Change Details** holds the optional record preview.

- **Accept Update?** uses the modern Microsoft 365 native in-cell Checkbox control.
- Checked = accept this file update.
- Unchecked = retain its previously approved data for this proposal.
- Add an optional Owner Note when useful.
- Click **Save File Choices** once after making the choices.
- **Check All** / **Uncheck All** are available for bulk selection.

The owner environment supports native in-cell checkboxes. In this production build, failure to create these controls is treated as a visible error; the UI does not silently downgrade to visible TRUE/FALSE text.

Source workbooks are never edited or deleted by this choice.

## Keep and save review inputs

Source choices, Owner Notes and Review Flag inputs remain drafts until saved.
Refresh keeps those edits, and saving the workbook before closing keeps them
available when it is reopened. Owner Notes are retained for accepted files
as well as deferred files.

- Use **Save File Choices** to bind choices to the current proposal without rescanning.
  Approval stays blocked until all file choices and review decisions are saved.
- Use **Save Review Decisions** to record flag decisions, comments and status.
  Clearing a previously saved field is also a review change that must be saved.
- **Select?** chooses flags for a parser-fix report; this checkbox alone does
  not block approval.
- Use **Review Saved Drafts** to inspect edits retained from a replaced proposal.
  Earlier-proposal edits are displayed for review and are not applied to the
  new proposal. Enter any changes you still want in the current proposal and save them.

If Excel cannot safely read or retain a draft, refresh and approval stop and
the Error Log explains the failure. Keep the workbook and its package folder
together when saving or moving your work.

## Review Flags and parser/mapping fixes

`Review Flags` is reserved for genuine actionable extraction/data anomalies.

The first column is **Select?**, implemented with the same native Microsoft 365 in-cell checkbox control.

To report extraction that needs parser or mapping correction:

1. Open the flagged **Source File** and **Worksheet Name** and confirm normal data is being missed.
2. Tick **Select?** for one or more exact Review Flag rows.
3. Click **Report Selected Parser Fix**.
4. Add an explanation when useful; existing row comments are preserved.
5. The selected rows are recorded as `NEEDS PARSER/MAPPING FIX` and remain `OPEN`.
6. The workbook creates a dedicated `PARSER_FIX_REPORTS` folder beside `NMDC_Document_Index.xlsm`.
7. Every request uses the next sequential subfolder: `0001`, `0002`, `0003`, and so on.
8. Each numbered folder contains only `PARSER_FIX_REQUEST.md` and `PARSER_FIX_REQUEST.json` for that request.
9. Windows Explorer opens with that exact numbered request's Markdown report selected.
9. Give that report plus the affected source workbook(s) to ChatGPT / the project maintainer.
10. After a corrected parser/configuration is installed, click **Retry After Fix**.

**Select All** and **Clear Selection** operate on the Review Flags `Select?` column.

Saving a decision alone does not prove the extraction is fixed. Resolution is confirmed only when the corrected Full Rescan no longer reproduces the Review Flag.

## Custom Fields & Keywords

Use **Custom Fields & Keywords** from Home to add user-defined derived columns to Master Documents, for example `Vessel Names`. Define which Master Documents columns to search, then maintain the editable keyword/result dictionary. Core extracted NMDC fields are protected from overwrite.

Binary **Enabled?** choices use native Microsoft 365 in-cell checkboxes. Multi-choice settings continue to use normal Excel dropdowns. Package-local runtime backups are kept under the extracted package's `runtime` folder.

## Background scanning

The responsive scan launcher keeps the engine in a hidden Windows process and polls it from Excel. Scheduled callbacks are explicitly qualified with the local open workbook name, not a SharePoint/OneDrive URL.

## Safety

- Source registers are read-only.
- Every update is staged first.
- Hold and Reject do not change the approved index.
- Approve binds to the exact staged run shown in Excel.
- Reset clears indexed/staged runtime records only; it does not delete source DATA or configuration.
- Undo restores the previous approved index version when one exists.
- Parser, hash, and duplicate-record-key conflicts cannot be overridden.
- Existing production workbooks are backed up before replacement.
- PR #8 must remain unmerged until owner desktop acceptance is complete.

The package is not digitally signed. Corporate macro and executable policy may require IT approval before use.

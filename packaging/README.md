# NMDC Document Index production package

## User setup

1. Extract the complete ZIP to a normal Windows folder.
2. Double-click `Create_NMDC_Document_Index.vbs` once.
3. Open the created `NMDC_Document_Index.xlsm` in Microsoft Excel desktop.
4. Enable macros when Excel asks.
5. Press **Select Data Folder** and choose the folder containing `METHODS` and `TECH`.
6. Press **Update Changed Files** or **Full Rescan / Rebuild All**.
7. Review **Pending Update** and **Review Flags** before approving.

The one-time setup uses Microsoft Excel to attach all production VBA modules and workbook events. If Excel blocks that attachment, the setup displays the exact Trust Center setting needed and stops without changing the source workbook.

The generated workbook keeps the last-working owner controls — including **Reset All Records**, **Undo Last Approval**, Review/Approve/Hold/Reject and Help — and adds the Live Filter, integrated Pending Update source-selection panel, and Custom Fields & Keywords features. The setup explicitly embeds these modules rather than relying on a later opportunistic import.

### Dynamic Live Filter

Live Filter is available on the main data/review sheets and follows the owner-provided Dynamic Live Filter behavior:

1. Click **SELECT COLUMN** or press `Ctrl+Shift+F`.
2. Click the table header of the one column you want to search.
3. Type into the Live Filter text box. The visible rows update immediately with every keystroke; Enter/Tab is not required.
4. Normal text is a partial, case-insensitive search. Text wrapped in double quotes is an exact, case-sensitive search.
5. Click **RESET** to clear the search/filter.

There is no `ALL COLUMNS` mode and no hidden helper-column scan. The filter operates directly on the selected Excel Table column using the native AutoFilter path for normal searches.

### Source selection in Pending Update

Source selection is now part of **Pending Update** instead of a separate user-facing worksheet. The source panel on the right shows one row per source workbook, including **Project No.** and **Source File**.

- Checked = include the source in index scope.
- Unchecked = intentionally exclude the source.
- Add an optional Owner Note when useful.
- Click **Save Source Choices & Restage** once after making the choices.

Checkboxes are linked directly to their cells and do not call an individual checkbox macro. This avoids the previous `NMDC_SourceCheckboxClicked` macro-availability error. Source workbooks are never edited or deleted by this choice.

### Custom Fields & Keywords

Use **Custom Fields & Keywords** from Home to add user-defined derived columns to Master Documents, for example `Vessel Names`. Define which Master Documents columns to search, then maintain the editable keyword/result dictionary. Core extracted NMDC fields are protected from overwrite.

The package keeps runtime/cache under `%LOCALAPPDATA%` when available so normal operation does not create unnecessary sync activity inside a OneDrive-hosted package folder. If you move or rename the extracted package later, rerun setup from its new location so package paths are refreshed.

## Safety

- Source registers are read-only.
- Every update is staged first.
- Hold and Reject do not change the approved index.
- Approve binds to the exact staged run shown in Excel.
- Reset clears indexed/staged runtime records only; it does not delete source DATA or configuration.
- Undo restores the previous approved index version when one exists.
- Parser, hash, and duplicate-record-key conflicts cannot be overridden.
- Existing production workbooks are backed up before replacement.

The package is not digitally signed. Corporate macro and executable policy may require IT approval before use.

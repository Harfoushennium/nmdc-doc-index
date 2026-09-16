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

Live Filter is available on the main data/review sheets without ActiveX controls.

1. Click **SELECT COLUMN** or press `Ctrl+Shift+F`.
2. Click the table header of the one column you want to search.
3. Live Filter enters typing mode. Type normally and the visible rows update immediately after every character; Enter/Tab is not required to apply the filter.
4. **Backspace** edits the search text. **Delete** clears the current query. **Esc**, **Enter** or **Tab** exits typing mode.
5. Click the search display to resume typing for the currently selected column.
6. Normal text is a partial, case-insensitive native Excel Table filter. Text wrapped in double quotes performs an exact whole-cell match.
7. Click **RESET** to clear the filter.

The implementation intentionally does **not** use an ActiveX `Forms.TextBox.1`, because current Microsoft 365 security settings can block ActiveX controls. It also has no `ALL COLUMNS` helper scan. The filter operates directly on the one selected Excel Table column using native AutoFilter.

### Source selection in Pending Update

Source selection is part of **Pending Update** instead of a separate user-facing worksheet. The source panel on the right shows one row per source workbook, including **Project No.** and **Source File**.

- The **Include in Index?** column uses the modern Microsoft 365 in-cell **Checkbox** control (`CellControl.SetCheckbox`).
- Checked = include the source in index scope.
- Unchecked = intentionally exclude the source.
- Add an optional Owner Note when useful.
- Click **Save Source Choices & Restage** once after making the choices.

The checkbox is the Excel cell control itself: the cell stores `TRUE`/`FALSE`; no Form Control or ActiveX checkbox and no per-checkbox macro are required. Source workbooks are never edited or deleted by this choice.

### Custom Fields & Keywords

Use **Custom Fields & Keywords** from Home to add user-defined derived columns to Master Documents, for example `Vessel Names`. Define which Master Documents columns to search, then maintain the editable keyword/result dictionary. Core extracted NMDC fields are protected from overwrite.

Binary **Enabled?** choices use the same modern Excel in-cell Checkbox control. Multi-choice settings continue to use normal Excel dropdowns.

The package keeps runtime/cache under `%LOCALAPPDATA%` when available so normal operation does not create unnecessary sync activity inside a OneDrive-hosted package folder. If you move or rename the extracted package later, rerun setup from its new location so package paths are refreshed.

### Background scanning

The responsive scan launcher keeps the engine in a hidden Windows process and polls it from Excel. Scheduled callbacks are explicitly qualified with the local open workbook **name**, not a SharePoint/OneDrive URL. This avoids the earlier macro-resolution failure where Excel attempted to execute `NMDC_PollEngineAsync` through a SharePoint address.

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

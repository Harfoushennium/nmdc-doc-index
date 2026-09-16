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

The generated workbook keeps the last-working owner controls — including **Reset All Records**, **Undo Last Approval**, Review/Approve/Hold/Reject, Source Selection and Help — and adds the newer Live Filter and Custom Fields & Keywords features. The setup explicitly embeds these newer modules rather than relying on a later opportunistic import.

### Live Filter

Live Filter is available on the main data/review sheets. Type in the row-3 search cell and press Enter/Tab. Spaces or `+` mean AND, `-word` excludes a term, and a quoted whole search performs an exact case-sensitive match. `Ctrl+Shift+F` selects one target column; otherwise use `ALL COLUMNS`.

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

# NMDC Document Index production package

## User setup

1. Extract the complete ZIP to a normal Windows folder.
2. Double-click `Create_NMDC_Document_Index.vbs` once.
3. Open the created `NMDC_Document_Index.xlsm` in Microsoft Excel desktop.
4. Enable macros when Excel asks.
5. Press **Select Data Folder** and choose the folder containing `METHODS` and `TECH`.
6. Press **Update Changed Files** or **Full Rescan / Rebuild All**.
7. Review **Pending Update** and **Review Flags** before approving.

The one-time setup uses Microsoft Excel to attach the audited VBA modules to the
workbook. If Excel blocks that attachment, the setup displays the exact Trust
Center setting needed and stops without changing the source workbook.

## Safety

- Source registers are read-only.
- Every update is staged first.
- Hold and Reject do not change the approved index.
- Approve binds to the exact staged run shown in Excel.
- Parser, hash, and duplicate-record-key conflicts cannot be overridden.
- Existing production workbooks are backed up before replacement.

The package is not digitally signed. Corporate macro and executable policy may
require IT approval before use.

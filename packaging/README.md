# NMDC Document Index production package

## User setup

1. Extract the complete ZIP to a normal Windows folder.
2. Double-click `Create_NMDC_Document_Index.vbs` once.
3. Open the created `NMDC_Document_Index.xlsm` in Microsoft Excel desktop.
4. Enable macros when Excel asks.
5. Press **Select Data Folder** and choose the folder containing `METHODS` and `TECH`.
6. Press **Update Changed Files** for normal work, or **Full Rescan / Rebuild All** when a complete rebuild is intentionally required.
7. Review **Pending Update** and **Review Flags** before approving.

The one-time setup uses Microsoft Excel to attach the audited VBA modules to the
workbook. If Excel blocks that attachment, the setup displays the exact Trust
Center setting needed and stops without changing the source workbook.

Setup verifies the packaged engine, configuration files and required core VBA
modules before opening Excel. During the same one-time setup, the owner UX also
loads the packaged Live Filter and Custom Fields modules and installs their
workbook event handlers. The generated workbook therefore remains self-contained
for normal later use even if Trust Center access to the VBA project object model
is turned off again.

The generated workbook stores the required engine/runtime/configuration paths so
folders containing spaces work reliably. Runtime/cache data is kept under the
local application-data area when available. If the extracted package is moved or
renamed, rerun setup from its new location so package paths are refreshed.

## Live Filter

The main data/review sheets include a row-3 Live Filter. It is implemented with
normal worksheet cells rather than an ActiveX textbox for better reliability on
corporate PCs.

- Type normal text for case-insensitive partial matching.
- Use spaces or `+` between terms for AND matching.
- Prefix a term with `-` to exclude rows containing it.
- Put the complete search in double quotes for exact case-sensitive matching.
- Press `Ctrl+Shift+F` to choose one table column.
- Use `ALL COLUMNS` to search the complete row.
- Clear the search cell to remove only the Live Filter condition.

Live Filter is provided on Master Documents, Revisions, Transactions, Pending
Update, Review Flags, User Decisions, Update History and Error Log.

## Custom Fields & Keywords

Use **Custom Fields & Keywords** when you want an extra derived column in Master
Documents without changing the canonical NMDC extraction fields.

Example:

1. Add a field named `Vessel Names`.
2. Choose the Master Documents columns to search, such as
   `Document Title;Source File`, or use `ALL TEXT`.
3. Add keyword mappings such as `SAFEEN 3000 -> SAFEEN-3000`.
4. Choose a match type: `CONTAINS`, `ALL TERMS`, `EXACT`, or advanced
   `WILDCARD` (`?` fixed width / `*` variable width).
5. Choose `FIRST` or `ALL UNIQUE` output behavior.
6. Press **Apply Custom Fields**.

Custom-field definitions and keyword mappings are backed up in the local runtime
area. Core fields such as Project No., Document No., Discipline, Category,
Subcategory, Source File and internal technical keys are protected from overwrite.
A core table refresh may temporarily rebuild Master Documents from canonical CSV
output; enabled custom fields are restored automatically on the next Master
Documents interaction.

## Safety

- Source registers are read-only.
- Every update is staged first.
- Hold and Reject do not change the approved index.
- Approve binds to the exact staged run shown in Excel.
- Parser, hash, and duplicate-record-key conflicts cannot be overridden.
- User custom fields are an enrichment layer and do not rewrite canonical source extraction fields.
- Existing production workbooks are backed up before replacement.

The package is not digitally signed. Corporate macro and executable policy may
require IT approval before use.

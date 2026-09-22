# NMDC Document Index — User Product Requirements

**Revision:** 2.0 (Post-Stage 5 Stabilization & Real Excel Verification)  
**Collaboration ID:** `NMDC-DOC-INDEX-001`  
**Role:** Implementer / Test Engineer / Release Coordinator  

This file captures the owner-visible product requirements that govern the NMDC Document Index application, user interface, and backend engine.

---

## 1. Core Product Boundary

- **Excel is the primary user interface:** The end-user operates the system entirely within Microsoft Excel (`NMDC_Document_Index.xlsm`).
- **No Developer Tooling Required:** The owner does not need Python, PowerShell, Command Prompt, Git, or direct runtime-file editing for normal daily operations.
- **Read-Only Source Policy:** Source `DATA/` workbooks (`METHODS/` and `TECH/`) are historical records and are **strictly read-only**. The index pipeline never unmerges, saves, edits, normalizes, or repairs files in `DATA/`.
- **Staged Approval Model:** No scan or update may silently replace or modify the approved index. All extractions are initially staged into `Pending Update` for explicit operator review and approval.
- **Merge & Acceptance Authority:** The repository owner alone holds final acceptance and merge authority.

---

## 2. Reliability & Source Interpretation

- **Lossless Extraction:** Merged-cell hierarchies (`Project -> Document -> Revision -> Event`) must be preserved without uncontrolled fill-down leakage.
- **Valid Baseline Corpus:** Full Rescan on the known valid `DATA/` corpus must complete with zero actionable extraction Review Flags and zero Error Log entries.
- **Transient Error Handling:** File-sharing or lock errors on synchronized OneDrive files must be retried automatically before declaring a source unavailable.
- **Safe Fallbacks:** If a source file is temporarily inaccessible, the approved index remains protected and the source is not marked as deleted.
- **Empty Result Integrity:** Intentionally empty registers or empty search results must keep their named Excel tables structurally valid without crashing VBA.

---

## 3. Excel Workbook Architecture

The production workbook (`NMDC_Document_Index.xlsm`) consists of exactly **15 worksheets** and **15 unique named Excel tables (`ListObjects`)**:

1. **`Home`**: Executive dashboard, key metrics cards, and single-click action buttons.
2. **`Master Documents`** (`MasterDocuments` table): Unique document-level register.
3. **`Revisions`** (`RevisionRegister` table): Revision-level history with latest-revision flags.
4. **`Transactions`** (`EventRegister` table): Complete event/submission/transmittal transaction log.
5. **`Pending Update`** (`PendingUpdate` table): Staging area for new and changed records awaiting approval. Includes in-cell source-selection checkboxes.
6. **`Review Flags`** (`ReviewFlags` table): Flags documents requiring classification review, ambiguous numbers, or syntax warnings.
7. **`User Decisions`** (`UserDecisionLog` table): Persistent audit log of operator review decisions.
8. **`Configuration`** (`Configuration` table): Path configurations (`Data Folder`, runtime directories, engine executable path).
9. **`Rules & Mappings`** (`ClassificationRules` table): Non-coder editable document classification taxonomy.
10. **`Custom Fields`** (`CustomFields` table): Formula-driven custom columns and user-defined metadata.
11. **`Update History`** (`UpdateHistory` table): Historical record of all staged, approved, held, or rejected runs.
12. **`Error Log`** (`ErrorLog` table): Plain-English error reporting with recommended actions.
13. **`System Data`** (`BaselineCounts` & `SourceInventory` tables): Internal audit counts and source inventory.
14. **`Source Selection`** (`SourceSelection` table): Backup source selection registry.

---

## 4. Live Filter & User Experience Requirements

- **Dynamic Live Filter — owner REV03 reference architecture:**
  - A real worksheet ActiveX textbox named `TxtBox_Search` is the typing surface.
  - `Cls_LiveFilter_Listener` handles `MouseDown`, `Change`, and `KeyDown`; filtering updates from the textbox `Change` event on every keystroke.
  - `Ctrl+Shift+F` is used only to choose a target **table header** and store that target in `LiveFilter_Anchor`; it must not capture normal typing.
  - Placeholder/help text follows the reference tool: `Search <Column>... (+AND / -EXCLUDE) (Ctrl+Shift+F: Change Column)`.
  - Spaces or `+` mean required/AND terms; `-word` excludes rows containing that term; surrounding quotes invoke the reference exact-match behavior.
  - A visible **RESET SEARCH** button clears the table filter, clears the textbox, restores the placeholder, and returns focus to the search box.
  - The same control pattern is installed on Master Documents, Revisions, Transactions, Pending Update, Review Flags, User Decisions, Update History, and Error Log.
- **Visual Presentation Standards:**
  - Font: Clean Aptos 10 typography across all data cells.
  - Automatic font coloring with no harsh or arbitrary cell background fills.
  - Distinct headers and field-appropriate column alignments.
  - Row heights and column widths formatted to prevent visual clutter and maintain high scroll performance.
- **64-Bit & Enterprise Path Compatibility:**
  - VBA declarations use `PtrSafe` (e.g. `Declare PtrSafe Sub Sleep Lib "kernel32"`).
  - Commercial OneDrive paths (`https://...my.sharepoint.com/...`) are automatically translated to local filesystem paths.

---

## 5. Staged Review & Admin Controls

- **Source Scope Controls:** Native Microsoft 365 in-cell checkboxes inside `Pending Update` allow including or excluding entire source workbooks in a single **Save Selection & Restage** operation without repeatedly re-scanning. The owner environment supports native in-cell checkboxes; the production UX must not silently downgrade these controls to visible TRUE/FALSE text.
- **Non-Coder Rules Editor:** Plain-text match types (`CONTAINS`, `EXACT`, `STARTS_WITH`, `ENDS_WITH`). No complex regular expressions required for routine rules.
- **Actionable parser/mapping exception workflow:** `Review Flags` uses a native Microsoft 365 **Select?** checkbox column and separate **Source File** / **Worksheet Name** columns so the exact affected worksheet is explicit. When selected rows are reported for `NEEDS PARSER/MAPPING FIX`, the workbook must create `PARSER_FIX_REPORTS\0001`, `0002`, ... beside the workbook package. Each request folder is self-contained and must not overwrite or mix with an older request. It must explain that the installed executable cannot safely rewrite its own parser code. After a corrected parser/configuration is installed, **Retry After Fix** must reprocess the source through Full Rescan without modifying source DATA or approving the staged proposal automatically.
- **Safe Recovery Actions:**
  - **Reset All Records:** Deletes runtime cache and index records without touching source `DATA/` or configuration files.
  - **Undo Last Approval:** Restores the previous approved index version from audit history.

---

## 6. Acceptance & Release Criteria

- **Real Excel Simulation Matrix:** Must pass all automated gates in the 23-case real Microsoft Excel matrix (`test_real_excel_simulation.py`) with zero failures.
- **Production Package:** Verified standalone installation package containing `Create_NMDC_Document_Index.vbs`, pre-compiled engine, VBA modules, and SHA-256 verified base workbook chunks.
- **Release Status:** PR #8 remains in `CHANGES` until the repository owner completes final interactive acceptance and authorizes merge.

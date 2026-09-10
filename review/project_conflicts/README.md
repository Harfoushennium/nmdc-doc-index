# Project conflict validation package

This folder is for human validation only.

- The Excel files under `files/` are exact copies of the original source workbooks.
- The originals remain in `DATA/` and are not modified, renamed, or moved.
- `conflict_manifest.csv` records the project-number evidence that conflicts with the filename/path project.
- No project ownership has been automatically corrected. Every item remains `USER_VALIDATION_REQUIRED` until the owner confirms it.

Project identity evidence should be considered separately:
1. source filename/path project,
2. explicit project number inside the workbook,
3. document-number project prefix.

If the internal project number and document-number prefix agree against the filename, the source file may simply be misnamed; this package exists so that can be confirmed before the master index is produced.

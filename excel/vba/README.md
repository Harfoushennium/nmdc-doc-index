# Excel VBA Source

These modules are the auditable source for the future production `NMDC_Document_Index.xlsm` workbook.

## Modules

- `modNMDC_Engine.bas` — configuration lookup, silent engine launcher and Excel error logging.
- `modNMDC_Refresh.bas` — imports deterministic CSV exchange files into workbook sheets and refreshes the Home dashboard.
- `modNMDC_Actions.bas` — button macros for update, full rebuild, approve/hold/reject, folder selection, flags and support requests.
- `modNMDC_Startup.bas` — refreshes the dashboard when the workbook opens.

## Production packaging rule

The text modules in this folder are the source of truth. The production `.xlsm` is a packaged artifact that must be created/validated on Windows with Microsoft Excel and the approved VBA modules attached to the corresponding workbook buttons.

Do not claim the production workbook is complete merely because these source modules exist.

## User prerequisites target

Normal user prerequisites are intentionally minimal:

- Windows;
- Microsoft Excel desktop;
- macros/VBA permitted by corporate policy;
- permission to run the packaged `engine\nmdc_index_engine.exe` beside the workbook;
- access to the selected local/OneDrive/shared/network source folder.

The user should not need Python, Git, GitHub Desktop, PowerShell, Command Prompt or a code editor.

## Expected package layout

```text
NMDC_Document_Index.xlsm
engine/
  nmdc_index_engine.exe
runtime/
  ...generated local state...
```

`runtime/` is local generated state and is excluded from Git.

## Security / support behavior

- The engine is launched hidden and Excel waits for its exit result.
- A non-zero exit result does not silently update approved data.
- Friendly errors are written to the workbook Error Log.
- The user can create a support request for GPT/developer follow-up.
- Technical non-overridable conflicts remain blocked until corrected.

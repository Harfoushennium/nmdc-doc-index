# NMDC Document Index — User Product Requirements

**Written by:** ChatGPT  
**Role:** Implementer / Release Coordinator

This file captures the owner-visible product requirements that must remain true as the Excel interface and extraction engine evolve.

## Core product boundary

- Excel is the normal user interface.
- The owner must not need Python, PowerShell, Command Prompt, GitHub, or direct runtime-file editing for normal use.
- Source DATA workbooks are read-only inputs and must never be modified by the index.
- No staged update may silently replace the approved index.
- The owner remains the final approval and merge authority.

## Reliability requirements

- The current known valid DATA corpus must complete a Full Rescan with no actionable extraction Review Flags and no extraction/refresh Error Log entries.
- Known valid layout variations must be handled automatically.
- Intentionally empty register sheets must not be reported as extraction failures.
- Routine lifecycle information such as new/changed sources belongs in Pending Update/history, not Review Flags.
- Genuine future faults (locked/corrupt sources, true conflicts, ambiguous duplicates) must remain visible and fail safely.
- Transient Excel/OneDrive file-sharing errors must be retried automatically before a source is declared unavailable.
- If a source remains unavailable after retries, the approved index must be protected and the source must not be interpreted as removed.

## Excel workbook requirements

- Every required data sheet keeps a valid uniquely named Excel table even when empty.
- Refresh populates tables by exact ListObject/table name and never destroys titles/navigation areas.
- Active filters/sorts are reset before refreshed rows are replaced so hidden/sorted row state cannot mix data after updates.
- Source File and Document Link hyperlinks are row-specific and must open the correct target.
- No temporary worksheets may be left behind during refresh.
- Empty result tables must remain structurally valid.
- Dates must be real Excel dates, numeric counts real numbers, and document/revision identifiers preserved as text.

## Table presentation requirements

- Tables must be professionally formatted without manual cleanup.
- Body cells use Aptos 10, automatic font color and no forced fill color.
- Headers remain visually distinct and readable.
- Long text wraps only where helpful.
- Alignment is field-appropriate and consistent.
- Row heights and column widths must be suitable for the actual data while preserving performance on large tables.
- User-editable fields are clearly identifiable through notes, appropriate controls and guidance rather than relying on arbitrary body fill colors.
- Table columns are ordered according to the review workflow: business identifiers and decisions first, technical/audit keys later.

## User-input control requirements

- Use Excel checkboxes for direct binary owner choices when practical.
- Source workbook inclusion/exclusion is checkbox-first: checked means include in index scope; unchecked means intentionally exclude.
- Checkbox changes are collected first and applied in one **Save Selection & Restage** action; checking/unchecking must not trigger a scan per click.
- Provide **Check All** and **Uncheck All** controls for source selection.
- Excluding a source never edits or deletes the source workbook and never changes the approved index until a later explicit approval.
- Use dropdown menus when one field has three or more mutually exclusive choices, for example Review Flags decisions, resolution status, rule scope and match type.
- Keep explanatory comments/notes as free text.
- Do not add large numbers of unnecessary checkbox objects to large extraction tables because workbook responsiveness is a product requirement.

## Rules & Mappings requirements

- Rules & Mappings must be usable by a non-coder.
- The normal view must emphasize plain-language concepts: where to look, what to match, and what classification to assign.
- Normal matching uses plain text: `CONTAINS`, `EXACT`, `STARTS_WITH`, or `ENDS_WITH`.
- Shipped/default rules must not require REGEX.
- Normal categorical inputs use dropdown menus.
- Add Simple Rule should generate a safe Rule ID / Priority and sensible defaults.
- `CONTAINS` is the preferred normal-user match method.
- Invalid or duplicate rules must be blocked before staging rather than silently accepted.

## Review workflow requirements

- Pending Update is review-only and contains no normal row-level user input.
- Source-level scope decisions are made **inside Pending Update** in a same-sheet source-selection section, using modern Microsoft 365 in-cell checkboxes and an optional Owner Note. There is no separate owner-facing Source Selection worksheet.
- Review Flags is only for genuine actionable anomalies.
- Every table column has guidance explaining what the field means and whether it is system output, user input, or audit data.
- Review decision choices must explain their effect; saving a decision records the review and does not silently rewrite the source workbook.
- Reset All Records and Undo Last Approval must protect source data/configuration/audit history.

## UX/performance requirements

- No black command window during normal engine execution.
- No merged-cell warning on workbook open.
- No Excel slow-workbook warning caused by formatting the unused worksheet tail.
- Full Rescan and refresh show meaningful progress/activity feedback.
- The workbook should remain responsive during long engine operations.
- Normal **Update Changed Files** should avoid rereading/reprocessing unchanged sources.
- Profiling/extraction should use a persistent local cache where possible so OneDrive source workbooks are not repeatedly opened during normal scans.
- Runtime/state/cache files should be stored outside the synchronized package folder when `%LOCALAPPDATA%` is available.

## Release rule

PR #8 remains CHANGES and unmerged until the owner completes the final Excel acceptance sequence on a fresh production package and explicitly authorizes merge.

# NMDC Document Index V2 — RC1 candidate

Extract the entire portable ZIP and double-click Launch.exe. The app uses an existing browser and processes records locally. Python, pip, Node, Git, Office, SDKs and administrator access are not application prerequisites. Excel is required only to inspect exported XLSX files during physical acceptance. No operational DATA or user state is included.

The default workspace is %LOCALAPPDATA%/NMDCIndexV2. For acceptance, use Run-UAT.ps1 and UAT-GUIDE.md to create an isolated workspace, verify the manifest and record the identical build on work and home PCs. This RC is not production approved. Keep V1 and Candidate 09 available.

Select a read-only DATA folder in Sources & jobs, scan, review before/after differences, save exact file choices and flag decisions, then explicitly approve. Deferred sources retain their previously approved records. Changed or unavailable inputs, dirty drafts and stale concurrent reviews prevent unsafe approval. Correct blocking parser flags and retry; marking a flag resolved cannot override a parser failure.

The approved Events/Documents/Revisions views support search, exact/AND/exclusion terms, filters, groups, paging, column order, saved views and themes. Rules and mappings apply through validated editors with configuration history. Custom fields are derived Document columns; unsupported legacy wildcard constructs are explicitly rejected. Dashboard totals use the approved views.

Export defaults to full event history. Each XLSX has one INDEX sheet and one native table, including empty results. Both hyperlink columns and literal original targets are mandatory. Native hyperlink eligibility verifies syntax, not target availability. The preview includes unsupported-target exceptions with source provenance; inspect actual links on both PCs. Excel's 65,530 native-link limit requires explicit partition consent for the full corpus. No unverified formula-based alternative is shipped. CSV neutralizes formula prefixes; JSON retains literal values. Exports require a preview bound to the exact snapshot, configuration, format, filters and columns.

Settings provides read-only legacy preview, confirmed side-by-side import into a new workspace, verified backup/isolated restore, undo and reset. Unsupported or inaccessible legacy content blocks import instead of silently defaulting. Backups retain the newest ten verified bundles; preserve any required archival copies outside that managed folder before further operations. Reset retains settings and audit history and creates a verified backup. Use the returned workspace path with Launch.exe --workspace "path" to inspect an import/restore; the active workspace and legacy baseline are not overwritten.

Use Exit application to stop the local service. A second writer to the same workspace is refused. Interrupted jobs require rescan. Parser reports belong outside DATA, its project root and the app workspace. If cloud-backed inputs are unavailable, make them locally available and retry. Do not bypass organization policy or delete live state to force a scan.

Production approval remains blocked until the complete release matrix, independent QA, owner-workspace migration, physical Excel checks, both-PC measurements and explicit legacy-ticket dispositions are reconciled. No production merge or cutover is authorized by this candidate.

# NMDC V2 — Standalone architecture and implementation plan

Status: planning candidate, 2026-11-28. Collaboration NMDC-DOC-INDEX-001.
Baseline: PR #8 candidate commit 88711792308bbeaa924ed3f2a5b9499a52e1c91b. Do not merge PR #8 or change its branch as part of this work. The planning branch is not a release.
Canonical decision map: https://github.com/Harfoushennium/nmdc-doc-index/issues/22

## Confirmed requirements

- W01: local browser application, existing Python engine, selective optimization supported by paired benchmarks.
- W02: preserve existing file-level approval, migrate all settings and compatible data/history, retain and enhance parser reporting.
- W03: macro-free one INDEX worksheet with one native table; complete event history default, optional document/revision/filtered scopes, saved configurable column profiles with core fields, preserve both clickable targets where verified and always original paths; preflight limits, explicit partition consent; optional CSV/JSON.
- W04: sidebar, unified document/revision/event module, advanced filters/search/grouping/saved views, independent jobs/progress/cancel/retry, expandable file review and integrated flags, dashboard, persistent theme/density preferences.
- W05.3 owner clarification: standalone application; no separately installed Python or prerequisite tools/apps. Normal-user offline execution; all dependencies and UI assets included. Existing Windows browser is used, no separate browser/WebView2 installation required. No requirement to use a single physical binary. See issue 28 comment 7228810838.
- Remaining W05.1/2/4/5/6/7 are proposed, not approved by the owner's instruction to proceed with packaging. Storage ticket remains blocked by the real-corpus benchmark.

## Evidence and selected build candidate

At pinned baseline .github/workflows/production-package.yml, windows-engine-package uses Python 3.11, installs PyInstaller on the build runner, builds nmdc_index_engine.exe via --onefile and --add-data "config;config", smoke-tests export-excel, and packages legacy VBA/workbook resources. This is evidence of an existing build route, not proof that the V2 web app is packaged or accepted.
Reuse the Windows packaging approach, not the Excel installer or VBA modules. Evaluate PyInstaller onedir first (predictable startup and asset layout); compare onefile only if measured startup and antivirus behavior warrant. Pin dependency versions and hashes and use a pinned Windows build image, build on Windows, bundle Python, SQLite, chosen XLSX library, web assets and notices. Build tools remain CI-only. Never include operational DATA, cache, user configuration, credentials or private CAD code. Use synthetic fixtures in public CI.
Reference CAD architecture: 54596cd8052f87b0e78e54553c4f24abd3015eb5, local browser/host split, authenticated loopback, explicit job lifecycle and atomic verified releases.

## Runtime boundary

Launcher starts authenticated loopback HTTP service on an ephemeral port, waits for health, opens existing supported browser, records startup errors, and provides explicit exit. Reject non-loopback bind, unexpected Host/Origin and unauthorized file operations; use scoped folder capabilities, no arbitrary path endpoint, no external uploads. No external CDN, analytics or mandatory update check. Separate long-running worker with bounded events, cancel and retry, reconnect to active job after tab reload; application exit must resolve active jobs safely. Serve only packaged assets. Do not start a second writer on double launch. Native folder selection may use bundled system-dialog adapter; test normal-user access to spaces, Unicode, UNC and locally synced OneDrive. No assumption that all corporate machines permit executables: report policy blockers instead of bypassing them.

## Storage recommendation (pending W05)

Keep immutable approved engine snapshots/manifests/caches and atomic current pointer authoritative. Derived SQLite tied to snapshot/fingerprint/schema; rebuild when absent or stale, never dual authority. Keep active DB/cache local, not live-synced OneDrive. Settings, drafts and histories are durable; guard saved choices by run and baseline. Separate configurable reports/backups/exports outside source scan; original DATA strictly read-only. Preserve V1 and run migration dry-run with checksum, count, identity, configuration and history reconciliation; report incompatible entries. Recommend automatic backups before approval/migration/reset/upgrade, owner-configurable retention. No change to exact storage location or retention until confirmed.

## Implementation increments (not authorization to publish)

1. Packaging-first tracer: standalone launcher and local host with bundled static UI, synthetic fixture, query and one-table XLSX export; package manifest and clean Windows smoke. Do not ship requiring pip or interpreter installation.
2. Engine adapter: expose existing read-only discovery, incremental extraction, progress/cancel and proposal semantics without Excel exchange on hot path; validate fixture parity and preserve source hashes.
3. Query/data boundary: version-bound derived SQLite candidate, paged API, search grammar parity, dashboard; profile real workload before optimization.
4. UI and review: all approved W04 modules, durable choices, flags, reports, settings/rules/custom fields and undo. Preserve accepted/deferred new/changed/removed semantics, no silent approval.
5. Export: approved snapshot consistency, configured profiles and document/revision/event/filtered scopes, verified one-table Excel, links and limits, CSV/JSON, safe literal strings, atomic publication.
6. Migration: isolated dry-run, compatibility register, audit reconciliation, backup/restore, side-by-side acceptance.
7. Release: signed or hash-verified package and offline guide, clean normal-user Windows acceptance and paired benchmarks on both PCs. Release Latest/History staging and atomic promotion only after owner acceptance. Existing issues and PR remain open until replacement verified or explicitly withdrawn.

## Standalone acceptance matrix

All below must be measured, not inferred from synthetic benchmarks. Results today: NOT RUN for V2 Windows package.
- P01: fresh supported Windows 10/11 x64 normal user, no Python/pip, Node, Git, Office, separate WebView2 or SDK installed: unpack and launch using existing supported browser. No admin, network or prerequisite prompt.
- P02: offline launch, scan synthetic input, search, review/approve, export XLSX/CSV/JSON, exit/restart; no external network requests. Validate workbook structurally without Office; test actual link interaction separately on an Office-equipped acceptance PC.
- P03: launch from paths with spaces and Unicode; read local and synced source folders; verify source tree hashes before/after; source temporary lock must not become deletion.
- P04: repeated launch and browser refresh, active job reconnect, cancel, retry, controlled exit, no orphan process or duplicate writer.
- P05: reject wrong Host/Origin/token, out-of-scope paths and traversal; no network-accessible bind or accidental exposure.
- P06: package manifest covers every shipped file, hash verification, no bundled user data/secrets, license notices and pinned toolchain evidence.
- P07: versioned package upgrade and rollback; preserve user settings/history and report location; failed publication leaves old release intact.
- P08: actual approved dataset parity: records, revisions, event keys, dates, links, classifications, source metadata, file choices, flags and audit. No duplicates or unexplained losses.
- P09: paired traces on both PCs with identical authorized corpus; full scan, unchanged refresh, one changed workbook, approval, query, export, memory and disk. Proposed targets: startup <=5s, search/first page p95 <=0.5s, unchanged refresh <=3s excluding sync, 40,494-row export <=10s; full-scan target after baseline.
- P10: independent UI and engineering QA, all applicable legacy tickets mapped and accepted before cutover.

## Current blockers

No Windows build or clean-machine runtime has been executed for V2. The real corpus and both PCs are not accessible in the current execution session; synthetic evidence is not paired V1/V2 performance proof. Browser UI interaction from earlier prototype was unverified. The current architecture direction and standalone requirement are approved; other W05 choices are conditional proposals. This branch is planning documentation only, not a production package or merge request.

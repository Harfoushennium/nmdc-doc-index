# NMDC V2 — Final release and approval plan (RP-NMDC-V2-001 v1.0)

Status: proposed execution/acceptance contract, 2026-11-28. Owner approval of release architecture and criteria is pending. Governing project: #22; delivery PR #29. This supersedes the implementation sequence (not the requirements) of NMDC-V2-standalone-implementation-plan.md. No production cutover, merge, or source edits authorized by this plan.

## 1. Baseline and finite scope

Candidate 06: e4662243d1c13db8a6cfa42247349e1ebaee1e0c; CI run 46098642285, all three jobs success. 286 inherited tests on each of Linux and Windows; 15 query parity cases; standalone restricted-PATH and Edge workflow; 56 repository workbooks, 40,494 unique events, zero review/blocking flags, exact serial/parallel and incremental stage parity, unchanged source hashes. Packaged full job 11.736 s; verified unchanged median 0.276 s; query median 14.064 ms (includes cold model builds, not p95). Artifact 15243828172, SHA256 outer artifact ZIP 257ccc7ce73bdc1542e4c14f01ef0929a5fdffdbb878f20a4f542f174849b934. No claim of clean owner PC, complete legacy migration, actual Excel link clicks or production approval.

The first release is feature completion, migration, reliability, usability and independent acceptance, not open-ended optimization. Retain W01-W04 and mandatory W05.3 requirements. Recommend freezing Python + bounded workers + verified unchanged gate + immutable existing authoritative approved records + disposable snapshot-bound SQLite query model + direct XLSX. Do not merge the canonical SQLite prototype or C++ research into first release unless owner specifically chooses it and complete migration/semantic parity is proved. PRs #30 and #31 research are closed unmerged. One authoritative store, no dual writes. Existing PR #8 and legacy package remain rollback options. No operational data, private credentials, or raw diagnostic records added to public Git; use local owner evidence and sanitized aggregates. User data is read-only at source. No prerequisite installs or admin/network requirements for ordinary app operation. Existing supported Windows browser is permitted. Physical Excel is needed only for actual Excel-output acceptance, not app execution.

## 2. Work packages, dependencies and exit gates

| Package | Scope | Depends | Exit gate |
|---|---|---|---|
| RP0 | Freeze requirements and define dispositions; inspect legacy configuration schemas, state formats, existing migration/manual instructions; inventory all V1 functions; define precise acceptance cases and exclusions | Candidate 06 | One traceable contract, each legacy requirement mapped, no ambiguous destructive behavior |
| RP1 | Complete application workflow and owner controls | RP0 | All W02-W04 functionality demonstrated in browser and API; tests F01–F14 |
| RP2 | Migration, durability, reset and recovery | RP0; integrate RP1 | Tests M01–M08, no silent loss, isolated dry-run and rollback |
| RP3 | Excel delivery and performance closure | RP0; integrate RP1/RP2 | Tests E01–E07 and B01–B05; owner-approved thresholds or documented exceptions |
| RP4 | Security, standalone package, regression and independent QA | RP1–RP3 | Tests P01–P07; all CI green for pinned RC; independent findings register |
| RP5 | Two-PC acceptance and ticket disposition | RP4 | Tests U01–U05 on work/home PCs, with exact build/data; all applicable issues closed or expressly superseded, no unapproved waivers |
| RP6 | Approval dossier and immutable release candidate | RP5 | Single owner approval package and decision; production publication only after explicit approval |

RP1, RP2 and RP3 may develop concurrently on isolated work branches; integration must be serial with explicit baseline reconciliation. Keep Candidate 06 usable. No arbitrary new feature scope during RC hardening. Every failing test becomes a tracked finding with reproduction, cause, patch and rerun; do not restart the planning cycle. No merge to master until owner decision.

### RP1: finish functional product

- Dashboard: reconcile counts by document/revision/event/source, processing status, meaningful charts and quick actions; persisted preferences.
- Source inventory: native folder selection, permanent include/exclude with reason, duplicate newest-modified policy, scope/config invalidation, read-only source and unavailable/locked-file handling. No missing source interpreted as intentional removal without explicit review.
- Job lifecycle: full and incremental scan, clear stages/progress, cooperative cancel with no premature publication, retry, browser reload/reconnect, explicit exit, persisted interrupted job/audit status, bounded waits and actionable error messages. No orphan worker or second writer.
- File Review: expand each affected workbook into actual before/after records and changed fields, worksheet/row/source evidence, accept/defer individually, saved vs draft status and current run identity; show prior-run drafts without replaying them. Accept changed/new/removed only as selected; deferred changed or removed retain approved, deferred new absent. No autoapproval.
- Review Flags: selectable findings, exact source file and worksheet, note/decision persistence and audit, resolve/hold/reject paths consistent with existing policy, non-overridable conflicts blocked. Report numbered parser-fix packages in a separate configurable folder outside DATA/project root, include source+worksheet+row/cell, expected correction, diagnostics and retry-after-fix. Reports never themselves resolve or approve.
- Configuration: rules, category/subcategory dictionary, project overrides/aliases, source exclusions, custom fields (definition/type/enabled/default/mapping, existing behavior), owner choices and validation; edit/save/reopen, safe revert, rule changes require new stage. Reuse existing schemas, do not invent unsupported semantics.
- Index: Document/Revision/Event tabs, existing AND/exclude/exact syntax, selected-column visible control, reset, multi-value filters, stable sort, paging, grouping, visible column ordering, saved views, persistent export profiles, keyboard navigation/focus and responsive layout. No hidden source identity loss. Dashboard/index/export all use same pinned approved version.

### RP2: non-destructive migration and recovery

- Inventory owner V1 package/workspace and separately any settings held only in XLSM tables. Read-only dry-run with versioned adapters; inspect all relevant config/durable state directories and actual workbook table schemas. Do not run macros, overwrite legacy data, or infer missing values.
- Migrate all compatible canonical records, historical versions, manifest/cache references, source choices/notes, pending runs/drafts, flags/decisions, rule mappings/aliases/custom fields and logs. Preserve original IDs, paths and timestamps. If a setting exists only in workbook, explicitly parse or report unsupported. Map old paths to new only with owner-selected root and verified evidence.
- Compatibility report: source/destination inventory, SHA256, record and key counts, exact field/value comparison, source metadata, rules counts/values, flags/notes/history, incompatible or missing elements. No silent defaults. Unresolved critical incompatibility blocks acceptance.
- Fresh isolated V2 workspace, import only after preview; verify same approved snapshot/queries/export. Back up before migration, approval, reset, undo, upgrade per proposed policy, with configurable retention and sufficient disk check. Test restore into separate workspace and explicit rollback; no unverified archive extraction.
- Reset needs confirmation, creates verified backup, clears runtime approved/staged views only; preserve sources, settings, rules, custom fields and audit history; valid empty index/export; rescan returns expected records. Never delete legacy/source tree.
- Test atomic promotion, interrupted jobs, stale certificate/cache invalidation, same-metadata changed file, partial approved/deferred merge, duplicate/conflict rejection, interrupted restore, corrupt package/archive, disk-full/write failure and recovery. Verify no partial publication or approval-pointer changes on failure.

### RP3: export and performance

- One macro-free XLSX containing only INDEX sheet and one native named table (including zero-record valid output; define supported Excel empty-table representation, no fabricated record). Default complete event history; optional Document/Revision/filtered view; stable column order and profile persistence; preserve mandatory core/provenance fields and leading zeros. Actual dates/numeric values use appropriate cell types; untrusted formula-like text stays literal.
- Preserve document and source-workbook hyperlinks plus literal original paths; resolve relative targets from source workbook, support local/UNC/OneDrive paths and special characters. Use native links or safe verified alternative only; preflight all Excel limits and long/unavailable paths, show explicit exceptions; never silently omit links. Because two native links per row can exceed the 65,530 worksheet cap at 40,494 rows, verify single-file safe alternative and explicit partition option; owner decides when limitations require split. Validate actual Excel click behavior on both PCs. No hidden sheets, VBA, repair prompt, unexpected external links or formula injection.
- Preflight row/column/link limits, explicit confirmation before multi-file split, each file one sheet/table, manifest and full record/identity/order reconciliation; export pinned snapshot, atomic output, safe cancel/failure, no approved-state mutation. Optional CSV/JSON; document neutralization semantics.
- Reuse actual 56-workbook corpus; paired same data/config/build machine measurements at least 5 warm and 3 cold trials for startup/full/unchanged/changed source/approval/export, and >=30 representative query requests for p95. Include cache builds, result serialization, disk writes, peak RSS and disk usage. Include controlled new/changed/removed/deferred fixture and 100k-event synthetic scaling. Separate source sync/locks from compute and report. Do not claim isolated primitive timings as end-to-end.
- Proposed release ceilings on supported owner PCs for local available corpus: startup <=5s; first page/search p95 <=0.5s; verified unchanged <=3s median; export 40,494 rows <=10s median; full scan <=15s median (new proposed baseline); changed-source and approval benchmark with agreed ceilings after measured reference. Stretch optimization goals (not automatic release blockers): unchanged <=0.5s, full <=8s, approval <=0.5s, query <=10ms, export <=5s. If mandatory threshold missed, fix or seek explicit documented owner exception, never silently downgrade. Memory/disk report and compare; no unbounded growth across repeat runs. No new rewrite merely to chase stretch targets.

### RP4: release packaging/security and regression

- Pinned dependency versions and toolchain/build image, include runtime/UI/assets/config defaults/license notices, package manifest of SHA256 and sizes, verify extracted ZIP and exact release commit. No source DATA, user state, private artifacts or secrets shipped. Avoid nested ZIP if packaging permits; include offline quick start, troubleshooting, release notes and migration/rollback guide. Versioned release folders; previous verified candidate retained.
- Normal-user Windows 10/11 supported target, existing browser; test offline, clean PATH, no Python/pip/Node/Git/Office/SDK/WebView2 install; launch paths spaces/Unicode; no admin, external network, firewall exception or unwanted install. Restricted PATH is not clean-machine certification.
- Loopback only, random port/token, strict Host/Origin, session isolation, no arbitrary paths/traversal, no executable/config injection, no external asset requests, read-only authorized DATA, secure local state permissions, export/report/backup destination separation. Test wrong token/origin/host, malformed input, XSS/formula-like strings, recovery archive traversal/oversize, second launcher and blocked writes. No unsafe bypass of organization policy.
- Execute inherited 286-test baseline suites on Linux/Windows plus all new unit/integration/regression suites, packaged executable and Edge tests, exact real-corpus verification, functional browser assertions not just API invocation, read-model parity and cleanup, full change/defer/approval/undo/reset/restart flows. Add coverage for discovered failures, never remove failed case to pass.
- Independent QA by separate reviewer/test pass against pinned RC; reproduce audit of actual exported XLSX ZIP/XML and physical Excel acceptance in RP5. Findings register: check, exact location, expected/actual, evidence link, priority, proposed correction, status. PASS/FAIL/BLOCKED/NOT RUN. P0/P1 blockers cannot pass; P2/P3 exceptions require explicit owner acceptance. Regression rerun on same final hash after any material change.

### RP5: owner acceptance and ticket disposition

Use same immutable candidate on work and home PC, isolated DATA copy/legacy workspace backup, normal user. Actual environment differences recorded. Run UAT once per PC, capture timings/OS/Excel version/build hash, sanitized screenshots and evidence. If PC or permission not accessible, mark BLOCKED, do not simulate approval. Owner can use the supplied local UAT runner/checklist; assistant can analyze returned evidence. Legacy Excel-specific controls only close as verified fixes on V1 or expressly superseded by accepted V2 replacement, not because UI differs. Retain original ticket history and link replacement acceptance. No legacy issue closure before owner agrees disposition.

### RP6: approval handover

Deliver versioned portable ZIP, SHA256, manifest, exact commit, release notes, user guide, migration/recovery instructions, acceptance matrix with evidence and two-PC results, consolidated findings register, performance comparison, ticket closeout report, known limitations/accepted exceptions, explicit final approval choice. Ready for approval means all mandatory tests PASS on pinned artifact or formally accepted noncritical exception, no P0/P1, all compatible migration reconciled, rollback proved, actual UAT complete, applicable legacy dispositions recorded. Ready is distinct from approved. After explicit owner approval only: publish Latest/History and merge authorized PRs; leave V1 backup accessible and never delete historic data. No production changes during preparation.

## 3. Acceptance matrix (all gates must be recorded)

| ID | Verification | Required result |
|---|---|---|
| F01 | Folder selection, inventory, duplicate/newest and exclusions | Correct source list; read-only; scope change invalidates proposal |
| F02 | Full scan real corpus | 56/40,494 baseline when unchanged; exact semantic parity and hashes |
| F03 | Unchanged refresh | Same approved pointer, no proposal; verified inputs |
| F04 | New/changed/removed and locked/unavailable fixtures | Correct decisions, no silent deletion or lost data |
| F05 | Cancel/retry/reconnect/double launch/exit | No orphan/second writer; no partial publication |
| F06 | File expansion/diffs | Accurate old/new fields and worksheet/row evidence |
| F07 | Save drafts/reload/restart/supersede | Durable notes/checkboxes; old drafts retained, no stale replay |
| F08 | Mixed file approval/deferral | Accepted changes only; deferred removed/changed preserved; deferred new absent |
| F09 | Unsaved/stale/duplicate/conflict and hold/reject | No unauthorized promotion; actionable errors and audit |
| F10 | Flags and parser reports | Durable decisions; separate numbered reports; explicit worksheet; retry without autoapproval |
| F11 | Rules/aliases/exclusions/custom fields | Validate, save, reopen, restage, rollback; no silent data loss |
| F12 | Index queries and sorting | All scopes exact reference parity; typed input/focus/reset; grouping and filters |
| F13 | Saved views/columns/profiles/theme/density | Persist across restart; consistent filtered export |
| F14 | Dashboard | Counts reconcile with approved version; functional job/status actions |
| M01 | Legacy read-only inventory/dry-run | Full compatibility report, unknown formats flagged |
| M02 | Migration into new workspace | Exact keys, fields, metadata, settings, histories; no overwrite |
| M03 | Migrated pending/approved histories | Same choices/flags/notes/identity; no unauthorized replay |
| M04 | Backup policy/retention/restore | Verified bundle; preserve required areas; restore to isolated workspace |
| M05 | Undo with/without previous version | Exact prior state or clear refusal, audit preserved |
| M06 | Reset/empty/re-scan | Settings/history retained; valid empty UI/output; rescan restores |
| M07 | Crash/write failure/cache corruption | No partial promotion; recover prior version; invalidate stale cache |
| M08 | Upgrade and rollback | Version compatibility and preserved settings; reversible |
| E01 | Full single-table XLSX | One INDEX/one native table, exact rows/keys/types, no macros |
| E02 | Empty and filtered/document/revision exports | Valid table and scope, no fabricated rows |
| E03 | Native and alternate hyperlink behavior | Real document/source links work on actual Excel; literal original targets retained |
| E04 | Capacity/partitioning/long link | Explicit preflight/consent, no missing records/links silently |
| E05 | Formula/text/date/leading zeros | Literal untrusted text and correct native types |
| E06 | CSV/JSON and saved profiles | Expected schema/order/safety and literal JSON values |
| E07 | Export failure/snapshot consistency | No mixed versions or partial delivered file; state unchanged |
| B01 | Cold launch and first page/search | Approved ceilings and p95; no query semantic regression |
| B02 | Full/unchanged/changed-source scans | Paired trials, source hash/parity, approved ceilings |
| B03 | Approval/undo/export | Complete measured operations; not isolated primitive timings |
| B04 | 100k scale and repeated operations | Correctness, bounded memory/disk, no accumulation |
| B05 | Both PCs comparable traces | Data/config/build identical; sync overhead separately reported |
| P01 | Fresh normal-user offline Windows | Launch/scan/search/review/export/restart without installs |
| P02 | Paths Unicode/spaces/UNC/OneDrive | Read correctly or actionable safe failure |
| P03 | Authentication/loopback/input/path controls | Deny unauthorized access; no remote exposure |
| P04 | Package manifest/ZIP/licensing | Verified all files, no operational data/secrets |
| P05 | Full regression on final SHA | Inherited+new tests PASS Linux/Windows, actual packaged browser PASS |
| P06 | Independent QA | No P0/P1, evidence attached, noncritical exceptions explicit |
| P07 | Upgrade/recovery instructions | Offline guide validated on clean test workspace |
| U01 | Work PC end-to-end | Actual pinned package and local data; evidence |
| U02 | Home PC end-to-end | Same package/corpus/settings; evidence |
| U03 | Actual Excel output on both PCs | No repair prompt; table/date/links correct |
| U04 | Migration/backup/restore on owner workspace copy | Full reconciliation and reversible cutover readiness |
| U05 | Ticket/requirements traceability and approval dossier | Every ticket disposition justified; owner can approve or reject |

## 4. Open ticket disposition map

| Existing ticket | Release package / evidence | Disposition gate |
|---|---|---|
| #10 | RP1/RP5 F01–F09,U01–U03 | V2 replacement accepted; native V1 issue superseded only with owner agreement |
| #11 | RP1/RP2 F11,M01–M03,M06 | Rule/config integrity and migration; V1 issue disposition |
| #12 | RP2/RP3 M06,B02 | Reset/empty/rescan, preservation and timings |
| #13 | RP4/RP5 P01,P03,U01/U02 | V1 modal harness expressly superseded or original native verification |
| #14 | RP2 M05,M07 | Undo positive/negative, failure/timeout; V1 modal disposition |
| #15 | RP5 U01–U05 | Both-PC physical matrix; retain original 23-case traceability |
| #16 | RP1 F12,F13 | Visible search-column control, keyboard/focus and reset; legacy shortcut disposition |
| #17 | RP1 F07,F08,F10 | Working browser review/flag controls; old native checkbox disposition |
| #18 | RP1/RP2 F07–F10,M03 | Durable drafts, dirty approval block, historical access, failures |
| #19 | RP3 B02,B03,B05 | Actual repeat/changed/failed update timings, parity and user input protection |
| #20 | RP1/RP3 F12,B01,B04,B05 | Syntax/physical typing/reset/query performance |
| #21 | RP5 | All five legacy slices accepted or expressly superseded |
| #24 | RP3/RP5 B01–B05 | Measured performance report and thresholds disposition |
| #28 | RP0/RP2/RP5 M01–M08,U04 | Confirm storage freeze, full migration/recovery and rollback |
| #22 | RP6 | All applicable requirements and dispositions complete, approval decision |
| PR #29 | RP4–RP6 | Keep open/draft until pinned RC evidence and owner approval; merge only after authorization |

## 5. Decisions requested with plan approval

D1: Freeze hybrid architecture for V2.0; defer canonical SQLite and C++ research. Alternative: mandate new canonical store now, extending RP2/RP4 with complete dual-candidate migration and crash/parity testing, never dual authority.
D2: Proposed release ceilings vs stretch targets as above; approve changed-source/approval ceilings after measured comparable baseline. No silent waivers.
D3: Side-by-side read-only migration and backups before mutation; configurable local storage/report/export paths and retention; no destructive cutover before acceptance.
D4: Legacy tickets closed only after verified replacement and explicit supersession approval or original acceptance, not blanket "fixed."
D5: Mandatory two-PC UAT before final readiness; owner may expressly change target environments, never fabricate results.

Approval of this plan authorizes implementation/tests and evidence-based closeout; does not itself approve application or production cutover.

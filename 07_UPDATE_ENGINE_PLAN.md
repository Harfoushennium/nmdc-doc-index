# 07 — Incremental Update & Staged Approval Engine

## Owner goal

The normal user must work only inside the final Excel product. This PR builds the backend update foundation that the future Excel buttons will call.

## User-visible behavior to support

1. **Update Changed Files**
   - detect NEW / CHANGED / UNCHANGED / REMOVED source workbooks;
   - reprocess only NEW and CHANGED workbooks;
   - reuse cached canonical data for UNCHANGED workbooks;
   - never append blindly to the approved master dataset.

2. **Full Rescan / Rebuild All**
   - intentionally ignore extraction cache;
   - process all currently selected source workbooks again;
   - rebuild a fresh staged dataset.

3. **Review Pending Update**
   - compare the proposed staged dataset with the currently approved dataset;
   - show Added / Modified / Removed / Unchanged records;
   - show plain-English flags and conflicts.

4. **Approve Update / Hold / Reject**
   - the approved master dataset must not change until explicit approval;
   - Hold/Reject leaves the approved dataset untouched;
   - Approval promotes the staged dataset and source manifest atomically.

## Technical design

### Source manifest

Store deterministic metadata for every candidate source workbook:

- relative path;
- SHA-256 content hash;
- file size;
- modified time;
- source selection status;
- parser version;
- configuration fingerprint;
- last processed run id.

The content hash is authoritative for change detection; modified time and size are supporting evidence only.

### Canonical cache

Cache extracted canonical rows by source workbook. On an incremental run:

- NEW -> extract;
- CHANGED -> re-extract the entire workbook;
- UNCHANGED -> reuse cached rows;
- REMOVED -> exclude cached rows from the staged rebuild and report the removal for user approval.

### Staging

Each run produces a staged package containing:

- proposed full canonical records;
- proposed source manifest;
- change summary;
- user-facing flags;
- run log / run metadata.

### Approval

Approval copies the staged package into the approved state. Rejection records the decision but does not replace the approved state.

## Plain-English flag levels

- 🟢 OK — normal record/update.
- 🟡 REVIEW — unusual data that may still be usable.
- 🔴 CONFLICT — must not be accepted silently.

Initial update-related flag codes:

- `SOURCE_NEW`
- `SOURCE_CHANGED`
- `SOURCE_REMOVED`
- `SOURCE_HASH_ERROR`
- `CACHE_MISSING`
- `CACHE_STALE`
- `PARSER_VERSION_CHANGED`
- `CONFIG_CHANGED`
- `APPROVED_OVERRIDE_CONTRADICTED`
- `UNRECOGNIZED_LAYOUT`
- `PARSER_ERROR`

## Safety

- `DATA/` remains read-only.
- Approved data is immutable until explicit approval.
- No source workbook is modified or deleted.
- No GitHub merge is automatic.
- Unknown/conflicting conditions remain visible.

## Scope boundary

This PR builds and tests the update/staging engine. It does not build the final `.xlsm` interface, package the Windows executable, or implement the future project-folder/hyperlink scanner.

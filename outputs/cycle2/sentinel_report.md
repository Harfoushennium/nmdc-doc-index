# Cycle 2 Sentinel Extraction Report

- Sentinel cases: **5**
- Event records: **568**
- Distinct documents: **153**
- Distinct revisions: **271**
- Records with hyperlinks: **31**

## Case reconciliation

- `M2369_PROC` — INCLUDE; rows=63; events=119; docs=21; revisions=63; warnings=none
- `M2891_INCOMING` — INCLUDE; rows=17; events=17; docs=15; revisions=17; warnings=DOCUMENT_NO_FALLBACK_COMPANY
- `M2171_NONSTANDARD` — INCLUDE; rows=104; events=321; docs=31; revisions=103; warnings=none
- `T2820_PIPELINE` — INCLUDE; rows=88; events=111; docs=86; revisions=88; warnings=DATE_TEXT_PRESERVED
- `T3291_CLIENT` — EXCLUDED; rows=0; events=0; docs=0; revisions=0; warnings=none

## Safety

Cycle 2 is sentinel-only. `DATA/` is read-only. No full-source extraction and no final XLSX are produced.

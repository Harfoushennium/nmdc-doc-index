# Project-number conflict validation package

These are unchanged copies of the original Excel registers that triggered a project-identity mismatch during Cycle 3. They are provided only for owner examination. The originals under `DATA/` remain untouched and authoritative as source evidence.

## Files and observed conflicts

1. `2035 DOCUMENT REGISTER.xlsx`
   - Filename/path suggests project **2035**.
   - Worksheet `Specification` contains document-number evidence for **2136**.

2. `2035 Offshore Construction Engineering Register (Pipeline & Cables).xlsx`
   - Filename/path suggests project **2035**.
   - Worksheet `Specification` contains document-number evidence for **2136**.

3. `2705 -DOCUMENT REGISTER.xlsx`
   - Filename/path suggests project **2705**.
   - Internal worksheet header shows **NPCC PROJECT NO. 2824** on relevant sheets.
   - Example document evidence includes `2824-NN-0001`.

4. `2745-PP-GE-001-MDR Rev_2.xlsx`
   - Filename suggests project **2745**.
   - Internal register explicitly shows **CONTRACTOR Project No. 8405**.
   - Many document numbers begin `8405-...`, including anchor-pattern and DP records.

## Validation rule

Do not decide project ownership from the filename alone. For Cycle 3, filename/path is only one signal. Explicit internal project labels and consistent document-number patterns are stronger evidence. Until the owner validates these four registers, the conflicting worksheets must remain held for review and must not be silently reassigned.

## What the owner should check

Open each workbook and confirm which NMDC/NPCC project number should own the records. If a file was simply saved under the wrong filename, record the correct project number here or in the PR review. No original `DATA/` file needs to be renamed or edited for this validation.

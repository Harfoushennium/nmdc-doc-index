# Real-data Excel review package

This package is a safe review snapshot built from the Cycle 3 extraction outputs.
It is not an approved master index and it does not change `DATA/`.

## What is in the workbook

- Home dashboard with real baseline counts;
- the complete 7,498-document view, 10,663-revision view, and 39,344-event view;
- all 15 existing safe-review worksheet flags;
- the current classification rules and source-selection audit;
- update-history, error-log, help, and user-decision areas.

The delivered workbook contains the complete tables. An optional package folder
also carries the same exchange in CSV form for an independent audit trail.

## How to test it in Excel

1. Open `NMDC_Document_Index_Real_Data_Review.xlsx`.
2. Start on **Home** and click the coloured controls. They are native Excel links,
   not dead visual blocks, and take you to the relevant review page.
3. Open **Master Documents**, **Revisions**, and **Transactions / Events**. Use the
   table filters to inspect projects, document numbers, revisions, statuses, and
   source locations.
4. Open **Review Flags**. Confirm that the 15 unusual/empty worksheet layouts remain
   visible as `REVIEW` and are not silently guessed into the index.
5. Open **Rules & Mappings** and **System Data** to inspect the editable rule view and
   source-selection evidence.
6. Return to **Home** with the `← Home` link on any review sheet.

The buttons labelled Update, Full Rescan, Approve, Hold, Reject, and Select Data
Folder deliberately route to the review guidance in this snapshot. The macro-enabled
production workbook and packaged Windows engine are not attached yet, so this file
does not pretend to run those actions or change approved master data.

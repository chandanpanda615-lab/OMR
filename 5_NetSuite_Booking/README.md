# 5_NetSuite_Booking — folder structure

This folder is the working root for the HUL → NetSuite vendor-bill pipeline.
All scripts in `..\0_Scripts` self-locate this folder as `OUT_ROOT`, so the
names below are a **contract**: the files/folders marked LIVE must keep their
exact names and location or the scripts break.

## Layout

```
5_NetSuite_Booking\
├─ <Hub>\                     LIVE  one folder per fulfillment centre
│   ├─ PDFs\<date>\*.pdf      LIVE  downloaded invoices waiting to be booked
│   └─ Booked\*.pdf           LIVE  auto-filled by mark_booked.py after import
│                                   (created the first time that hub is marked)
│   Hubs: Byrathi, Chrompet, Gouribidanur, Hosur, Mysore Road, OMR,
│         Pondicherry, Soukya, Tumkur, Yeshwantpura
│
├─ Master_Entry.xlsx         LIVE  the day's entry sheet — you FILL the yellow
│                                   cells (Amount_5%, Amount_18%, IGST?), or let
│                                   fill_tax_from_pdfs.py fill them. Column D
│                                   ("Attached file") is a click-to-open S3 PDF link.
│
│   generate_master_csvs.py writes these 4 (overwritten each day):
├─ 1_ALL_Bills_Header.csv    LIVE  NetSuite import — bill headers
├─ 2_ALL_Bills_Expenses.csv  LIVE  NetSuite import — expense lines
├─ ALL_Tax_Verification.csv  LIVE  check the tax split before importing
├─ ALL_Invoices.zip          LIVE  every referenced PDF, for File Cabinet upload
│
├─ _booked.csv               LIVE  the ONE dedup ledger — every booked invoice.
│                                   Never delete. download + build scripts skip
│                                   anything already in it.
├─ STATUS.csv                LIVE  written by status.py (dashboard)
├─ MISSING_PDFS.xlsx         LIVE  the MISSING PDFs LIST (tab "Chase"): every invoice not booked yet +
│                                   status / remark / bad_file / days_waiting. Rows are only added
│                                   or updated; a row leaves only when the invoice is in _booked.csv.
├─ CDMS_Recovered.csv        LIVE  links CDMS_Tool\download_invoices.py found for the Missing PDFs list;
│                                   download + build read it, archive_batch.py archives it
├─ Gemini\                   LIVE  every Gemini answer, kept forever (never delete - it is paid for)
│   ├─ raw\<md5>.json               one file per PDF = exactly what Gemini read; reused, never re-paid
│   └─ Gemini_Results.xlsx          those answers vs what went into the import files (SAME / DIFFERENT)
└─ Archive\                  finished batches — nothing here is read by the scripts
    ├─ Combined_2026-09-17\
    ├─ Reimport_2026-09-18\
    ├─ OMR_2026-09-13\  OMR_Sep05-10_15bills\  Yeshwantpura_2026-09-15\
    ├─ CDMS_Sorted_retired_2026-09-22\   the old separate CDMS batch (lane merged into the daily flow)
    └─ old_per_hub\<Hub>\   old per-hub-flow outputs (Entry.xlsx, per-hub CSVs,
                             source CSVs, per-hub zips) — superseded by the master flow
```

## Daily workflow (one lane — CDMS recovery is part of it)

(`RUN_DAY.bat "path\to\GRN.csv"` in `0_Scripts\` = steps 1 + 2.)
1. `python download_hul_pdfs.py "path\to\GRN.csv"` → first re-checks the Missing PDFs list in CDMS
   (token refreshes itself - see `CDMS_Tool\HOW_TO_RUN.txt`), then downloads GRN + CDMS-found
   PDFs into `<Hub>\PDFs\<yyyy-mm-dd>\` and prints per hub: Downloaded now / Already on disk /
   Booked (skip) / No link (list). No-link rows go on the Missing PDFs list.
2. `python build_master_workbook.py "path\to\GRN.csv"` → ONE `Master_Entry.xlsx` (GRN + recovered),
   pre-filled from saved Gemini answers. Refuses to overwrite a `Master_Entry.xlsx` that has amounts.
3. In `Master_Entry.xlsx`, per row: type the amounts yourself, OR type `x` in `No Gemini`, OR type a
   `Remark` if you can't book it (→ Missing PDFs list as "wrong file"). Then optionally
   `python fill_tax_from_pdfs.py` - text PDFs (copyable text, e.g. Chrompet) are read on your PC first,
   free; only scans go to Gemini. Answers are saved in `Gemini\` first (never paid/read twice).
   Unsure rows are red REVIEW with the printed tax shown - fill those by hand. Save.
4. `python generate_master_csvs.py` → the 4 `ALL_*` import files. STOPS if a row's amounts don't match
   the tax printed on its PDF. Updates the Missing PDFs list and `Gemini\Gemini_Results.xlsx`.
5. In NetSuite: upload `ALL_Invoices.zip` to the File Cabinet **FIRST**, then
   Import CSV Records (header + expenses) with **RUN SERVER SUITESCRIPT ✔**.
6. `python mark_booked.py`
   → logs today's invoices in `_booked.csv`, MOVES their PDFs from `<Hub>\PDFs\` to
   `<Hub>\Booked\`, and updates the Missing PDFs list (booked leave, not-booked join).
7. `python archive_batch.py` → today's batch files (incl. `CDMS_Recovered.csv`) → `Archive\Master_<date>\`.

Full step table: `0_Scripts\WORKFLOW.md`.

## Rules of thumb
- Anything under `Archive\` is finished — safe to ignore, safe to delete.
- Never move or rename a LIVE file/folder; the scripts hard-code these paths.
- `_booked.csv` is the source of truth for "already done" — keep it forever.

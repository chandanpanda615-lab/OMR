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
│         PNCY, Pondicherry, Soukya, Tumkur, Yeshwantpura
│
├─ Master_Entry.xlsx         LIVE  the day's entry sheet — you FILL the yellow
│                                   cells (Amount_5%, Amount_18%, IGST?)
├─ Master_Links.xlsx         LIVE  same row order — click column E to open each PDF
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
├─ _pending_not_booked.csv   LIVE  written by status.py
├─ STATUS.csv                LIVE  written by status.py (dashboard)
├─ MISSING_PDFS.csv          LIVE  written by download_hul_pdfs.py — invoices with
│                                   a blank S3 link (rewritten every download run)
├─ Missing_PDF_Invoices_16-17Sep.csv   manual record of 24 no-link HUL SAMADHAN
│                                        invoices (Chromepet + OMR) held out of booking
│
├─ CDMS_Sorted\              past one-off CDMS batch (self-contained: own CSVs + zip + PDFs)
└─ Archive\                  finished batches — nothing here is read by the scripts
    ├─ Combined_2026-09-17\
    ├─ Reimport_2026-09-18\
    ├─ OMR_2026-09-13\  OMR_Sep05-10_15bills\  Yeshwantpura_2026-09-15\
    └─ old_per_hub\<Hub>\   old per-hub-flow outputs (Entry.xlsx, per-hub CSVs,
                             source CSVs, per-hub zips) — superseded by the master flow
```

## Daily workflow (master flow)

1. `python download_hul_pdfs.py "path\to\GRN.csv"`
   → PDFs land in `<Hub>\PDFs\<date>\`; blank-link rows go to `MISSING_PDFS.csv`.
2. `python build_master_workbook.py "path\to\GRN.csv"`
   → builds `Master_Entry.xlsx` + `Master_Links.xlsx`.
   ⚠ This OVERWRITES `Master_Entry.xlsx`. Run it BEFORE you fill amounts, never after.
3. Fill the yellow cells in `Master_Entry.xlsx` (use `Master_Links.xlsx` side-by-side
   to open each PDF), save.
4. `python generate_master_csvs.py`
   → writes the 4 `ALL_*` import files above.
5. In NetSuite: upload `ALL_Invoices.zip` to the File Cabinet **FIRST**, then
   Import CSV Records (header + expenses) with **RUN SERVER SUITESCRIPT ✔**.
6. `python mark_booked.py`
   → logs today's invoices in `_booked.csv` and MOVES their PDFs from
   `<Hub>\PDFs\` to `<Hub>\Booked\`, so `PDFs\` only ever shows unbooked work.

## Rules of thumb
- Anything under `Archive\` is finished — safe to ignore, safe to delete.
- Never move or rename a LIVE file/folder; the scripts hard-code these paths.
- `_booked.csv` is the source of truth for "already done" — keep it forever.

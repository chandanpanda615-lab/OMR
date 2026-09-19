# HUL → NetSuite Vendor Bill pipeline

Books HUL / HUL SAMADHAN purchase invoices into Oracle NetSuite via the 2-file CSV import.
Human enters only the 5% / 18% taxable amounts; everything else is automated.

> Quick self-serve steps: see **QUICKSTART.md**.
> Visual step-by-step chart of both flows: see **WORKFLOW.md**.

## Folders
- `0_Scripts/`            – these tools (one shared set for every hub)
- `5_NetSuite_Booking/`   – working area. **Master flow** writes here at the top level:
  - `Master_Entry.xlsx`   – ONE entry sheet, every hub (fill amounts) ← the master report; the
    **Attached file** cell of each row is a clickable link to its PDF (no separate links file)
  - `1_ALL_Bills_Header.csv`, `2_ALL_Bills_Expenses.csv` – combined import files
  - `ALL_Tax_Verification.csv`, `ALL_Invoices.zip` – tax check + PDFs bundled for File Cabinet
  - `_booked.csv`   – ledger of every invoice already imported (never re-listed)
  - `<Location>/PDFs/<date>/`   – signed invoice PDFs;  `<Location>/Booked/` – PDFs after import
  - `Archive/<Hub>_<date>/`     – move a batch here once imported
  - Per-hub files (`<Location>/<Location>_Entry.xlsx`, `1_<HUB>_...csv`) – only if you use the per-hub alternative

## Steps — Master flow (run from `0_Scripts/`, one file for every hub)
1. **Download PDFs**  (HUL + HUL SAMADHAN only) — *you already did this*
   `python download_hul_pdfs.py "C:\path\to\grn_copy_upload_data.csv"`
   → PDFs land in `5_NetSuite_Booking\<Location>\PDFs\<date>\`. Already-booked invoices are skipped.
2. **Build the master report**
   `python build_master_workbook.py "C:\path\to\grn_copy_upload_data.csv"`  (no path = uses the default GRN CSV)
   → writes `5_NetSuite_Booking\Master_Entry.xlsx` (the master report). Each row's **Attached file**
     cell links straight to its S3 PDF — click it to open, no separate links workbook.
     Already-booked invoices are skipped; invoices whose PDF link is missing are reported, not listed.
     An **About** sheet stamps the build date, the source CSV path, and the counts —
     so an old file opened later can explain which day / which CSV it came from.
3. **Fill amounts.** Open `Master_Entry.xlsx`, fill the YELLOW cells (`Amount_5%`, `Amount_18%`,
   and `IGST?` only for a rare inter-state row), save. Click a row's **Attached file** cell to open its
   PDF and read the taxable amounts. **Can't book a row** (wrong PDF / bad details)? Leave its amounts
   blank and type a reason in the **`Remark`** column (last column) — e.g. `wrong file`. Step 6
   collects every not-booked row + its remark into `MISSING_PDFS.xlsx` for you to chase.
4. **Generate the combined import files**
   `python generate_master_csvs.py`                     (address label included)
   `python generate_master_csvs.py --no-vendor-select`  (omit address)
   → `1_ALL_Bills_Header.csv`, `2_ALL_Bills_Expenses.csv`, `ALL_Tax_Verification.csv` (check the tax math),
     `ALL_Invoices.zip` (every referenced PDF).  Tax-math self-test: `python generate_master_csvs.py --selfcheck`.
5. **Import to NetSuite** (browser — see below), using the `ALL_` files + `ALL_Invoices.zip`.
6. **Mark booked** — only after a successful import
   `python mark_booked.py`  → appends every External ID to `_booked.csv` and moves each booked PDF
   from `<hub>\PDFs\` to `<hub>\Booked\`. Run it twice → 0 rows added (dedupe check).
7. **Archive the batch** — clears the top level for the next day
   `python archive_batch.py`  → moves the 6 batch files (`Master_*.xlsx`, `*ALL_*`) into
   `5_NetSuite_Booking\Archive\Master_<today>\` (date auto-filled; pass a date to override).
   `_booked.csv` and the hubs' `Booked\` PDFs stay put.

> **Per-hub alternative** (one hub at a time): `build_entry_workbook.py OMR` → fill `OMR_Entry.xlsx`
> → `generate_import_csvs.py OMR` [`--no-vendor-select`] → import → `mark_booked.py OMR`.

> **No GRN CSV, only pasted rows / links?** Use `download_from_links.py` in place of steps 1–2: paste your
> rows between the `PASTE = r"""` … `"""` markers, then `python download_from_links.py` → it downloads the
> PDFs **and** builds `<Hub>_Entry.xlsx` in one go. Continue from step 3.

## NetSuite import (manual, in browser)
1. `ALL_Invoices.zip` → File Cabinet (`Ganesh Folder`) via **Advanced Add** (Unzip ✔).
2. Setup ▸ Import/Export ▸ Import CSV Records → **Multiple files**:
   Primary = `1_ALL_Bills_Header.csv`, Linked = `2_ALL_Bills_Expenses.csv`, link by **External ID**.
3. Step 2 Advanced Options: **[✔] RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS**  (or GST/TDS won't calculate).
   Custom Form = `Ripplr Vendor Bill`.
4. Save & Run.  Then `python mark_booked.py`.

## Missing / wrong-file invoices → chase from the CDMS portal
Missing invoices collect in **`5_NetSuite_Booking\MISSING_PDFS.xlsx`** (two tabs):
- **`Missing_Links`** — written by `download_hul_pdfs.py`: invoices whose S3 link was blank.
- **`Wrong_File_Chase`** — written by `mark_booked.py` / `cdms_mark_booked.py`: rows listed in an entry
  workbook but not booked, with the `Remark` you typed (e.g. `wrong file`).

To pull those PDFs from CDMS (folder `0_Scripts\CDMS_Tool\`, full guide in its `HOW_TO_RUN.txt`):
1. Copy the `invoice_no` column from either tab into `CDMS_Tool\invoices.txt` (one per line).
2. Refresh the login token: Chrome ▸ CDMS GRN page ▸ F12 ▸ Network ▸ Fetch/XHR ▸ search an invoice ▸
   click the `list` request ▸ Request Headers ▸ copy the `authorization` value ▸ paste into `CDMS_Tool\token.txt`.
   (Token expires every 24 h; the script strips a leading `Bearer `.)
3. `cd CDMS_Tool` then `python download_invoices.py` (or double-click `run.bat`).
   → PDFs to `Desktop\CDMS_Invoices`, status report to `Desktop\CDMS_PDF_Result.xlsx`
   (`Downloaded` / `PDF Not attached` / `Not found`).

## Booking the CDMS-recovered invoices (a SEPARATE batch, kept in `CDMS_Sorted\`)
These are a different problem from the daily GRN flow, so they get their own entry workbook and
never mix with the top-level Master flow. Same columns as `Master_Entry.xlsx` (+ `Remark`), so
booking is identical. **The tracker only changes after NetSuite accepts** — mirrors the daily flow.
1. `python build_cdms_entry.py` → `CDMS_Sorted\CDMS_Entry.xlsx`, one row per **downloaded** invoice;
   moves each PDF off the Desktop into `CDMS_Sorted\<Hub>\PDFs\<date>\` (kept for the import zip). Each
   row's **Attached file** cell links to the **S3 URL** (opens in the browser, not Adobe). Ones that had
   a wrong PDF before are pre-flagged in `Remark` (VERIFY them).
2. Work `CDMS_Entry.xlsx`: **good** → fill `Amount_5%`/`Amount_18%`; **wrong/unusable** → type a
   `Remark` (e.g. `wrong file`), leave amounts blank. Click a row's **Attached file** cell to open its PDF.
3. `python cdms_generate.py` → `CDMS_Sorted\1_CDMS_Bills_Header.csv` + expenses + verification + zip.
   **Changes no tracker** — safe to re-run after a NetSuite rejection.
4. Import to NetSuite (same browser steps, `CDMS_` files + `CDMS_Invoices.zip`).
5. **Only after NetSuite accepts:** `python cdms_mark_booked.py` → books them in `_booked.csv`,
   moves their PDFs `CDMS_Sorted\<Hub>\PDFs\` → `<Hub>\Booked\`, drops them from `Missing_Links`,
   and sends any `Remark`-only rows to `Wrong_File_Chase`.
6. **Archive the CDMS batch** — `python archive_cdms.py` → moves that run's `CDMS_Entry.xlsx` +
   `1_CDMS_*` / `2_CDMS_*` / `CDMS_Tax_Verification.csv` / `CDMS_Invoices.zip` into
   `CDMS_Sorted\Archive\CDMS_<today>\` (date auto-filled), keeping a clean per-run record.

## Filing CDMS PDFs back into the hubs
`python sort_cdms.py`  → dry-run plan;  `python sort_cdms.py --go` → move.
Files `CDMS_Sorted\<Hub>\PDFs\<date>\*.pdf` into the real hub folders: booked → `<Hub>\Booked\`,
not-yet-booked → `<Hub>\PDFs\<date>\` (skips files already there). Maps `Mysore`→`Mysore Road`,
`YPR`→`Yeshwantpura`.

## Notes
- Check what's downloaded and what's booked: `python status.py` (or `python status.py <Hub>`) → per-hub
  summary + `5_NetSuite_Booking\STATUS.csv`, showing each invoice in plain words as BOOKED or not.
- HSN codes: `4090000` = 5% GST, `33059011` = 18% GST. TDS = 194Q @ 0.1%.
- Tax type is automatic: vendor GSTIN state == Place of Supply state → CGST+SGST, else IGST. Karnataka(29)→Karnataka(29) = CGST+SGST (NOT IGST). IGST only for the rare inter-state invoice (tick `IGST?` in the Entry sheet).
- Vendor Select address label (OMR): `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` (exact, trailing comma).
- Hub config (vendor/GSTIN/POS/brand/address) lives in the `HUBS` table in `build_entry_workbook.py`. All 10 hubs are set (OMR, Chrompet, Yeshwantpura, Soukya, Byrathi, Mysore Road, Gouribidanur, Hosur, Tumkur, Pondicherry). The master flow reads it directly — no per-hub Config sheet needed. Add a new hub there.

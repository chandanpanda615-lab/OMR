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
  - `MISSING_PDFS.xlsx` – the **Missing PDFs list** (tab `Chase`): every invoice not booked yet + why
  - `Gemini/raw/*.json` – every Gemini answer, one file per PDF (never paid twice);
    `Gemini/Gemini_Results.xlsx` – those answers vs what went into the import files
  - `CDMS_Recovered.csv` – PDFs CDMS found for the Missing PDFs list today (read by steps 1–2, archived in step 7)
  - `<Location>/PDFs/<date>/`   – signed invoice PDFs;  `<Location>/Booked/` – PDFs after import
  - `Archive/<Hub>_<date>/`     – move a batch here once imported
  - Per-hub files (`<Location>/<Location>_Entry.xlsx`, `1_<HUB>_...csv`) – only if you use the per-hub alternative

## Steps — Master flow (run from `0_Scripts/`, one file for every hub)
> **Steps 1–2 in one go:** `RUN_DAY.bat "C:\path\to\grn_copy_upload_data.csv"`
> Every step prints plain counts and updates the **Missing PDFs list** (`MISSING_PDFS.xlsx`).

1. **Download PDFs** (HUL + HUL SAMADHAN only)
   `python download_hul_pdfs.py "C:\path\to\grn_copy_upload_data.csv"`
   - **CDMS check first**, if `CDMS_Tool\token.txt` is still valid (refresh it by hand — see
     `CDMS_Tool\HOW_TO_RUN.txt`): every invoice on the Missing PDFs list is looked up in CDMS; PDFs
     found are downloaded in this same step. Token expired → "CDMS check SKIPPED", the rest still runs.
   - Prints a per-hub table:
     | Column | Meaning |
     |---|---|
     | **Downloaded now** | new PDF saved today into `<Hub>\PDFs\<yyyy-mm-dd>\` |
     | **Already on disk** | PDF was downloaded in an earlier run — not downloaded again (nothing lost) |
     | **Booked (skip)** | invoice is already in NetSuite (`_booked.csv`) — ignored |
     | **No link (list)** | GRN has no PDF link yet — put on the Missing PDFs list |
2. **Build the master report**
   `python build_master_workbook.py "C:\path\to\grn_copy_upload_data.csv"`
   → `5_NetSuite_Booking\Master_Entry.xlsx`: today's GRN rows + CDMS-recovered rows. Each row's
     **Attached file** cell links to its S3 PDF (opens in the browser). Rows whose PDF Gemini already
     read are **pre-filled from the saved answer** (free). Prints what was listed and what was not
     (already booked / no link / same wrong PDF / unknown hub — the last three are on the Missing PDFs list).
     An unknown hub name is printed as **UNKNOWN HUB** (add it to `ALIASES`/`HUBS` in `build_entry_workbook.py`).
     Refuses to overwrite a `Master_Entry.xlsx` that already has amounts (finish that batch first).
3. **Fill amounts — you choose, row by row.** Open `Master_Entry.xlsx`:
   - **type it yourself:** fill the YELLOW `Amount_5%` / `Amount_18%` (and `IGST?` for a rare inter-state row);
   - **no Gemini for this row:** type `x` in **`No Gemini (type x)`** (e.g. for the 5–6 you want to do by hand);
   - **can't book it** (wrong PDF / bad details): leave amounts blank, type a reason in **`Remark`** →
     it goes on the Missing PDFs list as "wrong file" and that same PDF is never listed again.
     A pre-filled "VERIFY" remark = this invoice had a wrong PDF before: check the new one: correct -> type the amounts (booked); still wrong -> type your own remark. Left untouched = listed again next day.
   - then, for all other rows, optionally: `python fill_tax_from_pdfs.py` (Gemini, needs `GEMINI_API_KEY`).
     It **never touches** a row you typed or marked `x`. Every answer is saved first in
     `5_NetSuite_Booking\Gemini\raw\` (one file per PDF) and the Master is filled from there — the same
     PDF is never paid for twice. A row is filled only when the tax matches the printed tax (±₹1);
     otherwise it is red **REVIEW** with the printed tax in **`Printed Tax (Gemini)`** — fill those by hand.
   - Save.
4. **Generate the combined import files**
   `python generate_master_csvs.py`                     (address label included)
   `python generate_master_csvs.py --no-vendor-select`  (omit address)
   → `1_ALL_Bills_Header.csv`, `2_ALL_Bills_Expenses.csv`, `ALL_Tax_Verification.csv`, `ALL_Invoices.zip`.
   **Money check:** if a row's amounts don't match the tax printed on its PDF (`Printed Tax (Gemini)`,
   more than ₹1 off), it STOPS and writes nothing — fix those rows. Only if you checked the PDF and
   Gemini misread the printed tax: `--accept=<invoice>,<invoice>`.
   Also updates the Missing PDFs list (rows left out of this batch) and
   `5_NetSuite_Booking\Gemini\Gemini_Results.xlsx` (every Gemini answer vs what went into the import).
   Tax-math self-test: `python generate_master_csvs.py --selfcheck`.
5. **Import to NetSuite** (browser — see below), using the `ALL_` files + `ALL_Invoices.zip`.
6. **Mark booked** — only after a successful import
   `python mark_booked.py`  → appends every External ID to `_booked.csv`, moves each booked PDF
   from `<hub>\PDFs\` to `<hub>\Booked\`, and updates the Missing PDFs list (booked leave it, not-booked
   rows join it). Run it twice → 0 rows added (dedupe check).
7. **Archive the batch** — clears the top level for the next day
   `python archive_batch.py`  → moves the batch files (`Master_*.xlsx`, `*ALL_*`, `CDMS_Recovered.csv`) into
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

## The Missing PDFs list — missing / wrong-file invoices (`5_NetSuite_Booking\MISSING_PDFS.xlsx`, tab `Chase`)
One row per invoice that is not booked yet, with `status`, your `remark`, the `bad_file` (the PDF
that was wrong), `first_seen` and `days_waiting`. Rows are only added or updated — no script wipes
another script's rows — and a row leaves the list only when the invoice is in `_booked.csv`.
| Status | Written by | What to do |
|---|---|---|
| `no S3 link (GRN)` | `download_hul_pdfs.py` | nothing — step 1 re-checks CDMS (when the token is valid) |
| `CDMS: PDF not attached yet` | `download_invoices.py` | ask the hub to upload the signed invoice |
| `CDMS: not found in portal` | `download_invoices.py` | check the invoice number |
| `wrong file …` / `CDMS: still the same wrong file` | `mark_booked.py` / `download_invoices.py` | ask the hub to re-upload the right PDF |
| `CDMS: PDF found …` | `download_invoices.py` | nothing — it is in the next `Master_Entry.xlsx` |
| `left blank in Master_Entry` | `mark_booked.py` | nothing — listed again next day |
| `unknown hub '…'` | `build_master_workbook.py` | add the spelling to `ALIASES` (or a new hub to `HUBS`) in `build_entry_workbook.py` |

CDMS check on its own / only some invoices: `python CDMS_Tool\download_invoices.py [9633128587 ...]`.
The old separate CDMS batch (`CDMS_Sorted\`, `build_cdms_entry` / `cdms_generate` / `cdms_mark_booked` /
`archive_cdms` / `sort_cdms`) was retired on 2026-09-22: scripts in `_legacy_cdms\`, data in
`5_NetSuite_Booking\Archive\CDMS_Sorted_retired_2026-09-22\`.

## Notes
- Check what's downloaded and what's booked: `python status.py` (or `python status.py <Hub>`) → per-hub
  summary + `5_NetSuite_Booking\STATUS.csv`, showing each invoice in plain words as BOOKED or not.
- HSN codes: `4090000` = 5% GST, `33059011` = 18% GST. TDS = 194Q @ 0.1%.
- Tax type is automatic: vendor GSTIN state == Place of Supply state → CGST+SGST, else IGST. Karnataka(29)→Karnataka(29) = CGST+SGST (NOT IGST). IGST only for the rare inter-state invoice (tick `IGST?` in the Entry sheet).
- Vendor Select address label (OMR): `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` (exact, trailing comma).
- Hub config (vendor/GSTIN/POS/brand/address) lives in the `HUBS` table in `build_entry_workbook.py`. All 10 hubs are set (OMR, Chrompet, Yeshwantpura, Soukya, Byrathi, Mysore Road, Gouribidanur, Hosur, Tumkur, Pondicherry). The master flow reads it directly — no per-hub Config sheet needed. Add a new hub there.

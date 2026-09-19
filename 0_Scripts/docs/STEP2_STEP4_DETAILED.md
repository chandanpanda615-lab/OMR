# Step 2 (build Excel) & Step 4 (generate import files) — detailed cases

Run everything from `C:\Users\chandan.p\Desktop\OMR\0_Scripts`.
`<Hub>` = the Location name, e.g. `OMR`, `Yeshwantpura`.

---

## STEP 2 — build the entry Excel

### Case A — you HAVE a grn CSV
The CSV must have columns: `fc_name, brand_name, invoice_no, invoice_date, s3_file_link`.
```
python download_hul_pdfs.py "C:\path\to\grn.csv"     # step 1: get the PDFs first
python build_entry_workbook.py <Hub> "C:\path\to\grn.csv"
```
- If the CSV is the default (`Downloads\grn_copy_upload_data (3).csv`) you can drop the path:
  `python build_entry_workbook.py <Hub>`
- It picks only that hub's rows + HUL/HUL SAMADHAN + rows that have an S3 link.
- One CSV can hold many hubs → run `build_entry_workbook.py` once **per hub**.
- Output: `5_NetSuite_Booking\<Hub>\<Hub>_Entry.xlsx`

### Case B — you DON'T have a CSV, only GRN links (pasted)
Do NOT run build directly. Use the paste tool — it does step 1 **and** step 2 together:
1. Open `download_from_links.py`, paste your rows/links between the `PASTE = r"""` … `"""` markers.
2. `python download_from_links.py`
   - downloads PDFs → `<Hub>\PDFs\<date>\`
   - writes `<Hub>\<Hub>_source.csv`
   - auto-builds `<Hub>\<Hub>_Entry.xlsx`
- Needs the hub to be known (in `HUBS` + `LOC`). New hub → do STEP 0 (see NEW_HUB_CHECKLIST.md) first.

### After you fill the Excel
- Fill yellow cells: `Amount_5%`, `Amount_18%`, `IGST?` (blank unless the invoice shows IGST).
- Cross-check the live `CGST 2.5% / SGST 2.5% / CGST 9% / SGST 9%` columns against each invoice.
- **SAVE the file** (Ctrl+S). If Excel stays open that's fine for reading, but the numbers must be saved.
  If Step 4 later gives a *permission* error, close Excel and re-run.
- Then go to STEP 4.

---

## STEP 4 — generate the import files
```
python generate_import_csvs.py <Hub>
```
Reads `<Hub>\<Hub>_Entry.xlsx` and writes into `5_NetSuite_Booking\<Hub>\`:
- `1_<Hub>_Bills_Header.csv`
- `2_<Hub>_Bills_Expenses.csv`
- `<Hub>_Tax_Verification.csv`
- `<Hub>_Invoices.zip`  (PDFs it found under `<Hub>\PDFs\`)

**Check the printout before importing:**
- `Tax split: X intra, Y IGST` — matches what you expect?
- `Zip: N/N PDFs` — must be all found. If it says `3/4`, one PDF is missing → download it first, else the
  NetSuite `custbody11` link fails.
- Any `WARN: GSTIN … fails checksum` → the GSTIN is wrong; fix it in the config and re-run.

### Flags
- `python generate_import_csvs.py <Hub> --no-vendor-select` → omit the address column
  (use if you get `Invalid billaddresslist reference key`).
- `python generate_import_csvs.py <Hub> --skip=9633119086,9629071934` → exclude already-booked invoices.

### "I generated the CSV but I need it again"
Just run `generate_import_csvs.py <Hub>` again — it **overwrites** the CSVs + zip from the current Excel.
Safe to run any number of times. Common reasons:
- You fixed an amount in the Excel → re-run to refresh the CSVs.
- Some invoices already booked, only the rest remain → add `--skip=<booked ids>`.
- You deleted/lost the files → re-run.
(External ID also blocks duplicates in NetSuite, so a re-import of an already-booked invoice just fails safely.)

---

## MULTIPLE EXCELS on the SAME DATE

### Different hubs, same day  → no conflict
Each hub has its own folder and its own `<Hub>_Entry.xlsx`. Just run the pipeline per hub:
```
python build_entry_workbook.py OMR
python build_entry_workbook.py Yeshwantpura
```
They never overwrite each other. Import each hub separately.

### Same hub, two separate batches, same day  → the Excel name would collide
`build_entry_workbook.py <Hub>` always writes the same `<Hub>_Entry.xlsx`, so building again **overwrites**.
Pick one:
- **Best (lazy) — make it ONE batch:** paste all the links together (or use one CSV that has all rows),
  build once, fill once, generate once, import once. Fewer imports, no collision.
- **Keep them separate:** finish batch 1 fully (build → fill → generate → import into NetSuite), then
  archive it — move `5_NetSuite_Booking\<Hub>\` files to `Archive\<Hub>_<date>_batch1\` — and only then
  build batch 2. This keeps a clean record of each batch.

Rule of thumb: **one Excel = one import**. If it can be one batch, keep it one batch.

---

## Quick recap
| You have | Do |
|---|---|
| grn CSV | `download_hul_pdfs.py <csv>` → `build_entry_workbook.py <Hub> <csv>` |
| only links | paste into `download_from_links.py` → run it (PDFs + Excel in one go) |
| filled Excel | save → `generate_import_csvs.py <Hub>` → check zip = N/N, no GSTIN warn |
| need CSVs again | re-run `generate_import_csvs.py <Hub>` (overwrites) |
| many hubs, same day | one folder + one Excel each; run per hub |
| same hub, 2 batches | combine into one, or finish+archive batch 1 before batch 2 |

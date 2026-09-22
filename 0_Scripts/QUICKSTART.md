# QUICKSTART — book a hub yourself (no Claude Code needed)

Run everything from this folder:
```
cd C:\Users\chandan.p\Desktop\OMR\0_Scripts
```

## A. If you have only pasted rows/links (no CSV)
1. Open `download_from_links.py`, paste your rows between the `PASTE = r"""` … `"""` lines (replace what's there).
2. Run it — downloads PDFs **and** builds the Excel in one go:
   ```
   python download_from_links.py
   ```
3. Open the built `..\5_NetSuite_Booking\<Hub>_Entry.xlsx`, go to step **C**.

## B. If you have a grn CSV
```
python download_hul_pdfs.py "C:\path\to\grn_copy.csv"
python build_entry_workbook.py <Hub>            # e.g. OMR  or  Yeshwantpura
```
Then open `..\5_NetSuite_Booking\<Hub>_Entry.xlsx`, go to step **C**.

## C. Fill the Excel (yellow cells only)
- **Amount_5%** = 5% taxable value from the invoice (blank if none)
- **Amount_18%** = 18% taxable value (blank if none)
- **IGST?** = leave blank normally; type `y` only if the invoice shows IGST
- Cross-check: the sheet shows **CGST 2.5% / SGST 2.5% / CGST 9% / SGST 9% / IGST** live —
  match each against the invoice's tax lines. If they match, the amounts are right.
- **Save** the file.

> Gemini auto-fill (`fill_tax_from_pdfs.py`) applies to the **Master** flow, not this per-hub quickstart — see **WORKFLOW.md**.

## D. Make the import files (CSVs + PDF zip)
```
python generate_import_csvs.py <Hub>
```
Produces in `..\5_NetSuite_Booking\<Hub>\`:
- `1_<Hub>_Bills_Header.csv`, `2_<Hub>_Bills_Expenses.csv`
- `<Hub>_Invoices.zip` (the PDFs, ready for File Cabinet)
Flags: `--no-vendor-select` (omit address), `--skip=id1,id2` (partial re-run).

Everything for a hub lives in its own folder: `5_NetSuite_Booking\<Hub>\`
(Entry.xlsx, source.csv, the 2 CSVs, the zip, and `PDFs\<date>\`).

## E. NetSuite (browser — the only manual part)
1. **Documents ▸ Files ▸ File Cabinet ▸ Ganesh Folder** → **Advanced Add** the `*_Invoices.zip`,
   tick **Unzip Files ✔** + **Overwrite ✔** → Add.  (Do this FIRST or the PDF link fails.)
2. **Setup ▸ Import/Export ▸ Import CSV Records** → Transactions ▸ Vendor Bill →
   2 files (Header + Expenses), link by **External ID**, Data Handling **Add**,
   **RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS ✔**, Custom Form **Ripplr Vendor Bill** → Save & Run.
3. Check **View CSV Import Status**. Green = done.

## New hub not yet known?
Add its row to `HUBS` in `build_entry_workbook.py` (copy the OMR/Yeshwantpura block) and, for the
paste flow, add its link-code to `LOC` in `download_from_links.py`. Gather the real values per
`docs\HUB_ONBOARDING_AND_PITFALLS.md` (Location, Vendor, GSTIN, address label — verify against a live bill).
```
```

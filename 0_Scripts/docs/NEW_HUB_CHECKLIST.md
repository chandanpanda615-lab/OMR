# New Location — how to approach it (self-serve)

A new hub needs **one-time setup** (gather its reference values + add them to the config).
After that, every run is the same 3–5 steps. Run all commands from:
```
cd C:\Users\chandan.p\Desktop\OMR\0_Scripts
```

---

## STEP 0 — One-time setup for a NEW hub (do this once per hub)

### 0a. Gather 6 reference values — from a REAL booked bill, not the tracker
Open a Vendor Bill for that hub that was **already booked manually** in NetSuite and read:

| Value | Where in NetSuite | Trap seen |
|---|---|---|
| **Location** (exact) | Bill ▸ Classification ▸ LOCATION | Not the short code (was `YSPR`, real = `Yeshwantpura`) |
| **Vendor** (`ID NAME`) | Bill ▸ VENDOR field | Copy verbatim (spelling e.g. `UNILIVER`) |
| **GSTIN** | Bill ▸ Tax Details ▸ Vendor Tax Reg. Number | Tracker GSTINs have been wrong/invalid — read the live one |
| **Place of Supply** | Bill ▸ PLACE OF SUPPLY | The hub's own state, e.g. `29-Karnataka` |
| **BRANDS / PROJECTS** | Bill ▸ BRANDS / PROJECTS | e.g. `HUL_YPR` |
| **Address LABEL** | Vendor record ▸ Address subtab ▸ **LABEL** column | Must be EXACT incl. typos/commas |

> Full detail + every trap: `HUB_ONBOARDING_AND_PITFALLS.md` (Section A).

### 0b. Add the hub to `build_entry_workbook.py` (copy an existing block)
Inside `HUBS = { ... }` add:
```python
"<hubname_lowercase>": {
    "Vendor": "RPP-DIST-XXXX HINDUSTAN UNILEVER LIMITED_HUL_XXX",
    "Location": "<exact Location>", "BRANDS / PROJECTS": "HUL_XXX",
    "Place of Supply": "29-Karnataka",
    "GSTIN intra": "<GSTIN from live bill>",
    "Address intra": "<exact address LABEL>",
    "GSTIN igst": "", "Address igst": "",   # fill ONLY if the hub ever gets inter-state invoices
},
```
- Intra-only hub → leave `GSTIN igst` / `Address igst` blank.
- Can get IGST sometimes → fill them with the vendor's *other* GSTIN + *other* label (like Yeshwantpura).

### 0c. If you'll use the paste flow, add its link-code
In `download_from_links.py`, add one line to `LOC`, e.g. `"BYR": "Byrathi"`
(the code is the first part of the PDF filename, e.g. `YSPR-HUL-07092026-...`).

---

## STEPS 1–5 — the routine (same for every hub, every day)

**1. Get the PDFs**
- Have a grn CSV → `python download_hul_pdfs.py "path\to\grn.csv"`
- Only pasted links → paste rows into `download_from_links.py`, then `python download_from_links.py`
  (downloads PDFs **and** builds the Excel in one run)

**2. Build the entry Excel** (skip if the paste flow already built it)
```
python build_entry_workbook.py <Hub>
```

**3. Fill the Excel**  → `5_NetSuite_Booking\<Hub>\<Hub>_Entry.xlsx`
- Yellow cells only: **Amount_5%**, **Amount_18%**, **IGST?** (blank unless the invoice shows IGST)
- Cross-check the live **CGST 2.5% / SGST 2.5% / CGST 9% / SGST 9%** columns against each invoice's tax lines → **save**
- Gemini auto-fill (`fill_tax_from_pdfs.py`) applies to the Master flow, not this per-hub checklist — see WORKFLOW.md.

**4. Make the import files**
```
python generate_import_csvs.py <Hub>
```
→ `1_<Hub>_Bills_Header.csv`, `2_<Hub>_Bills_Expenses.csv`, `<Hub>_Invoices.zip` (in the hub folder)
- If you see **`WARN: GSTIN … fails checksum`** → the GSTIN is wrong; fix it in the config before importing.

**5. NetSuite (browser — the only manual part)**
1. **Documents ▸ Files ▸ File Cabinet ▸ Ganesh Folder** ▸ **Advanced Add** the zip, **Unzip ✔** + **Overwrite ✔** — do this FIRST.
2. **Setup ▸ Import/Export ▸ Import CSV Records** ▸ Transactions ▸ Vendor Bill ▸ 2 files (Header + Expenses),
   link by **External ID**, Data Handling **Add**, **RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS ✔**,
   Custom Form **Ripplr Vendor Bill** ▸ **Save & Run**.
3. Check **View CSV Import Status**. Green = done → move the batch to `Archive\<Hub>_<date>\`.

---

## The 3 things that cause 90% of errors
1. **Wrong Location / GSTIN / address label** → gather from a live bill, exact (Step 0a).
2. **Forgot RUN SERVER SUITESCRIPT** → tax posts as 0.00.
3. **Imported the CSV before uploading the zip** → PDF link (`custbody11`) fails.

---

## Tax rule (never enter it by hand)
Vendor GSTIN first 2 digits **==** Place-of-Supply state code → **CGST + SGST**; different → **IGST**.
So Karnataka(29)→Karnataka(29) = CGST+SGST (NOT IGST). IGST only for the rare inter-state invoice (tick `IGST?`).
HSN `4090000` = 5% (2.5+2.5) · `33059011` = 18% (9+9) · TDS 194Q @ 0.1% of taxable.

## Shortcut
Send me the new hub's name + one screenshot of a manually-booked bill (Location, Vendor, GSTIN, POS, Brand)
+ the address LABEL from the vendor record → I'll drop the config block in, then you run steps 1–5 yourself.

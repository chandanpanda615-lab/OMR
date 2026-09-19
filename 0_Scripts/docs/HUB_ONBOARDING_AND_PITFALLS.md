# New-Hub Onboarding + Pitfalls Runbook (HUL → NetSuite)

Read this before booking a hub for the first time. It exists because the handover
tracker's reference values were sometimes **wrong**, and NetSuite rejects on exact-match.
Get every reference value from the **live system**, not the tracker.

---

## A. Reference details to collect ONCE per new hub

Collect all of these before generating CSVs. Best sources in brackets.
The tracker (`Purchase - Tracker (2).xlsx`) is a starting hint only — **verify each**.

| Field | Where to read it (authoritative) | Trap seen |
|---|---|---|
| **Location** (exact) | A manually-booked bill of that hub → **Classification ▸ LOCATION** | Tracker said `YSPR`; real value is `Yeshwantpura`. The short code is NOT the location. |
| **Vendor** (ledger name) | Same bill → **VENDOR** field, or Vendors list. Format = `<VENDOR ID> <NAME>` | Name has odd spelling `UNILIVER` (not UNILEVER) — copy verbatim. |
| **Vendor Tax Reg. Number** = GSTIN (intra) | Same bill → **Tax Details ▸ Vendor Tax Reg. Number**, or Vendor record ▸ tax subtab | Tracker's `29AAACH1004N1Z3` was **checksum-invalid** and rejected. Real = `29AAACH1004N1ZQ`. |
| **Place of Supply** | Same bill → **PLACE OF SUPPLY** | Constant per hub (hub's own state). E.g. Yeshwantpura always `29-Karnataka`. |
| **BRANDS / PROJECTS** | Same bill → **BRANDS / PROJECTS** | Yeshwantpura = `HUL_YPR` (not YSPR). |
| **Address LABEL** (Vendor Select) | Vendor record ▸ **Address** subtab ▸ **LABEL** column (not the address body) | Must be EXACT incl. typos + punctuation. OMR label ends with a trailing comma; YSP label misspells `LOGITICS` (not LOGISTICS). |
| **IGST variant** (if hub ever gets inter-state invoices) | The vendor's *other* tax reg + *other* address label | For YSP: IGST uses GSTIN `33AAACH1004N1Z1` + the "Ponds House, Chennai" label. |

Then add the hub as a row in `HUBS` inside `build_entry_workbook.py` (copy the OMR/Yeshwantpura block),
and add the hub's **paste link-code → Location** to the `LOC` map in `download_from_links.py`.

### A.1 The spelling traps that cause almost every rejection

NetSuite matches on **exact strings**. Every hub differs — never reuse another hub's value or "clean up" a typo. Verified across all hubs so far:

1. **Vendor name is spelled differently per hub — copy verbatim.**
   - Company spelling varies: `HINDUSTAN UNILEVER LIMITED` (OMR), `HINDUSTAN UNILIVER LIMITED` (Byrathi/Mysore Road/YSP — *UNILIVER* typo), `HINDUSTAN UNILEVER LTD.` (Gouribidanur/Hosur), mixed case `Hindustan Unilever Limited` (Tumkur), `Hindustan Uniliver Limited` (Soukya).
   - Suffix varies: most are `..._HUL_<Loc>`, but **Tumkur has NO `_HUL_`** — it's just `..._Tumkur`.
2. **Brand: the space after `HUL_` is inconsistent.** No space: `HUL_OMR`, `HUL_Chrompet`, `HUL_Soukya`, `HUL_Mysore Road`, `HUL_Gauribidanur`. **Space:** `HUL_ Byrathi`, `HUL_ Tumkur`, `HUL_ Hosur`.
3. **Address LABEL ≠ address body.** Copy the **LABEL column** (Vendor ▸ Address subtab) exactly — it is often truncated and carries typos/odd punctuation (misspelled `LOGITICS`, `CFA` vs `C/O`, `74-2` vs `74/2`, trailing/double commas, `Ponda` vs `Ponds`). Same physical warehouse → still a different label per hub.
4. **Location field spelling can differ from Vendor/Brand.** Gouribidanur: Location = `Gouribidanur` (OU) but Vendor/Brand use `Gauribidanur` (AU).
5. **Hosur is REVERSED (see Section B + E).** Its POS is Tamil Nadu, so its intra/IGST GSTIN+address roles are the opposite of the Karnataka hubs.

---

## B. Tax type is automatic — never enter it

Rule (SuiteTax uses the same one): compare **first 2 digits of Vendor GSTIN** vs **Place of Supply state code**.
- Same → **CGST + SGST** (e.g. 29→29, 33→33)
- Different → **IGST** (e.g. 29→33)

Because a hub can receive both kinds, the entry sheet has an **`IGST?`** column: leave blank for
normal (intra) invoices, type `y` only for the rare inter-state one. The generator then swaps
GSTIN **and** address to the hub's IGST variant, and NetSuite computes the tax. Place of Supply stays the hub's own state.

Rates: 5% slab (HSN `4090000`) = 2.5% CGST + 2.5% SGST, or 5% IGST. 18% slab (HSN `33059011`) = 9% + 9%, or 18% IGST. TDS = 194Q @ 0.1% of taxable.

> **Orientation matters.** The `Address intra`/`GSTIN intra` fields hold the variant whose **GSTIN state == the hub's POS state** (the CGST+SGST case); `*_igst` holds the other-state variant. For Karnataka hubs (POS 29) intra = the Karnataka `29…` GSTIN. For **Hosur (POS 33-Tamil Nadu)** it flips: intra = the Chennai `33…` GSTIN + Chennai address; IGST = the Karnataka `29…` GSTIN + Karnataka address. Set intra/igst by which GSTIN matches POS, not by "Karnataka = intra".

---

## C. Errors seen → cause → fix (quick lookup)

| Error in results.csv | Cause | Fix |
|---|---|---|
| `The field Amount (Req) has been removed from mapping` | Single-file CSV, sublist columns clash | Use the **2-file** import (Header + Expenses), link by External ID. |
| `Invalid custbody11 reference key <file>.pdf` | PDF not in File Cabinet yet | Upload the ZIP to File Cabinet (**Advanced Add ▸ Unzip ✔**) **before** the CSV import. |
| Bill saved but `Tax = 0.00`, no CGST/SGST | Server script didn't run | Step 2 Advanced Options: tick **RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS**. |
| `Invalid billaddresslist reference key <text> for entity <id>` | Vendor Select ≠ exact address LABEL | Use the exact LABEL text (typos/commas included), or omit Vendor Select (address blank, tax still correct). |
| `Invalid location reference key <X>` | Wrong Location string | Use the exact NetSuite Location name from a real bill (e.g. `Yeshwantpura`, not `YSPR`). |
| `Invalid entitytaxregnum reference key <gstin> for entity <id>` | GSTIN not registered on that vendor / bad checksum | Use the vendor's real GSTIN from a manual bill; generator now checksum-validates and warns. |
| Row fails as duplicate on re-run | External ID already booked | Expected & safe — External ID blocks duplicates. Use `--skip=<id>` to exclude already-booked rows cleanly. |

---

## D. The pipeline (commands)

Run from `0_Scripts\`. Outputs land in `..\5_NetSuite_Booking\`.

1. `python download_hul_pdfs.py "<grn_copy csv>"` → PDFs by Location\date, blanks logged to MISSING_PDFS.xlsx (Missing_Links tab)
2. `python build_entry_workbook.py <Hub fc_name>` → `<Location>_Entry.xlsx` (fill YELLOW amounts; tick IGST? if needed)
3. `python generate_import_csvs.py <Location>` → header + expenses CSV + tax-verification CSV
   - `--no-vendor-select` to omit the address column
   - `--skip=<id1,id2>` to exclude already-booked invoices on a partial re-run

Then in NetSuite: zip PDFs → File Cabinet (Advanced Add, Unzip) → Import CSV Records (2 files, link by External ID, **RUN SERVER SUITESCRIPT ✔**, Custom Form `Ripplr Vendor Bill`) → Save & Run. Archive the batch to `5_NetSuite_Booking\Archive\<Hub>_<date>\` when done.

---

## E. Hub configs (all values are exact — copy verbatim)

Status legend: **✅ verified** = booked through the pipeline in NetSuite; **🟡 wired** = in code + bill-verified, first pipeline run still pending; **⬜ not wired** = still needs Section-A collection.

### Paste link-code → Location (the `LOC` map in `download_from_links.py`)
The S3 filename prefix is **not** obvious; confirm from a real filename each time.

| Code | Location | | Code | Location |
|---|---|---|---|---|
| `OMR` | OMR | | `TLBL` | Mysore Road |
| `YSPR` | Yeshwantpura | | `GBDR` | Gouribidanur |
| `CRMP` | Chrompet | | `TMKR` | Tumkur |
| `SUKA` | Soukya | | `HSRA` | Hosur |
| `BYTI` | Byrathi | | | |

### Per-hub reference (HUBS key = Location lowercased)

**OMR** ✅ — Tamil Nadu, intra-only (33/33) · paste `OMR`
- Vendor `RPP-DIST-0396 HINDUSTAN UNILEVER LIMITED_HUL_OMR` · Brand `HUL_OMR` · POS `33-Tamil Nadu`
- Intra: GSTIN `33AAACH1004N1Z1` · label `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` (trailing comma) · no IGST

**Yeshwantpura** ✅ — Karnataka, mostly intra (29/29), sometimes IGST · paste `YSPR` · vendor id 13151
- Vendor `RPP-DIST-0014 HINDUSTAN UNILIVER LIMITED_HUL_YSP` · Brand `HUL_YPR` · POS `29-Karnataka`
- Intra: GSTIN `29AAACH1004N1ZQ` · label `HINDUSTAN UNILEVER LIMITED C/O LINFOX LOGITICS INDIA PVT LTD S NO. 74/2, GUVALAKANAHALLI NH NO.7. CHIKKABALLAPUR`
- IGST: GSTIN `33AAACH1004N1Z1` · label `Hindustan Unilever Ltd Ponds House, No:101 Santhome High Road`

**Chrompet** 🟡 — Tamil Nadu, intra-only (33/33) · paste `CRMP`
- Vendor `RPP-DIST-0388 HINDUSTAN UNILEVER LIMITED_HUL_Chrompet` · Brand `HUL_Chrompet` · POS `33-Tamil Nadu`
- Intra: GSTIN `33AAACH1004N1Z1` · label `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` · no IGST

**Soukya** 🟡 — Karnataka, intra + IGST · paste `SUKA`
- Vendor `RPP-DIST-0250 Hindustan Uniliver Limited_Hul_Soukya` · Brand `HUL_Soukya` · POS `29-Karnataka`
- Intra: GSTIN `29AAACH1004N1ZQ` · label `C/O LINFOX LOGITICS INDIA PVT LTD`
- IGST: GSTIN `33AAACH1004N1Z1` · label `101, PONDS HOUSE, SANTHOME HIGH ROAD, ,` (double comma)

**Byrathi** 🟡 (booked manually before, not yet via pipeline) — Karnataka, intra + IGST · paste `BYTI`
- Vendor `RPP-DIST-0370 HINDUSTAN UNILIVER LIMITED_HUL_Byrathi` · Brand `HUL_ Byrathi` (space) · POS `29-Karnataka`
- Intra: GSTIN `29AAACH1004N1ZQ` · label `HINDUSTAN UNILEVER LIMITED C/O LINFOX LOGITICS INDIA PVT LTD S NO. 74/2, GUVALAKANAHALLI NH NO.7. CHIKKABALLAPUR`
- IGST: GSTIN `33AAACH1004N1Z1` · label `No.101, Santhome High Road` (no trailing comma)

**Mysore Road** 🟡 (booked manually before, not yet via pipeline) — Karnataka, intra + IGST · paste `TLBL`
- Vendor `RPP-DIST-0017 HINDUSTAN UNILIVER LIMITED_HUL_MYSORE ROAD` · Brand `HUL_Mysore Road` · POS `29-Karnataka`
- Intra: GSTIN `29AAACH1004N1ZQ` · label `C/O LINFOX LOGITICS INDIA PVT LTD S NO. 74/2, GUVALAKANAHALLI NH NO.7.` (ends with period)
- IGST: GSTIN `33AAACH1004N1Z1` · label `No.101, Santhome High Road,` (trailing comma)

**Gouribidanur** 🟡 — Karnataka, intra-only · paste `GBDR`
- Vendor `RPP-DIST-0170 HINDUSTAN UNILEVER LTD._HUL_GAURIBIDANUR` · Brand `HUL_Gauribidanur` · POS `29-Karnataka`
- Location spelled **Gouribidanur (OU)**; Vendor/Brand spelled **Gauribidanur (AU)**
- Intra: GSTIN `29AAACH1004N1ZQ` · label `C/O. LINFOX LOGISTICS INDIA PVT LTD,S.NO.74/2,` · no IGST (vendor has only the KA address)

**Tumkur** 🟡 — Karnataka, intra-only · paste `TMKR`
- Vendor `RPP-DIST-0346 Hindustan Unilever Limited_Tumkur` (**no `_HUL_`**, mixed case) · Brand `HUL_ Tumkur` (space) · POS `29-Karnataka`
- Intra: GSTIN `29AAACH1004N1ZQ` · label `C/o linfox, Logistics india pvt ltd S No, 74/2,` · no IGST (vendor has only the KA address)

**Hosur** 🟡 — **Tamil Nadu, REVERSED orientation** (intra = Chennai/33, IGST = Karnataka/29) · paste `HSRA`
- Vendor `RPP-DIST-0162 HINDUSTAN UNILEVER LTD._HUL_HOSUR` · Brand `HUL_ Hosur` (space) · POS `33-Tamil Nadu`
- Intra (33/33 CGST+SGST, default): GSTIN `33AAACH1004N1Z1` · label `Ponda House, No:101 Santhome High` (`Ponda` not Ponds; truncated)
- IGST (29 vs 33): GSTIN `29AAACH1004N1ZQ` · label `CFA LINFOX LOGISTICS INDIA PVT LTD, S NO 74-2` (`CFA` not C/O; `74-2` hyphen)

**Pondicherry** ⬜ not wired — POS `34-Puducherry`, Vendor `RPP-DIST-0084 ..._HUL_Pondicherry`, Brand `HUL_Pondicherry`. Gather real values per Section A before first import (no invoices yet as of 2026-09-16).

Do NOT trust the old tracker GSTINs — checksum-validate each (the generator does this automatically).

---

## F. Constants (same for all hubs unless noted)
Account `50020 Cost of Goods Sold : Purchase` · Class `Distribution` · Department `Sales Cost` ·
India Tax Section Code `194Q TDS on Purchases` · India Tax Nature `Goods` · Currency `INR` · Exchange Rate `1.00` ·
Memo `Being Purchase for the month of Sep'26` · dates `DD/MM/YYYY` · Posting Period `Mmm YYYY` (update per month) ·
Custom Form `Ripplr Vendor Bill` · HSN `4090000`=5%, `33059011`=18%.

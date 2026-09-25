# GRN ↔ NetSuite Reconciliation

Daily check that each HUL hub's GRN (goods received) matches what was booked in
NetSuite, by comparing the **taxable value** per invoice.

## Daily steps

1. Create a folder for the day: `RECO_<day>` (e.g. `RECO_19th`).
2. Put into it:
   - **All hub GRN files** — the "Product Wise Purchase" export per hub, **renamed to the hub
     name**: `Byrathi.xlsx`, `Chrompet.xlsx`, `Gouribidanur.xlsx`, `Hosur.xlsx`, `Mysore Road.xlsx`,
     `OMR.xlsx`, `Soukya.xlsx`, `Tumkur.xlsx`, `Yeshwantpura.xlsx` (same spelling as the NetSuite
     Location). One file per hub. A name that is not a hub is printed as a WARNING.
     LeverEDGE codes: 41B534 Hosur, 41B878 Chrompet, 41B990 OMR, 43A573 Yeshwantpura,
     43A593 Byrathi, 43A633 Mysore Road, 43A766 Gouribidanur, 43A808 Soukya, 43A919 Tumkur
     (Pondicherry: no GRN export yet - its bills show as NS_ONLY).
   - **The NetSuite export, renamed to exactly `NS.xlsx`.** This is the only
     way the script tells the NS file apart from the GRN files.
3. Run it:
   - Double-click **`RUN_RECO.bat`**, or
   - `python reconcile.py` (lists the folders and asks which one), or
   - `python reconcile.py RECO_19th` to point at a specific folder.
4. Open **`RECO_<day>/Reconciliation_result.xlsx`**.

> Close `Reconciliation_result.xlsx` in Excel before re-running, or the script
> can't overwrite it.

## Output sheets

| Sheet | What's in it |
|-------|--------------|
| `Hub_Summary` | **Pivot per hub** (all 10 hubs + TOTAL): GRN vs NS taxable, Match / Mismatch, GRN only split into booked-not-in-NS-yet / on Missing PDFs list / **NOT booked, NOT chased**, NS only. **Click a hub name to jump to its own sheet** |
| one sheet per hub (`OMR`, `Chrompet`, …) | That hub's invoices only, same columns as `Reconciliation`. `<< Hub_Summary` in A1 goes back |
| `Reconciliation` | One row per invoice: Hub, GRN taxable, NS taxable, Diff, Status, **Track** |
| `Summary` | Count of invoices per Status |
| `GRN_Consolidated_EGIR` | All EGIR GRN lines from every hub, with a `Hub` column |

### Status meanings

| Status | Meaning | Action |
|--------|---------|--------|
| `MATCH` | Taxable agrees within ₹1 | none |
| `MISMATCH` | Both sides have it, taxable differs > ₹1 | check the invoice |
| `HUB_MISMATCH` | Booked under a different hub than the GRN | fix the bill's vendor/location |
| `GRN_ONLY (not in NS)` | Received but **not booked** in NetSuite | see **Track**: booked after this NS export / on the Missing PDFs list / **NOT BOOKED, NOT CHASED** = fell through the cracks, act now |
| `NS_ONLY (not in GRN)` | Booked, but no GRN file supplied | send that hub's GRN and re-run |

## The rule (how taxable is computed)

Verified against NetSuite to the paisa on invoice `9633114313` (Hosur).

- GRN files have the header on **row 8**, data from row 9; the trailing
  `GRAND TOTAL` row is dropped.
- Keep **`Invoice Type == EGIR`** only. DSE / STN lines are ignored (DSE lines
  can be negative and would distort the total).
- **GRN taxable** = `NetAmt − CGST Amt − SGST Amt − IGST Amt`, summed per
  `Supplier Invoice No`.
- **NS taxable** = `Amount`, summed per `Document Number` (a bill can span two
  GL rows).
- Match key = the invoice number (`Supplier Invoice No` ↔ `Document Number`).
- Invoices with **zero taxable** are ignored on both sides.

## Settings (top of `reconcile.py`)

- `TOL = 1.0` — rupee tolerance; a diff below this counts as `MATCH`. Raise it
  if a few rupees of rounding should not be flagged.
- `NS_NAME = "NS.xlsx"` — the required NetSuite filename.

## If a run fails

- **"No 'NS.xlsx' found"** — you didn't rename the NetSuite export to `NS.xlsx`.
- **Permission / cannot write** — `Reconciliation_result.xlsx` is open in Excel.
- **Wrong / empty numbers** — the GRN column layout changed (header no longer on
  row 8, or columns renamed). The script expects the standard Product Wise
  Purchase layout.

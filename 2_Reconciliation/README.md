# GRN ↔ NetSuite Reconciliation

Daily check that each HUL hub's GRN (goods received) matches what was booked in
NetSuite, by comparing the **taxable value** per invoice.

## Daily steps

1. Create a folder for the day: `RECO_<day>` (e.g. `RECO_19th`).
2. Put into it:
   - **All hub GRN files** — the "Product Wise Purchase" export per hub
     (`Byrathi.xlsx`, `Chromepet.xlsx`, `OMR.xlsx`, …). One file per hub.
   - **The NetSuite export, renamed to exactly `NS.xlsx`.** This is the only
     way the script tells the NS file apart from the GRN files.
3. Run it:
   - Double-click **`RUN_RECO.bat`**, or
   - `python reconcile.py` (auto-uses the newest `RECO_*` folder), or
   - `python reconcile.py RECO_19th` to point at a specific folder.
4. Open **`RECO_<day>/Reconciliation_result.xlsx`**.

> Close `Reconciliation_result.xlsx` in Excel before re-running, or the script
> can't overwrite it.

## Output sheets

| Sheet | What's in it |
|-------|--------------|
| `Reconciliation` | One row per invoice: GRN taxable, NS taxable, Diff, Status |
| `Summary` | Count of invoices per Status |
| `GRN_Consolidated_EGIR` | All EGIR GRN lines from every hub, with a `Hub` column |

### Status meanings

| Status | Meaning | Action |
|--------|---------|--------|
| `MATCH` | Taxable agrees within ₹1 | none |
| `MISMATCH` | Both sides have it, taxable differs > ₹1 | check the invoice |
| `GRN_ONLY (not in NS)` | Received but **not booked** in NetSuite | book it |
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

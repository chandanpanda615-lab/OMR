"""Make the CDMS batch import files from CDMS_Sorted\\CDMS_Entry.xlsx. CSV + zip ONLY -
it changes NO tracker/ledger, so it is safe to re-run after a NetSuite rejection.
Run cdms_mark_booked.py only AFTER NetSuite has accepted the bills.

    python cdms_generate.py                    # address label included
    python cdms_generate.py --no-vendor-select # omit the Vendor Select address column

Rows with amounts -> booked-ready. Rows with only a Remark are listed (for your eye) but
NOT written anywhere yet - cdms_mark_booked.py handles them after booking.

Reads : CDMS_Sorted\\CDMS_Entry.xlsx
Writes: CDMS_Sorted\\1_CDMS_Bills_Header.csv, 2_CDMS_Bills_Expenses.csv,
        CDMS_Tax_Verification.csv, CDMS_Invoices.zip
"""
import csv, os, sys, zipfile
from openpyxl import load_workbook
from build_entry_workbook import HUBS, config_rows, ENTRY_HEADERS
from generate_import_csvs import HEADER_COLS, EXP_COLS, process_entry_row, fmt_date

OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
CDMS_ROOT = os.path.join(OUT_ROOT, "CDMS_Sorted")
ENTRY = os.path.join(CDMS_ROOT, "CDMS_Entry.xlsx")
HUB_IDX, REM_IDX = len(ENTRY_HEADERS), len(ENTRY_HEADERS) + 1


def cfg_maps():
    cfg_by_loc, pos_by_loc = {}, {}
    for k in HUBS:
        C = dict(config_rows(k)); loc = HUBS[k]["Location"]
        cfg_by_loc[loc] = C
        pos_by_loc[loc] = str(C["Place of Supply"]).split("-")[0].strip()
    return cfg_by_loc, pos_by_loc


def main():
    include_vs = "--no-vendor-select" not in sys.argv
    cfg_by_loc, pos_by_loc = cfg_maps()

    header_rows, exp_rows, verify_rows, warnings = [], [], [], []
    flagged, undecided = [], 0
    for r in load_workbook(ENTRY, data_only=True)["Entry"].iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        inv = str(r[0]).strip()
        hub = str(r[HUB_IDX]).strip() if len(r) > HUB_IDX and r[HUB_IDX] else ""
        remark = str(r[REM_IDX]).strip() if len(r) > REM_IDX and r[REM_IDX] else ""
        a5, a18 = r[4], r[5]
        if not (a5 in (None, "") and a18 in (None, "")):
            if hub not in cfg_by_loc:
                warnings.append(f"{inv}: unknown hub '{hub}' - skipped"); continue
            h, e, v, w = process_entry_row(inv, fmt_date(r[1]), r[2], r[3], a5, a18, r[6],
                                           cfg_by_loc[hub], pos_by_loc[hub])
            warnings += w
            if h is None:
                continue
            header_rows.append(h); exp_rows += e; verify_rows.append(v)
        elif remark:
            flagged.append((inv, remark))
        else:
            undecided += 1

    assert header_rows, ("No booked-ready rows (no amounts filled in CDMS_Entry.xlsx). "
                         "Fill amounts, or if all remaining are wrong/blank run cdms_mark_booked.py.")

    header_cols = list(HEADER_COLS)
    if not include_vs:
        header_cols.remove("Vendor Select")
        for h in header_rows:
            h.pop("Vendor Select", None)
    hp = os.path.join(CDMS_ROOT, "1_CDMS_Bills_Header.csv")
    ep = os.path.join(CDMS_ROOT, "2_CDMS_Bills_Expenses.csv")
    vp = os.path.join(CDMS_ROOT, "CDMS_Tax_Verification.csv")
    with open(hp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=header_cols); w.writeheader(); w.writerows(header_rows)
    with open(ep, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=EXP_COLS); w.writeheader(); w.writerows(exp_rows)
    with open(vp, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Invoice", "TaxType", "Taxable5%", "Taxable18%", "TotTaxable",
                    "CGST", "SGST", "IGST", "TotGST", "TDS194Q", "NetPayable"])
        for row in verify_rows:
            w.writerow([row[0], row[1]] + [str(x) for x in row[2:]])

    wanted = [h["Attached file"] for h in header_rows if h.get("Attached file")]
    found = {}
    for loc in {h["Location"] for h in header_rows}:
        for dp, _, files in os.walk(os.path.join(CDMS_ROOT, loc, "PDFs")):
            for fn in files:
                if fn in wanted and fn not in found:
                    found[fn] = os.path.join(dp, fn)
    zp = os.path.join(CDMS_ROOT, "CDMS_Invoices.zip")
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        for fn in wanted:
            if fn in found:
                z.write(found[fn], fn)
    for fn in wanted:
        if fn not in found:
            warnings.append(f"PDF not found under CDMS_Sorted: {fn}")

    igst_n = sum(1 for v in verify_rows if v[1] == "IGST")
    print(f"Header : {hp}  ({len(header_rows)} bills)")
    print(f"Expenses: {ep}  ({len(exp_rows)} lines)")
    print(f"Verify : {vp}")
    print(f"Zip : {zp}  ({sum(1 for fn in wanted if fn in found)}/{len(wanted)} PDFs)")
    print(f"Tax split: {len(verify_rows) - igst_n} intra (CGST+SGST), {igst_n} IGST")
    print(f"Not booked yet - flagged(remark): {len(flagged)} | blank(to work): {undecided}")
    print("Import these to NetSuite. AFTER it accepts, run:  python cdms_mark_booked.py")
    for w_ in warnings:
        print("WARN:", w_)


if __name__ == "__main__":
    main()

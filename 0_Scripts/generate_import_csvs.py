"""Generate the 2 NetSuite import CSVs (+ a tax verification sheet) from a filled <HUB>_Entry.xlsx.

Usage:
    python generate_import_csvs.py OMR
    python generate_import_csvs.py YSPR
    python generate_import_csvs.py YSPR --no-vendor-select   # omit address (Option B)

Per invoice, the IGST? flag in the Entry sheet swaps GSTIN + address between the hub's
intra and IGST variants. Tax type (CGST+SGST vs IGST) is derived automatically by
comparing the GSTIN state code to the Place-of-Supply state code -- never asked.

Reads : <OUT_ROOT>\\<LOC>_Entry.xlsx
Writes: 1_<LOC>_Bills_Header.csv, 2_<LOC>_Bills_Expenses.csv, <LOC>_Tax_Verification.csv
"""
import csv, sys, os, zipfile
from datetime import datetime, date as _date
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import load_workbook

# Project root = parent of 0_Scripts; self-locating so the top folder can be renamed/moved.
OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")

HEADER_COLS = ["External ID", "Reference No.", "Vendor", "Vendor Tax Reg. Number",
               "Vendor Select", "Date", "Posting Period", "Invoice Date",
               "Invoice Receipt Date", "Due Date", "Tax Point Date", "Approval Status",
               "Currency", "Exchange Rate", "Department", "Class", "Location",
               "BRANDS / PROJECTS", "Place of Supply", "Attached file", "Memo"]
EXP_COLS = ["External ID", "Line", "Account", "Amount", "Memo", "Department", "Class",
            "Location", "BRANDS / PROJECTS", "INDIA TAX SECTION CODE",
            "INDIA TAX HSN OR SAC CODE", "India Tax Nature"]
TRUTHY = {"y", "yes", "igst", "1", "true", "x"}


def money(v):
    return f"{float(v):.2f}" if v not in (None, "") else None


def fmt_date(v):
    # NetSuite wants DD/MM/YYYY. A manually edited Excel date cell comes back from
    # openpyxl as a datetime -> serialises "2026-09-12 00:00:00", which NetSuite
    # rejects. Untouched cells are already the text build_entry_workbook wrote, so
    # pass those straight through.
    return v.strftime("%d/%m/%Y") if isinstance(v, (datetime, _date)) else v


_CP = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def gstin_ok(g):
    # validates the 15th-char checksum; catches bad tracker GSTINs before import
    g = (g or "").strip()
    if len(g) != 15:
        return False
    factor, s = 2, 0
    for ch in reversed(g[:14]):
        if ch not in _CP:
            return False
        d = factor * _CP.index(ch)
        factor = 1 if factor == 2 else 2
        s += d // 36 + d % 36
    return _CP[(36 - s % 36) % 36] == g[14]


def q(x):
    return Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP)


def process_entry_row(inv, idate, brand, pdf, a5, a18, igst_flag, C, pos_state):
    """Turn one filled Entry row into (header_dict, expense_rows, verify_row, warnings).
    header_dict is None when the row has no amounts. Vendor Select is always populated;
    callers running --no-vendor-select drop the column afterwards. Shared by the per-hub
    and master generators so the tax math lives in exactly one place."""
    warnings = []
    if a5 in (None, "") and a18 in (None, ""):
        return None, [], None, [f"{inv}: no amounts entered - skipped"]

    is_igst = str(igst_flag or "").strip().lower() in TRUTHY and bool(C.get("Vendor Tax Reg. Number (IGST)"))
    gstin = C["Vendor Tax Reg. Number (IGST)"] if is_igst else C["Vendor Tax Reg. Number"]
    if not gstin_ok(gstin):
        warnings.append(f"{inv}: GSTIN '{gstin}' fails checksum - NetSuite will reject it")
    vs = C["Vendor Select (IGST address label)"] if is_igst else C["Vendor Select (address label)"]

    header = {
        "External ID": inv, "Reference No.": inv, "Vendor": C["Vendor"],
        "Vendor Tax Reg. Number": gstin, "Vendor Select": vs,
        "Date": idate, "Posting Period": C["Posting Period"], "Invoice Date": idate,
        "Invoice Receipt Date": idate, "Due Date": idate, "Tax Point Date": idate,
        "Approval Status": C["Approval Status"], "Currency": C["Currency"],
        "Exchange Rate": C["Exchange Rate"], "Department": C["Department"],
        "Class": C["Class"], "Location": C["Location"],
        "BRANDS / PROJECTS": C["BRANDS / PROJECTS"], "Place of Supply": C["Place of Supply"],
        "Attached file": pdf, "Memo": C["Memo"],
    }

    exp = []
    line = 0
    for amt, hsn in ((a5, C["HSN 5%"]), (a18, C["HSN 18%"])):
        if amt in (None, "") or float(amt) == 0:
            continue
        line += 1
        exp.append({
            "External ID": inv, "Line": line, "Account": C["Account"],
            "Amount": money(amt), "Memo": C["Memo"], "Department": C["Department"],
            "Class": C["Class"], "Location": C["Location"],
            "BRANDS / PROJECTS": C["BRANDS / PROJECTS"],
            "INDIA TAX SECTION CODE": C["INDIA TAX SECTION CODE"],
            "INDIA TAX HSN OR SAC CODE": hsn, "India Tax Nature": C["India Tax Nature"],
        })

    # verification: tax type from GSTIN state vs POS state
    t5 = Decimal(str(a5 or 0)); t18 = Decimal(str(a18 or 0))
    intra = gstin[:2] == pos_state
    cgst = sgst = igst = Decimal("0")
    if intra:
        cgst = q(t5 * Decimal("0.025")) + q(t18 * Decimal("0.09")); sgst = cgst
    else:
        igst = q(t5 * Decimal("0.05")) + q(t18 * Decimal("0.18"))
    taxable = t5 + t18
    tds = q(taxable * Decimal("0.001"))
    payable = taxable + cgst + sgst + igst - tds
    verify = [inv, "IGST" if not intra else "CGST+SGST", t5, t18, taxable,
              cgst, sgst, igst, cgst + sgst + igst, tds, payable]
    return header, exp, verify, warnings


def main():
    loc = sys.argv[1] if len(sys.argv) > 1 else "OMR"
    include_vs = "--no-vendor-select" not in sys.argv
    # --skip 111,222 : External IDs already booked (partial re-run), excluded from output.
    skip = set()
    for a in sys.argv:
        if a.startswith("--skip="):
            skip = {s.strip() for s in a.split("=", 1)[1].split(",") if s.strip()}

    hub_dir = os.path.join(OUT_ROOT, loc)   # per-hub folder holds Entry, CSVs, zip, PDFs
    wb = load_workbook(os.path.join(hub_dir, f"{loc}_Entry.xlsx"), data_only=True)
    C = {r[0].value: r[1].value for r in wb["Config"].iter_rows(min_row=2) if r[0].value}
    pos_state = str(C["Place of Supply"]).split("-")[0].strip()

    header_rows, exp_rows, verify_rows, warnings = [], [], [], []
    for r in wb["Entry"].iter_rows(min_row=2, values_only=True):
        inv = r[0]
        if inv is None:
            continue
        if str(inv) in skip:
            continue
        h, e, v, w = process_entry_row(inv, fmt_date(r[1]), r[2], r[3], r[4], r[5], r[6], C, pos_state)
        warnings += w
        if h is None:
            continue
        header_rows.append(h); exp_rows += e; verify_rows.append(v)

    if not include_vs:
        HEADER_COLS.remove("Vendor Select")
        for h in header_rows:
            h.pop("Vendor Select", None)

    h_path = os.path.join(hub_dir, f"1_{loc}_Bills_Header.csv")
    e_path = os.path.join(hub_dir, f"2_{loc}_Bills_Expenses.csv")
    v_path = os.path.join(hub_dir, f"{loc}_Tax_Verification.csv")
    with open(h_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HEADER_COLS); w.writeheader(); w.writerows(header_rows)
    with open(e_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=EXP_COLS); w.writeheader(); w.writerows(exp_rows)
    with open(v_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Invoice", "TaxType", "Taxable5%", "Taxable18%", "TotTaxable",
                    "CGST", "SGST", "IGST", "TotGST", "TDS194Q", "NetPayable"])
        for row in verify_rows:
            w.writerow([row[0], row[1]] + [str(v) for v in row[2:]])

    # Bundle the PDFs these bills reference into a zip for File Cabinet (Advanced Add > Unzip).
    wanted = [h["Attached file"] for h in header_rows if h.get("Attached file")]
    found = {}
    for dp, _, files in os.walk(os.path.join(hub_dir, "PDFs")):
        for fn in files:
            if fn in wanted and fn not in found:
                found[fn] = os.path.join(dp, fn)
    z_path = os.path.join(hub_dir, f"{loc}_Invoices.zip")
    with zipfile.ZipFile(z_path, "w", zipfile.ZIP_DEFLATED) as z:
        for fn in wanted:
            if fn in found:
                z.write(found[fn], fn)
    for fn in wanted:
        if fn not in found:
            warnings.append(f"PDF not found under {loc}\\PDFs: {fn}")

    print(f"Header : {h_path}  ({len(header_rows)} bills)")
    print(f"Expenses: {e_path}  ({len(exp_rows)} lines)")
    print(f"Verify : {v_path}")
    print(f"Zip : {z_path}  ({sum(1 for fn in wanted if fn in found)}/{len(wanted)} PDFs)")
    igst_n = sum(1 for v in verify_rows if v[1] == "IGST")
    print(f"Tax split: {len(verify_rows) - igst_n} intra (CGST+SGST), {igst_n} IGST")
    print(f"Vendor Select (address): {'INCLUDED' if include_vs else 'OMITTED'}")
    for w_ in warnings:
        print("WARN:", w_)
    assert header_rows, "No bills generated - did you fill the amounts and save?"


if __name__ == "__main__":
    main()

"""Consolidate hub GRN files + reconcile taxable value against the NetSuite export.

GRN files: header on row 8, data from row 9, trailing "GRAND TOTAL" row dropped.
Rule (verified on invoice 9633114313, Hosur): keep Invoice Type == EGIR only;
GRN taxable = NetAmt - CGST - SGST - IGST per line, summed per Supplier Invoice No.
NS taxable = Amount summed per Document Number. Match on the invoice number.
Zero-taxable invoices are ignored on both sides.

Usage: python reconcile.py [folder]
  - folder given  -> reconcile exactly that folder (name relative to this script, or a full path)
  - no folder     -> list the day-folders and ask which one (no silent auto-pick)
"""
import sys, glob, os, re
import pandas as pd
from openpyxl.worksheet.hyperlink import Hyperlink

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "0_Scripts"))
from build_entry_workbook import HUBS, hub_key    # hub names = NetSuite Location names
from mark_booked import load_booked              # _booked.csv
import missing_tracker as mt                     # Missing PDFs list
HUB_NAMES = [h["Location"] for h in HUBS.values()]
NOT_CHASED = "NOT BOOKED, NOT CHASED"


def resolve(folder):
    """Accept a full path or a plain folder name next to this script."""
    for cand in (folder, os.path.join(HERE, folder)):
        if os.path.isdir(cand):
            return cand
    sys.exit(f"Folder not found: {folder}")


def choose_folder():
    """No folder given -> list day-folders that hold xlsx files and ask. Never guess."""
    dirs = sorted(d for d in os.listdir(HERE)
                  if os.path.isdir(os.path.join(HERE, d))
                  and glob.glob(os.path.join(HERE, d, "*.xlsx")))
    if not dirs:
        sys.exit("No folder with xlsx files found. Create one (e.g. RECO_19th) and drop the GRN + NS files in it.")
    print("Which folder to reconcile?")
    for i, d in enumerate(dirs, 1):
        print(f"  {i}. {d}")
    pick = input("Enter number or folder name: ").strip()
    if pick.isdigit() and 1 <= int(pick) <= len(dirs):
        return os.path.join(HERE, dirs[int(pick) - 1])
    return resolve(pick)


FOLDER = resolve(sys.argv[1]) if len(sys.argv) > 1 else choose_folder()
TOL = 1.0  # rupees; below this a diff is treated as a rounding match
NS_NAME = "NS.xlsx"  # rename the NetSuite export to exactly this
OUT_NAME = "Reconciliation_result.xlsx"  # our own output — never read it back as a GRN

TAX_COLS = ["CGST Amt", "SGST Amt", "IGST Amt"]
NUM_COLS = TAX_COLS + ["NetAmt"]


def load_grn(path):
    df = pd.read_excel(path, header=7, dtype=str)
    df = df[df["Sr No"].notna() & (df["Item Name"] != "GRAND TOTAL")].copy()
    df = df[df["Invoice Type"] == "EGIR"]                       # ignore DSE/STN/etc.
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Taxable"] = df["NetAmt"] - df[TAX_COLS].sum(axis=1)
    name = os.path.splitext(os.path.basename(path))[0]      # rename each GRN file to its hub: OMR.xlsx
    k = hub_key(name)
    df["Hub"] = HUBS[k]["Location"] if k in HUBS else name
    if k not in HUBS:
        print(f"WARNING: '{os.path.basename(path)}' is not a hub name - rename it to one of: {', '.join(HUB_NAMES)}")
    return df


def hub_summary(rec, grn_hubs):
    """One row per hub (all hubs, even with no GRN file) + TOTAL: the pivot of the Reconciliation sheet."""
    rows = []
    for hub in HUB_NAMES + sorted(set(rec["Hub"]) - set(HUB_NAMES)):
        x = rec[rec["Hub"] == hub]
        go = x[x["Status"].str.startswith("GRN_ONLY")]
        nc = go[go["Track"] == NOT_CHASED]
        ns_only = x[x["Status"].str.startswith("NS_ONLY")]
        rows.append({
            "Hub": hub,
            "GRN file": "yes" if hub in grn_hubs else "NO GRN FILE",
            "GRN invoices": int((x["GRN_Taxable"] != 0).sum()),
            "GRN taxable": round(x["GRN_Taxable"].sum(), 2),
            "NS taxable": round(x["NS_Taxable"].sum(), 2),
            "GRN - NS": round(x["GRN_Taxable"].sum() - x["NS_Taxable"].sum(), 2),
            "Match": int((x["Status"] == "MATCH").sum()),
            "Mismatch (amount/hub)": int(x["Status"].isin(["MISMATCH", "HUB_MISMATCH"]).sum()),
            "GRN only": len(go),
            "..booked, not in NS export yet": int(go["Track"].str.startswith("booked").sum()),
            "..on Missing PDFs list": int(go["Track"].str.startswith("Missing").sum()),
            "..NOT booked, NOT chased": len(nc),
            "..NOT chased taxable": round(nc["GRN_Taxable"].sum(), 2),
            "NS only": len(ns_only),
            "NS only taxable": round(ns_only["NS_Taxable"].sum(), 2),
        })
    s = pd.DataFrame(rows)
    total = {c: (s[c].sum() if c not in ("Hub", "GRN file") else "") for c in s.columns}
    total["Hub"] = "TOTAL"
    return pd.concat([s, pd.DataFrame([total])], ignore_index=True)


def write_hub_sheets(xl, hubs, rec):
    """One sheet per hub that has rows; hub name in Hub_Summary links to it, and back."""
    summary_ws = xl.book["Hub_Summary"]
    for i, hub in enumerate(hubs["Hub"], start=2):      # row 1 = header
        rows = rec[rec["Hub"] == hub]
        if hub == "TOTAL" or rows.empty:
            continue
        name = re.sub(r"[\\/*?:\[\]]", "-", str(hub))[:31]
        rows.to_excel(xl, sheet_name=name, index=False, startrow=1)  # row 1 kept for the back link
        back = xl.book[name]["A1"]
        back.value, back.style = "<< Hub_Summary", "Hyperlink"
        back.hyperlink = Hyperlink(ref="A1", location="Hub_Summary!A1")   # internal link needs location, not target
        xl.book[name].freeze_panes = "A3"
        cell = summary_ws.cell(row=i, column=1)
        cell.style = "Hyperlink"
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{name}'!A1")


def main():
    files = [f for f in glob.glob(os.path.join(FOLDER, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")
             and os.path.basename(f) != OUT_NAME]
    ns_path = next((f for f in files if os.path.basename(f).lower() == NS_NAME.lower()), None)
    if ns_path is None:
        sys.exit(f"No '{NS_NAME}' found in {FOLDER}. Rename the NetSuite export to '{NS_NAME}' and re-run.")
    grn_files = [f for f in files if f != ns_path]
    if not grn_files:
        sys.exit(f"No GRN files in {FOLDER} - add one Product Wise Purchase file per hub, named like OMR.xlsx")
    print(f"Folder    : {FOLDER}")
    print(f"NS export : {os.path.basename(ns_path)}")
    print(f"GRN hubs  : {[os.path.basename(f) for f in grn_files]}")

    # --- consolidate GRN (EGIR only) ---
    grn = pd.concat([load_grn(f) for f in grn_files], ignore_index=True)

    grn_inv = (grn.groupby("Supplier Invoice No")
                  .agg(GRN_Hub=("Hub", lambda s: "/".join(sorted(set(s)))),
                       GRN_Lines=("Taxable", "size"),
                       GRN_Taxable=("Taxable", "sum"))
                  .reset_index().rename(columns={"Supplier Invoice No": "Invoice"}))
    grn_inv["GRN_Taxable"] = grn_inv["GRN_Taxable"].round(2)
    grn_inv = grn_inv[grn_inv["GRN_Taxable"] != 0]                # ignore zero-taxable

    # --- NS ---
    ns = pd.read_excel(ns_path, header=0, dtype=str)
    ns.columns = [str(c).strip() for c in ns.columns]
    loc_col = 'BRANDS / PROJECTS: filter by "Location"'
    ns["amt"] = pd.to_numeric(ns["Amount"], errors="coerce").fillna(0)
    ns_inv = (ns.groupby("Document Number")
                .agg(NS_Location=(loc_col, lambda s: "/".join(sorted(set(s.dropna().astype(str))))),
                     NS_Taxable=("amt", "sum"))
                .reset_index().rename(columns={"Document Number": "Invoice"}))
    ns_inv["NS_Taxable"] = ns_inv["NS_Taxable"].round(2)
    ns_inv = ns_inv[ns_inv["NS_Taxable"] != 0]                   # ignore zero-taxable

    # --- reconcile ---
    rec = pd.merge(grn_inv, ns_inv, on="Invoice", how="outer", indicator=True)
    rec["GRN_Taxable"] = rec["GRN_Taxable"].fillna(0)
    rec["NS_Taxable"] = rec["NS_Taxable"].fillna(0)
    rec["Diff"] = (rec["GRN_Taxable"] - rec["NS_Taxable"]).round(2)

    def status(r):
        if r["_merge"] == "left_only":
            return "GRN_ONLY (not in NS)"
        if r["_merge"] == "right_only":
            return "NS_ONLY (not in GRN)"
        if r["GRN_Hub"] != r["NS_Location"]:
            return "HUB_MISMATCH"           # booked under another hub's vendor/location
        return "MATCH" if abs(r["Diff"]) < TOL else "MISMATCH"

    rec["Status"] = rec.apply(status, axis=1)
    rec = rec.drop(columns="_merge").sort_values(["Status", "Invoice"]).reset_index(drop=True)
    rec.insert(0, "Hub", rec["GRN_Hub"].fillna(rec["NS_Location"]))

    # --- what happened to each GRN_ONLY invoice (booked lane + Missing PDFs list) ---
    booked, chase = load_booked(), mt.load()

    def track(r):
        inv = str(r["Invoice"]).strip()
        if not r["Status"].startswith("GRN_ONLY"):
            return ""
        if inv in booked:
            return "booked - not in this NS export yet"
        if inv in chase:
            return "Missing PDFs list: " + chase[inv]["status"]
        return NOT_CHASED

    rec.insert(rec.columns.get_loc("Status") + 1, "Track", rec.apply(track, axis=1))

    # --- summary ---
    summ = rec["Status"].value_counts().rename_axis("Status").reset_index(name="Invoices")
    grn_hubs = set(grn["Hub"])
    hubs = hub_summary(rec, grn_hubs)

    out = os.path.join(FOLDER, "Reconciliation_result.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xl:
        hubs.to_excel(xl, sheet_name="Hub_Summary", index=False)
        rec.to_excel(xl, sheet_name="Reconciliation", index=False)
        summ.to_excel(xl, sheet_name="Summary", index=False)
        grn.to_excel(xl, sheet_name="GRN_Consolidated_EGIR", index=False)
        write_hub_sheets(xl, hubs, rec)

    print("\n=== SUMMARY ===")
    print(summ.to_string(index=False))
    print(f"\nGRN EGIR invoices: {len(grn_inv)} | NS invoices: {len(ns_inv)}")
    print("\n=== PER HUB ===")
    print(hubs[["Hub", "GRN file", "GRN invoices", "Match", "Mismatch (amount/hub)", "GRN only",
                "..NOT booked, NOT chased", "NS only", "GRN - NS"]].to_string(index=False))
    missing = [h for h in HUB_NAMES if h not in grn_hubs]
    if missing:
        print(f"\nNo GRN file for: {', '.join(missing)}  (their NetSuite bills show as NS_ONLY)")
    nc = rec[rec["Track"] == NOT_CHASED]
    if len(nc):
        print(f"\n!! {len(nc)} invoice(s) received but NOT booked and NOT on the Missing PDFs list "
              f"(taxable {nc['GRN_Taxable'].sum():,.2f}) - see Track column")
    print(f"Wrote: {out}")
    bad = rec[rec["Status"].isin(["MISMATCH"])]
    if len(bad):
        print("\nMISMATCHES:")
        print(bad[["Invoice", "GRN_Hub", "NS_Location", "GRN_Taxable", "NS_Taxable", "Diff"]].to_string(index=False))


if __name__ == "__main__":
    main()

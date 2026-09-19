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
import sys, glob, os
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


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
    df["Hub"] = os.path.splitext(os.path.basename(path))[0]
    return df


def main():
    files = [f for f in glob.glob(os.path.join(FOLDER, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")
             and os.path.basename(f) != OUT_NAME]
    ns_path = next((f for f in files if os.path.basename(f).lower() == NS_NAME.lower()), None)
    if ns_path is None:
        sys.exit(f"No '{NS_NAME}' found in {FOLDER}. Rename the NetSuite export to '{NS_NAME}' and re-run.")
    grn_files = [f for f in files if f != ns_path]
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
        return "MATCH" if abs(r["Diff"]) < TOL else "MISMATCH"

    rec["Status"] = rec.apply(status, axis=1)
    rec = rec.drop(columns="_merge").sort_values(["Status", "Invoice"]).reset_index(drop=True)

    # --- summary ---
    summ = rec["Status"].value_counts().rename_axis("Status").reset_index(name="Invoices")

    out = os.path.join(FOLDER, "Reconciliation_result.xlsx")
    with pd.ExcelWriter(out) as xl:
        rec.to_excel(xl, sheet_name="Reconciliation", index=False)
        summ.to_excel(xl, sheet_name="Summary", index=False)
        grn.to_excel(xl, sheet_name="GRN_Consolidated_EGIR", index=False)

    print("\n=== SUMMARY ===")
    print(summ.to_string(index=False))
    print(f"\nGRN EGIR invoices: {len(grn_inv)} | NS invoices: {len(ns_inv)}")
    print(f"Wrote: {out}")
    bad = rec[rec["Status"].isin(["MISMATCH"])]
    if len(bad):
        print("\nMISMATCHES:")
        print(bad[["Invoice", "GRN_Hub", "NS_Location", "GRN_Taxable", "NS_Taxable", "Diff"]].to_string(index=False))


if __name__ == "__main__":
    main()

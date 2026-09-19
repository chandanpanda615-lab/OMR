"""Build ONE Master_Entry.xlsx covering every hub in a single Entry sheet.

Usage:
    python build_master_workbook.py
    python build_master_workbook.py "C:\\path\\to\\grn_copy.csv"

The Hub column (last column) is auto-filled from the CSV's fc_name, so you never type it;
a dropdown is added only so a mis-derived hub can be corrected. Fill Amount_5% / Amount_18%
(and IGST? for the rare inter-state row) for all invoices in this one file, then run
generate_master_csvs.py to produce the combined NetSuite import CSVs.

Output: <OUT_ROOT>\\Master_Entry.xlsx   (only HUL / HUL SAMADHAN rows with an S3 link)
"""
import csv, os, sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation
from build_entry_workbook import (HUBS, BRANDS, DEFAULT_CSV, OUT_ROOT,
                                  ENTRY_HEADERS, ENTRY_WIDTHS, add_entry_row,
                                  to_ddmmyyyy, BOLD)
from mark_booked import load_booked   # _booked.csv is the single "already done" ledger

LINK = Font(color="0563C1", underline="single")   # the Attached file cell doubles as the PDF link


def add_about(wb, csv_path, total, already, missing, per_hub):
    """Add an 'About' tab recording when/from-what this file was built, so an old
    Master_Entry/Master_Links opened later can explain itself. Does not touch the
    data sheet, so Entry<->Links row sync is unaffected."""
    ws = wb.create_sheet("About")
    hubs = ", ".join(f"{loc} {per_hub[loc]}" for loc in sorted(per_hub))
    rows = [("Built", date.today().isoformat()),
            ("Source CSV", csv_path),
            ("New invoices", total),
            ("Already booked (skipped)", already),
            ("Missing PDF link", missing),
            ("Hubs", hubs)]
    for label, value in rows:
        ws.append([label, value])
        ws.cell(ws.max_row, 1).font = BOLD
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 80


def main():
    csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV
    loc_by_key = {k: HUBS[k]["Location"] for k in HUBS}   # fc_name(lower) -> display Location
    booked = load_booked()   # invoices already imported to NetSuite -> never list again

    wb = Workbook()
    ent = wb.active
    ent.title = "Entry"
    headers = ENTRY_HEADERS + ["Hub", "Remark"]
    ent.append(headers)
    for c in range(1, len(headers) + 1):
        ent.cell(1, c).font = BOLD
    hub_col = len(ENTRY_HEADERS) + 1        # Hub right after N (A..N formulas stay stable)
    # Remark (Hub+1): leave blank; type a reason (e.g. "wrong file") on a row you cannot book.
    # mark_booked.py collects every not-booked row + its remark into MISSING_PDFS.xlsx.

    per_hub, already, missing = {}, 0, 0
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if (r.get("brand_name") or "").strip().upper() not in BRANDS:
                continue
            key = (r.get("fc_name") or "").strip().lower()
            if key not in HUBS:
                continue
            inv = (r.get("invoice_no") or "").strip()
            if inv in booked:            # old link in the GRN -> already booked, skip
                already += 1
                continue
            url = (r.get("s3_file_link") or "").strip()
            if not url.startswith("https://cdms-signed-invoice.s3"):
                missing += 1             # new invoice but PDF link blank -> chase it
                continue
            row = add_entry_row(ent,
                                inv,
                                to_ddmmyyyy(r.get("invoice_date")),
                                (r.get("brand_name") or "").strip(),
                                os.path.basename(url),
                                igst_editable=True)   # any row may be flagged IGST
            ent.cell(row, hub_col).value = loc_by_key[key]
            af = ent.cell(row, 4); af.hyperlink = url; af.font = LINK   # click the filename to open the PDF
            per_hub[loc_by_key[key]] = per_hub.get(loc_by_key[key], 0) + 1

    if not per_hub:
        sys.exit(f"Nothing new to enter in {csv_path}  "
                 f"(already booked: {already}, missing PDF link: {missing})")
    total = sum(per_hub.values())

    hub_letter = ent.cell(1, hub_col).column_letter
    dv = DataValidation(type="list",
                        formula1='"%s"' % ",".join(loc_by_key[k] for k in HUBS),
                        allow_blank=False)
    ent.add_data_validation(dv)
    dv.add(f"{hub_letter}2:{hub_letter}{ent.max_row}")

    for i, w in enumerate(ENTRY_WIDTHS + [16, 30], start=1):   # Hub 16, Remark 30
        ent.column_dimensions[ent.cell(1, i).column_letter].width = w
    ent.freeze_panes = "E2"

    os.makedirs(OUT_ROOT, exist_ok=True)
    add_about(wb, csv_path, total, already, missing, per_hub)
    out = os.path.join(OUT_ROOT, "Master_Entry.xlsx")
    wb.save(out)

    print(f"Built: {out}  (click the Attached file cell to open its PDF)")
    print(f"NEW to enter: {total} across {len(per_hub)} hubs   "
          f"(already booked, skipped: {already} | new but PDF link missing: {missing})")
    for loc in sorted(per_hub):
        print(f"  {loc}: {per_hub[loc]}")
    print("Fill YELLOW cells (Amount_5%, Amount_18%, IGST? rare), save, then run generate_master_csvs.py")


if __name__ == "__main__":
    main()

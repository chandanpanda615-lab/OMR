"""Build ONE Master_Entry.xlsx covering every hub in a single Entry sheet.

Usage:
    python build_master_workbook.py
    python build_master_workbook.py "C:\\path\\to\\grn_copy.csv"

The Hub column (last column) is auto-filled from the CSV's fc_name, so you never type it;
a dropdown is added only so a mis-derived hub can be corrected. Fill Amount_5% / Amount_18%
(and IGST? for the rare inter-state row) for all invoices in this one file, then run
generate_master_csvs.py to produce the combined NetSuite import CSVs.

Also lists CDMS_Recovered.csv (PDFs CDMS_Tool\\download_invoices.py found for the chase list),
so recovered invoices are booked in the same batch. Chase-list rules:
  - hub not in HUBS/ALIASES   -> printed + put on the chase list (never silently skipped)
  - same PDF that was flagged wrong before -> not listed again
  - a NEW PDF for an invoice that was wrong -> listed with Remark "VERIFY" (type amounts to book)
Rows whose PDF Gemini already read (5_NetSuite_Booking\\Gemini\\raw) are pre-filled from that saved
answer - free, no API call. Columns after N: Hub, Remark, No Gemini (type x), Printed Tax (Gemini),
Gemini Check, Gemini note.

Output: <OUT_ROOT>\\Master_Entry.xlsx   (only HUL / HUL SAMADHAN rows with an S3 link)
"""
import os, sys
from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation
from build_entry_workbook import (HUBS, BRANDS, DEFAULT_CSV, OUT_ROOT, hub_key,
                                  ENTRY_HEADERS, ENTRY_WIDTHS, add_entry_row,
                                  to_ddmmyyyy, BOLD, MASTER_EXTRA, MASTER_EXTRA_WIDTHS)
from fill_tax_from_pdfs import fill_sheet   # re-apply SAVED Gemini answers (free, no API call)
from mark_booked import load_booked   # _booked.csv is the single "already done" ledger
import missing_tracker as mt

LINK = Font(color="0563C1", underline="single")   # the Attached file cell doubles as the PDF link


def add_about(wb, csv_path, total, counts, per_hub):
    """Add an 'About' tab recording when/from-what this file was built, so an old
    Master_Entry opened later can explain itself. Does not touch the data sheet."""
    ws = wb.create_sheet("About")
    hubs = ", ".join(f"{loc} {per_hub[loc]}" for loc in sorted(per_hub))
    rows = [("Built", date.today().isoformat()),
            ("Source CSV", csv_path),
            ("New invoices", total),
            ("  of which from CDMS recovery", counts["recovered"]),
            ("Already booked (skipped)", counts["already"]),
            ("Missing PDF link", counts["missing"]),
            ("Known wrong file (skipped)", counts["wrong"]),
            ("Unknown hub (skipped, on chase list)", counts["unknown"]),
            ("Hubs", hubs)]
    for label, value in rows:
        ws.append([label, value])
        ws.cell(ws.max_row, 1).font = BOLD
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 80


def refuse_if_filled(path):
    """Never overwrite a Master_Entry that already has amounts typed in (unfinished batch)."""
    if not os.path.exists(path):
        return
    for r in load_workbook(path, read_only=True)["Entry"].iter_rows(min_row=2, max_col=6, values_only=True):
        if r[0] is not None and (r[4] not in (None, "") or r[5] not in (None, "")):
            sys.exit(f"{path} already has amounts filled - finish that batch first "
                     f"(generate_master_csvs -> import -> mark_booked -> archive_batch), or move it away.")


def main():
    csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV
    out = os.path.join(OUT_ROOT, "Master_Entry.xlsx")
    refuse_if_filled(out)
    mt.assert_writable(mt.XLSX, out)
    booked = load_booked()   # invoices already imported to NetSuite -> never list again
    chase = mt.load()

    wb = Workbook()
    ent = wb.active
    ent.title = "Entry"
    headers = ENTRY_HEADERS + MASTER_EXTRA
    ent.append(headers)
    for c in range(1, len(headers) + 1):
        ent.cell(1, c).font = BOLD
    hub_col = len(ENTRY_HEADERS) + 1        # Hub right after N (A..N formulas stay stable)
    # Remark (Hub+1): type a reason (e.g. "wrong file") on a row you cannot book.
    # mark_booked.py puts every not-booked row + its remark on the chase list.

    per_hub, unknown_fc = {}, {}
    counts = dict(already=0, missing=0, wrong=0, unknown=0, recovered=0)
    for r in mt.grn_rows(csv_path):
        if (r.get("brand_name") or "").strip().upper() not in BRANDS:
            continue
        inv = (r.get("invoice_no") or "").strip()
        if inv in booked:                # old link in the GRN -> already booked, skip
            counts["already"] += 1
            continue
        key = hub_key(r.get("fc_name"))
        if key not in HUBS:              # new spelling / new hub -> loud, and chased
            counts["unknown"] += 1
            fc = (r.get("fc_name") or "").strip()
            unknown_fc[fc] = unknown_fc.get(fc, 0) + 1
            mt.upsert(chase, inv, hub=fc, brand=r.get("brand_name"), invoice_date=r.get("invoice_date"),
                      status=f"unknown hub '{fc}' - add it to HUBS or ALIASES in build_entry_workbook.py")
            continue
        if not mt.has_link(r):
            counts["missing"] += 1       # new invoice but PDF link blank -> download_hul_pdfs chased it
            continue
        fn = os.path.basename(r["s3_file_link"].strip())
        bad = chase.get(inv, {}).get("bad_file", "")
        if bad and bad == fn:            # the same PDF you already flagged as wrong
            counts["wrong"] += 1
            continue
        loc = HUBS[key]["Location"]
        row = add_entry_row(ent, inv, to_ddmmyyyy(r.get("invoice_date")),
                            (r.get("brand_name") or "").strip(), fn,
                            igst_editable=True)   # any row may be flagged IGST
        ent.cell(row, hub_col).value = loc
        af = ent.cell(row, 4); af.hyperlink = r["s3_file_link"].strip(); af.font = LINK   # click to open the PDF
        if bad:
            ent.cell(row, hub_col + 1).value = mt.VERIFY
        counts["recovered"] += r.get("_src") == "CDMS"
        per_hub[loc] = per_hub.get(loc, 0) + 1

    mt.save(chase)
    for fc, n in unknown_fc.items():
        print(f"UNKNOWN HUB: '{fc}' ({n} invoice(s)) - NOT listed; on the Missing PDFs list. "
              f"Add it to HUBS or ALIASES in build_entry_workbook.py, then re-run.")
    if not per_hub:
        sys.exit(f"Nothing new to enter in {csv_path}  "
                 f"(already booked: {counts['already']}, missing PDF link: {counts['missing']}, "
                 f"known wrong file: {counts['wrong']}, unknown hub: {counts['unknown']})")
    total = sum(per_hub.values())

    hub_letter = ent.cell(1, hub_col).column_letter
    dv = DataValidation(type="list",
                        formula1='"%s"' % ",".join(HUBS[k]["Location"] for k in HUBS),
                        allow_blank=False)
    ent.add_data_validation(dv)
    dv.add(f"{hub_letter}2:{hub_letter}{ent.max_row}")

    for i, w in enumerate(ENTRY_WIDTHS + MASTER_EXTRA_WIDTHS, start=1):
        ent.column_dimensions[ent.cell(1, i).column_letter].width = w
    ent.freeze_panes = "E2"
    g = fill_sheet(ent, OUT_ROOT)           # saved Gemini answers only - never a new (paid) call

    os.makedirs(OUT_ROOT, exist_ok=True)
    add_about(wb, csv_path, total, counts, per_hub)
    wb.save(out)

    print(f"STEP 2 - Master_Entry built: {total} invoice(s) to enter  -> {out}")
    print(f"  from today's GRN ................ {total - counts['recovered']}")
    print(f"  from CDMS recovery .............. {counts['recovered']}")
    print(f"  pre-filled from saved Gemini .... {g['passed']}  (+{g['review']} marked REVIEW, no cost)")
    print(f"  not listed - already booked ..... {counts['already']}")
    print(f"  not listed - no PDF link yet .... {counts['missing']}   (on the Missing PDFs list)")
    print(f"  not listed - same wrong PDF ..... {counts['wrong']}   (on the Missing PDFs list)")
    print(f"  not listed - unknown hub ........ {counts['unknown']}   (on the Missing PDFs list)")
    print("  per hub: " + ", ".join(f"{loc} {per_hub[loc]}" for loc in sorted(per_hub)))
    print("Next: fill the YELLOW cells yourself, and/or type x in 'No Gemini' for rows you will type,")
    print("      then  python fill_tax_from_pdfs.py  (Gemini fills the rest).  Then: python generate_master_csvs.py")


if __name__ == "__main__":
    main()

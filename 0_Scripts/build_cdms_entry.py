"""Build a SEPARATE Master-Entry batch for invoices just pulled from the CDMS portal.

The CDMS-recovered invoices are a different problem from the daily GRN flow, so they get
their own entry workbook that lives inside CDMS_Sorted and never mixes with the top-level
Master flow. Same columns as the normal Master_Entry (so booking works identically) + a
Remark column: fill amounts to book, or type a reason (e.g. "wrong file") to send the row
to MISSING_PDFS.xlsx -> Wrong_File_Chase later (see cdms_finalize.py).

    python build_cdms_entry.py            # reads Desktop\\CDMS_PDF_Result.xlsx + Desktop\\CDMS_Invoices

Writes into 5_NetSuite_Booking\\CDMS_Sorted\\:
    CDMS_Entry.xlsx   - one row per downloaded invoice (YELLOW = fill; Remark = flag)
    CDMS_Links.xlsx   - same rows, click-to-open link to each LOCAL pdf
And MOVES each downloaded PDF from the Desktop into CDMS_Sorted\\<Hub>\\PDFs\\<date>\\.
'PDF Not attached' invoices are skipped here (they stay in MISSING_PDFS.xlsx Missing_Links).
"""
import os, shutil, sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from build_entry_workbook import (HUBS, ENTRY_HEADERS, ENTRY_WIDTHS, add_entry_row, BOLD)

LINK = Font(color="0563C1", underline="single")   # the Attached file cell doubles as the PDF link

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
RESULT = os.path.join(DESKTOP, "CDMS_PDF_Result.xlsx")
SRC_PDFS = os.path.join(DESKTOP, "CDMS_Invoices")
OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
CDMS_ROOT = os.path.join(OUT_ROOT, "CDMS_Sorted")
MISSING_XLSX = os.path.join(OUT_ROOT, "MISSING_PDFS.xlsx")

VALID = {HUBS[k]["Location"] for k in HUBS}
FCMAP = {"Chromepet": "Chrompet"}   # CDMS portal name -> our hub folder name (rest match)
# Permanent public PDF base; the entry link uses the S3 URL (opens in the browser, not Adobe).
S3_BASE = "https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/"


def wrong_file_set():
    if not os.path.exists(MISSING_XLSX):
        return set()
    wb = load_workbook(MISSING_XLSX)
    if "Wrong_File_Chase" not in wb.sheetnames:
        return set()
    return {str(r[0]).strip() for r in wb["Wrong_File_Chase"].iter_rows(min_row=2, values_only=True) if r[0]}


def iso_to_ddmmyyyy(s):
    p = str(s)[:10].split("-")
    return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 else str(s)


def main():
    wrong = wrong_file_set()
    rows = list(load_workbook(RESULT)["PDF Result"].iter_rows(min_row=2, values_only=True))

    wb = Workbook(); ent = wb.active; ent.title = "Entry"
    headers = ENTRY_HEADERS + ["Hub", "Remark"]
    ent.append(headers)
    for c in range(1, len(headers) + 1):
        ent.cell(1, c).font = BOLD
    hub_col, rem_col = len(ENTRY_HEADERS) + 1, len(ENTRY_HEADERS) + 2

    per_hub, skipped, problems, flagged = {}, 0, [], 0
    for r in rows:
        inv = str(r[0]).strip() if r[0] else ""
        status = (r[5] or "")
        if not inv or not status.startswith("Downloaded"):
            skipped += 1                      # not attached / not found -> stays in Missing_Links
            continue
        hub = FCMAP.get((r[1] or "").strip(), (r[1] or "").strip())
        if hub not in VALID:
            problems.append(f"{inv}: CDMS FC '{r[1]}' -> '{hub}' not a known hub; row skipped")
            continue
        iso = str(r[4])[:10]
        fn = (r[6] or "").strip()
        # move the PDF off the Desktop into the quarantined CDMS_Sorted area
        dest_dir = os.path.join(CDMS_ROOT, hub, "PDFs", iso)
        dest = os.path.join(dest_dir, fn)
        src = os.path.join(SRC_PDFS, fn)
        if fn and os.path.exists(src) and not os.path.exists(dest):
            os.makedirs(dest_dir, exist_ok=True)
            shutil.move(src, dest)
        url = r[7].strip() if len(r) > 7 and r[7] else S3_BASE + fn   # exact URL if the report has it, else rebuild
        row = add_entry_row(ent, inv, iso_to_ddmmyyyy(r[4]), (r[2] or "").strip(), fn, igst_editable=True)
        ent.cell(row, hub_col).value = hub
        af = ent.cell(row, 4); af.hyperlink = url; af.font = LINK   # click the filename -> opens S3 PDF in browser
        if inv in wrong:
            ent.cell(row, rem_col).value = "was wrong before - VERIFY the new pdf before booking"
            flagged += 1
        per_hub[hub] = per_hub.get(hub, 0) + 1

    if not per_hub:
        sys.exit("Nothing downloaded to enter (no 'Downloaded' rows in CDMS_PDF_Result.xlsx).")

    for i, w in enumerate(ENTRY_WIDTHS + [16, 40], start=1):   # Hub 16, Remark 40
        ent.column_dimensions[ent.cell(1, i).column_letter].width = w
    ent.freeze_panes = "E2"
    os.makedirs(CDMS_ROOT, exist_ok=True)
    entry_out = os.path.join(CDMS_ROOT, "CDMS_Entry.xlsx")
    wb.save(entry_out)

    total = sum(per_hub.values())
    print(f"Built: {entry_out}  (click the Attached file cell -> opens the S3 PDF in the browser)")
    print(f"Rows to work: {total}   (skipped, not downloaded: {skipped})")
    for h in sorted(per_hub):
        print(f"  {h}: {per_hub[h]}")
    if flagged:
        print(f"Pre-flagged as VERIFY (were wrong before): {flagged}")
    for p in problems:
        print("PROBLEM:", p)
    print("Fill YELLOW amounts to book, or type a Remark (e.g. 'wrong file'), then run cdms_generate.py")


if __name__ == "__main__":
    main()

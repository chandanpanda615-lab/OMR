"""Close out the CDMS batch AFTER NetSuite has accepted the bills. This is the only
step that changes the ledger and the trackers, so run it only once booking succeeded.

    python cdms_mark_booked.py

Does, in order:
  1. reads CDMS_Sorted\\1_CDMS_Bills_Header.csv (what you imported) -> books those invoices
     in _booked.csv (deduped) and MOVES each PDF from CDMS_Sorted\\<hub>\\PDFs -> <hub>\\Booked
     (so booked CDMS invoices leave the quarantine and join the normal hubs).
  2. reads CDMS_Sorted\\CDMS_Entry.xlsx for rows you left with only a Remark ->
     upserts them into MISSING_PDFS.xlsx -> Wrong_File_Chase (DO NOT BOOK).
  3. removes every booked + flagged invoice from MISSING_PDFS.xlsx -> Missing_Links,
     and drops booked ones from Wrong_File_Chase.
If NetSuite rejected some, just don't include them (fix + re-run cdms_generate first);
nothing here touches a row until it is in the header CSV.
"""
import csv, os, shutil, sys
from datetime import date
from openpyxl import load_workbook
from build_entry_workbook import ENTRY_HEADERS
from mark_booked import OUT_ROOT, LEDGER, load_booked
import missing_tracker

CDMS_ROOT = os.path.join(OUT_ROOT, "CDMS_Sorted")
ENTRY = os.path.join(CDMS_ROOT, "CDMS_Entry.xlsx")
HEADER_CSV = os.path.join(CDMS_ROOT, "1_CDMS_Bills_Header.csv")
HUB_IDX, REM_IDX = len(ENTRY_HEADERS), len(ENTRY_HEADERS) + 1


def move_to_booked(hub, fn):
    """Move a booked PDF from CDMS_Sorted\\<hub>\\PDFs -> <hub>\\Booked."""
    if not (hub and fn):
        return False
    dest = os.path.join(OUT_ROOT, hub, "Booked", fn)
    if os.path.exists(dest):
        return True
    for dp, _, files in os.walk(os.path.join(CDMS_ROOT, hub, "PDFs")):
        if fn in files:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.move(os.path.join(dp, fn), dest)
            return True
    return False


def load_tab(name):
    if not os.path.exists(missing_tracker.XLSX):
        return None, []
    wb = load_workbook(missing_tracker.XLSX)
    if name not in wb.sheetnames:
        return None, []
    ws = wb[name]
    return [c.value for c in ws[1]], [list(r) for r in ws.iter_rows(min_row=2, values_only=True)
                                      if any(v is not None for v in r)]


def main():
    assert os.path.exists(HEADER_CSV), f"missing: {HEADER_CSV} - run cdms_generate.py first"
    booked = load_booked()
    today = date.today().isoformat()
    new_rows, already, moved, no_pdf, booked_now = [], 0, 0, 0, set()

    with open(HEADER_CSV, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            inv = str(row["External ID"]).strip()
            hub = str(row.get("Location", "")).strip()
            booked_now.add(inv)
            if move_to_booked(hub, str(row.get("Attached file", "")).strip()):
                moved += 1
            else:
                no_pdf += 1
            if inv in booked:
                already += 1
                continue
            booked.add(inv)
            new_rows.append([inv, hub, row.get("Invoice Date", ""), today])

    first_write = not os.path.exists(LEDGER)
    with open(LEDGER, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if first_write:
            w.writerow(["invoice_no", "hub", "invoice_date", "marked_on"])
        w.writerows(new_rows)

    # rows left with only a Remark (no amount) -> Wrong_File_Chase
    wrong_new = {}
    for r in load_workbook(ENTRY, data_only=True)["Entry"].iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        inv = str(r[0]).strip()
        remark = str(r[REM_IDX]).strip() if len(r) > REM_IDX and r[REM_IDX] else ""
        blank = (r[4] in (None, "")) and (r[5] in (None, ""))
        if inv not in booked_now and blank and remark:
            hub = str(r[HUB_IDX]).strip() if len(r) > HUB_IDX and r[HUB_IDX] else ""
            wrong_new[inv] = [inv, hub, r[1] if r[1] is not None else "", remark]

    # cascade the two trackers
    resolved = booked_now | set(wrong_new)
    ml_h, ml_rows = load_tab("Missing_Links")
    if ml_h:
        ic = ml_h.index("invoice_no")
        missing_tracker.write_sheet("Missing_Links", ml_h,
                                    [row for row in ml_rows if str(row[ic]).strip() not in resolved])
    wf_h, wf_rows = load_tab("Wrong_File_Chase")
    wf_h = wf_h or ["invoice_no", "hub", "invoice_date", "remark"]
    ic = wf_h.index("invoice_no")
    merged = {str(row[ic]).strip(): list(row) for row in wf_rows}
    for inv in booked_now:
        merged.pop(inv, None)
    merged.update(wrong_new)
    missing_tracker.write_sheet("Wrong_File_Chase", wf_h, list(merged.values()), widths=[14, 16, 12, 40])

    print(f"Booked now: {len(booked_now)}  (new to ledger: {len(new_rows)}, already: {already})")
    print(f"PDFs moved CDMS_Sorted -> Booked: {moved}   (not found: {no_pdf})")
    print(f"Wrong-file flagged -> Wrong_File_Chase: {len(wrong_new)}")
    print(f"Removed from Missing_Links: {len(resolved)}")
    print(f"Ledger: {LEDGER}  (total {len(booked)} invoices)")


if __name__ == "__main__":
    main()

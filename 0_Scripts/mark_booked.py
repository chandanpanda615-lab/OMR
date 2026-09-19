"""Record invoices as BOOKED after a successful NetSuite import, so they can never
be re-downloaded, re-listed, or re-booked by accident.

    python mark_booked.py                       # MASTER flow: reads 1_ALL_Bills_Header.csv
    python mark_booked.py "Mysore Road" Byrathi # per-hub flow: reads 1_<HUB>_Bills_Header.csv

Both modes append External IDs to 5_NetSuite_Booking\\_booked.csv
(invoice_no, hub, invoice_date, marked_on), deduped. build_master_workbook.py,
download_hul_pdfs.py and download_from_links.py all read this ledger and skip
anything already in it.

MASTER mode also MOVES each booked PDF from <hub>\\PDFs\\... to <hub>\\Booked\\...,
so the PDFs folder always shows only invoices still waiting to be booked.

Self-check: run it twice -> the second run reports every invoice as "already in
ledger" and adds 0 rows. That is the dedupe test, on real data.
"""
import csv, os, shutil, sys
from datetime import date

OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
LEDGER = os.path.join(OUT_ROOT, "_booked.csv")


def load_booked():
    booked = set()
    if os.path.exists(LEDGER):
        with open(LEDGER, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                booked.add(str(row["invoice_no"]).strip())
    return booked


def chase_rows(entry_path, booked):
    """Rows in Master_Entry that are NOT in the booked set (listed but skipped),
    each with its typed Remark -> the wrong-file / to-chase list. Deduped by invoice."""
    from openpyxl import load_workbook
    from build_entry_workbook import ENTRY_HEADERS
    hub_i, rem_i = len(ENTRY_HEADERS), len(ENTRY_HEADERS) + 1   # 0-based Hub, Remark columns
    rows, seen = [], set()
    for r in load_workbook(entry_path, data_only=True)["Entry"].iter_rows(min_row=2, values_only=True):
        inv = r[0]
        if inv is None:
            continue
        inv = str(inv).strip()
        if inv in booked or inv in seen:
            continue
        seen.add(inv)
        hub = str(r[hub_i]).strip() if len(r) > hub_i and r[hub_i] else ""
        remark = str(r[rem_i]).strip() if len(r) > rem_i and r[rem_i] else ""
        rows.append([inv, hub, r[1] if r[1] is not None else "", remark])
    return rows


def move_to_booked(hub, fn):
    """Move a booked PDF from <hub>\\PDFs\\... to <hub>\\Booked\\. Returns
    True if it now lives under Booked (moved or already there), False if not found."""
    if not (hub and fn):
        return False
    dest_dir = os.path.join(OUT_ROOT, hub, "Booked")
    dest = os.path.join(dest_dir, fn)
    if os.path.exists(dest):
        return True                      # already marked in a previous run
    for dp, _, files in os.walk(os.path.join(OUT_ROOT, hub, "PDFs")):
        if fn in files:
            os.makedirs(dest_dir, exist_ok=True)
            shutil.move(os.path.join(dp, fn), dest)
            return True
    return False


def main():
    hubs = [a for a in sys.argv[1:] if not a.startswith("-")]
    master = not hubs                    # no hub names -> master (ALL) flow
    booked = load_booked()
    new_rows, already, moved, no_pdf = [], 0, 0, 0
    today = date.today().isoformat()

    if master:
        p = os.path.join(OUT_ROOT, "1_ALL_Bills_Header.csv")
        assert os.path.exists(p), f"missing: {p} — run generate_master_csvs.py first"
        with open(p, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                inv = str(row["External ID"]).strip()
                if inv in booked:
                    already += 1
                    continue
                booked.add(inv)
                hub = str(row.get("Location", "")).strip()
                new_rows.append([inv, hub, row.get("Invoice Date", ""), today])
                if move_to_booked(hub, str(row.get("Attached file", "")).strip()):
                    moved += 1
                else:
                    no_pdf += 1
    else:
        for h in hubs:
            p = os.path.join(OUT_ROOT, h, f"1_{h}_Bills_Header.csv")
            assert os.path.exists(p), f"missing: {p} — generate its import CSV first"
            with open(p, newline="", encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    inv = str(row["External ID"]).strip()
                    if inv in booked:
                        already += 1
                        continue
                    booked.add(inv)
                    new_rows.append([inv, h, row.get("Invoice Date", ""), today])

    first_write = not os.path.exists(LEDGER)
    with open(LEDGER, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if first_write:
            w.writerow(["invoice_no", "hub", "invoice_date", "marked_on"])
        w.writerows(new_rows)
    print(f"Marked booked: {len(new_rows)} new  (already in ledger: {already})")
    if master:
        print(f"PDFs moved to Booked: {moved}   (PDF not found on disk: {no_pdf})")
    print(f"Ledger: {LEDGER}  (total {len(booked)} invoices)")

    if master:   # refresh the to-chase list from whatever is listed but still not booked
        entry_path = os.path.join(OUT_ROOT, "Master_Entry.xlsx")
        if os.path.exists(entry_path):
            from missing_tracker import write_sheet
            rows = chase_rows(entry_path, booked)
            write_sheet("Wrong_File_Chase",
                        ["invoice_no", "hub", "invoice_date", "remark"], rows,
                        widths=[14, 16, 12, 30])
            print(f"To chase (listed but not booked): {len(rows)} -> MISSING_PDFS.xlsx (Wrong_File_Chase tab)")


if __name__ == "__main__":
    main()

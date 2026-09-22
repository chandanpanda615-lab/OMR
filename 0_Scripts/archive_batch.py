"""Move a finished master batch into Archive\\Master_<date>\\ automatically.

    python archive_batch.py              # uses today's date
    python archive_batch.py 2026-09-19   # override the date

Moves the master/ALL files (+ CDMS_Recovered.csv) out of 5_NetSuite_Booking\\ into the dated folder.
_booked.csv (the permanent ledger) and each hub's Booked\\ PDFs stay put.
A 2nd batch on the same day goes to Master_<date>_2 (the 1st is never overwritten).
"""
import os, shutil, sys
from datetime import date

OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
FILES = ["1_ALL_Bills_Header.csv", "2_ALL_Bills_Expenses.csv",
         "ALL_Tax_Verification.csv", "ALL_Invoices.zip",
         "Master_Entry.xlsx", "Master_Links.xlsx", "CDMS_Recovered.csv"]


def main():
    stamp = sys.argv[1] if len(sys.argv) > 1 else date.today().isoformat()
    dest = base = os.path.join(OUT_ROOT, "Archive", f"Master_{stamp}")
    n = 1
    while os.path.exists(dest):          # 2nd batch the same day -> Master_<date>_2, never overwrite the 1st
        n += 1
        dest = f"{base}_{n}"
    if not any(os.path.exists(os.path.join(OUT_ROOT, fn)) for fn in FILES):
        sys.exit("Nothing to archive.")
    os.makedirs(dest)
    moved, missing = 0, []
    for fn in FILES:
        src = os.path.join(OUT_ROOT, fn)
        if not os.path.exists(src):
            missing.append(fn)
            continue
        shutil.move(src, os.path.join(dest, fn))
        moved += 1
    print(f"Archived {moved} files -> {dest}")
    for fn in missing:
        print(f"  (not found, skipped): {fn}")


if __name__ == "__main__":
    main()

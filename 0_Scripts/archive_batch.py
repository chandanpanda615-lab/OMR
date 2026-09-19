"""Move a finished master batch into Archive\\Master_<date>\\ automatically.

    python archive_batch.py              # uses today's date
    python archive_batch.py 2026-09-19   # override the date

Moves the 6 master/ALL files out of 5_NetSuite_Booking\\ into the dated folder.
_booked.csv (the permanent ledger) and each hub's Booked\\ PDFs stay put.
Safe to re-run: a file already in the archive folder is replaced, not duplicated.
"""
import os, shutil, sys
from datetime import date

OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
FILES = ["1_ALL_Bills_Header.csv", "2_ALL_Bills_Expenses.csv",
         "ALL_Tax_Verification.csv", "ALL_Invoices.zip",
         "Master_Entry.xlsx", "Master_Links.xlsx"]


def main():
    stamp = sys.argv[1] if len(sys.argv) > 1 else date.today().isoformat()
    dest = os.path.join(OUT_ROOT, "Archive", f"Master_{stamp}")
    os.makedirs(dest, exist_ok=True)
    moved, missing = 0, []
    for fn in FILES:
        src = os.path.join(OUT_ROOT, fn)
        if not os.path.exists(src):
            missing.append(fn)
            continue
        dst = os.path.join(dest, fn)
        if os.path.exists(dst):
            os.remove(dst)            # re-run: replace the older copy, don't error
        shutil.move(src, dst)
        moved += 1
    print(f"Archived {moved} files -> {dest}")
    for fn in missing:
        print(f"  (not found, skipped): {fn}")


if __name__ == "__main__":
    main()

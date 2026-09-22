"""Move a finished CDMS recovery batch into CDMS_Sorted\\Archive\\CDMS_<date>\\ automatically,
so each CDMS run keeps its own clean record (just like archive_batch.py does for the daily flow).

    python archive_cdms.py              # uses today's date
    python archive_cdms.py 2026-09-19   # override the date

Moves the CDMS batch working files out of CDMS_Sorted\\. _booked.csv, MISSING_PDFS.xlsx, and
each hub's PDFs\\ / Booked\\ folders stay put. Safe to re-run (replaces, never errors).
"""
import os, shutil, sys
from datetime import date

OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
CDMS_ROOT = os.path.join(OUT_ROOT, "CDMS_Sorted")
FILES = ["CDMS_Entry.xlsx", "1_CDMS_Bills_Header.csv", "2_CDMS_Bills_Expenses.csv",
         "CDMS_Tax_Verification.csv", "CDMS_Invoices.zip"]


def main():
    stamp = sys.argv[1] if len(sys.argv) > 1 else date.today().isoformat()
    dest = os.path.join(CDMS_ROOT, "Archive", f"CDMS_{stamp}")
    os.makedirs(dest, exist_ok=True)
    moved, missing = 0, []
    for fn in FILES:
        src = os.path.join(CDMS_ROOT, fn)
        if not os.path.exists(src):
            missing.append(fn)
            continue
        dst = os.path.join(dest, fn)
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
        moved += 1
    print(f"Archived {moved} CDMS files -> {dest}")
    for fn in missing:
        print(f"  (not found, skipped): {fn}")


if __name__ == "__main__":
    main()

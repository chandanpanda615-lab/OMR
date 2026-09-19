"""Download HUL / HUL SAMADHAN signed invoice PDFs, saved by Location / invoice_date.

Usage:
    python download_hul_pdfs.py                       # uses the default CSV below
    python download_hul_pdfs.py "C:\\path\\to\\some.csv"

- Filters to brand_name in {HUL, HUL SAMADHAN} only.
- Folder layout: PDF_Downloads_HUL\\<Location>\\<invoice_date>\\<original>.pdf
- Rows with a blank S3 link are listed in the "Missing_Links" tab of MISSING_PDFS.xlsx
  but create NO folder, so nothing is silently dropped and no empty folders pile up.
- Re-run anytime: files already on disk are skipped, so it resumes cleanly.
"""
import csv, os, sys, urllib.request
from mark_booked import load_booked   # _booked.csv is the single "already done" ledger
from missing_tracker import write_sheet   # -> MISSING_PDFS.xlsx (Missing_Links tab)

DEFAULT_CSV = r"C:\Users\chandan.p\Downloads\grn_copy_upload_data (3).csv"
# Project root = parent of 0_Scripts; self-locating so the top folder can be renamed/moved.
OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")  # per-hub: <Location>\PDFs\<date>
BRANDS = {"HUL", "HUL SAMADHAN"}

csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV


def safe(name):
    # ponytail: strip only chars Windows forbids in folder names
    return "".join(c for c in (name or "Unknown").strip() if c not in r'<>:"/\|?*') or "Unknown"


def main():
    ok = skipped = failed = already = 0
    booked = load_booked()   # already imported to NetSuite -> never re-download
    missing = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if (row.get("brand_name") or "").strip().upper() not in BRANDS:
                continue
            if str(row.get("invoice_no") or "").strip() in booked:
                already += 1
                continue
            url = (row.get("s3_file_link") or "").strip()
            if not url.startswith("https://cdms-signed-invoice.s3"):
                missing.append(row)   # blank link -> log only, do NOT create an empty folder
                continue

            folder = os.path.join(OUT_ROOT, safe(row.get("fc_name")), "PDFs", safe(row.get("invoice_date")))
            os.makedirs(folder, exist_ok=True)  # created only once we have a real PDF to save
            target = os.path.join(folder, os.path.basename(url))
            if os.path.exists(target):
                skipped += 1
                continue
            try:
                urllib.request.urlretrieve(url, target)
                ok += 1
                print(f"[OK]  {row.get('fc_name')} / {row.get('invoice_date')} / {os.path.basename(url)}")
            except Exception as e:
                failed += 1
                print(f"[FAIL] {url}\n       {e}")

    # Log the blank-link invoices so the user can chase the missing PDFs from CDMS.
    miss_rows = [[r.get("fc_name"), r.get("brand_name"), r.get("invoice_no"),
                  r.get("invoice_date"), r.get("brand_grn_no"), "no S3 link"] for r in missing]
    miss_path = write_sheet("Missing_Links",
                            ["fc_name", "brand_name", "invoice_no", "invoice_date", "brand_grn_no", "remark"],
                            miss_rows, widths=[16, 14, 14, 12, 16, 14])

    print(f"\nDone. Downloaded={ok}  Skipped(existing)={skipped}  AlreadyBooked={already}  Failed={failed}  Missing-link={len(missing)}")
    print(f"Saved under: {OUT_ROOT}")
    print(f"Missing-PDF list: {miss_path}  (Missing_Links tab)")
    assert failed == 0, f"{failed} download(s) failed - see [FAIL] lines above"


if __name__ == "__main__":
    main()

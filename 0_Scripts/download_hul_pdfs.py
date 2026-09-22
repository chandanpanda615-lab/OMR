"""STEP 1 - download the day's HUL / HUL SAMADHAN signed invoice PDFs.

Usage:
    python download_hul_pdfs.py "C:\\path\\to\\grn.csv"

1. CDMS check first (only if CDMS_Tool\\token.txt is still valid): every invoice on the Missing
   PDFs list is looked up in CDMS; PDFs found go into CDMS_Recovered.csv and download below too.
2. Downloads every not-yet-booked PDF into 5_NetSuite_Booking\\<Hub>\\PDFs\\<yyyy-mm-dd>\\
   (fc_name spellings like "Chromepet" are resolved to the real hub via hub_key).
3. Invoices with no PDF link go on the Missing PDFs list (MISSING_PDFS.xlsx, tab Chase) and
   create NO folder. Re-run anytime: files already on disk are not downloaded again.
"""
import os, sys, urllib.request
from mark_booked import load_booked   # _booked.csv is the single "already done" ledger
from build_entry_workbook import HUBS, hub_key
import missing_tracker as mt           # Missing PDFs list + GRN/CDMS_Recovered rows

DEFAULT_CSV = r"C:\Users\chandan.p\Downloads\grn_copy_upload_data (3).csv"
# Project root = parent of 0_Scripts; self-locating so the top folder can be renamed/moved.
OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")  # per-hub: <Location>\PDFs\<date>
BRANDS = {"HUL", "HUL SAMADHAN"}
KINDS = ["new", "on_disk", "booked", "no_link", "failed"]


def safe(name):
    # ponytail: strip only chars Windows forbids in folder names
    return "".join(c for c in (name or "Unknown").strip() if c not in r'<>:"/\|?*') or "Unknown"


def location(fc_name):
    k = hub_key(fc_name)
    return HUBS[k]["Location"] if k in HUBS else safe(fc_name)


def cdms_check():
    """Look the Missing PDFs list up in CDMS before downloading. Skipped (not an error) when
    the token is expired - paste a fresh one into CDMS_Tool\\token.txt to include it."""
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "CDMS_Tool"))
    import download_invoices
    try:
        download_invoices.main([], verbose=False)
    except SystemExit as e:
        print(f"CDMS check SKIPPED: {e}")


def main():
    csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV
    mt.assert_writable()
    print(f"STEP 1 - download PDFs   (GRN: {os.path.basename(csv_path)})")
    cdms_check()
    booked = load_booked()   # already imported to NetSuite -> never re-download
    chase = mt.load()
    per_hub, missing, new_chase = {}, [], 0

    def count(hub, kind):
        per_hub.setdefault(hub, dict.fromkeys(KINDS, 0))[kind] += 1

    for row in mt.grn_rows(csv_path):
        if (row.get("brand_name") or "").strip().upper() not in BRANDS:
            continue
        hub = location(row.get("fc_name"))
        if str(row.get("invoice_no") or "").strip() in booked:
            count(hub, "booked")
            continue
        if not mt.has_link(row):
            count(hub, "no_link")
            missing.append(row)   # blank link -> Missing PDFs list only, do NOT create an empty folder
            continue
        url = row["s3_file_link"].strip()
        # date folder always yyyy-mm-dd, whether the CSV says 20/09/2026 or 2026-09-20
        folder = os.path.join(OUT_ROOT, hub, "PDFs", safe(mt.iso(row.get("invoice_date"))))
        target = os.path.join(folder, os.path.basename(url))
        if os.path.exists(target):
            count(hub, "on_disk")
            continue
        os.makedirs(folder, exist_ok=True)  # created only once we have a real PDF to save
        try:
            urllib.request.urlretrieve(url, target)
            count(hub, "new")
            print(f"  [OK]   {hub} / {os.path.basename(url)}" + ("   (from CDMS)" if row.get("_src") else ""))
        except Exception as e:
            count(hub, "failed")
            print(f"  [FAIL] {url}\n         {e}")

    for r in missing:
        inv = (r.get("invoice_no") or "").strip()
        if inv and inv not in chase:      # keep a more specific status (CDMS / wrong file) if already listed
            new_chase += 1
            mt.upsert(chase, inv, hub=location(r.get("fc_name")), brand=r.get("brand_name"),
                      invoice_date=r.get("invoice_date"), status="no S3 link (GRN)")
    path, open_n = mt.save(chase)

    tot = {k: sum(h[k] for h in per_hub.values()) for k in KINDS}
    print()
    print(f"  {'Hub':<14}{'Downloaded':>11}{'Already':>9}{'Booked':>8}{'No link':>9}{'Failed':>8}")
    print(f"  {'':<14}{'now':>11}{'on disk':>9}{'(skip)':>8}{'(list)':>9}{'':>8}")
    for hub in sorted(per_hub):
        h = per_hub[hub]
        print(f"  {hub:<14}{h['new']:>11}{h['on_disk']:>9}{h['booked']:>8}{h['no_link']:>9}{h['failed']:>8}")
    print(f"  {'TOTAL':<14}{tot['new']:>11}{tot['on_disk']:>9}{tot['booked']:>8}{tot['no_link']:>9}{tot['failed']:>8}")
    print("""
  Downloaded now  = new PDF saved today into <Hub>\\PDFs\\<date>\\
  Already on disk = PDF was downloaded in an earlier run - not downloaded again (nothing lost)
  Booked (skip)   = invoice is already in NetSuite (_booked.csv) - ignored
  No link (list)  = GRN has no PDF link yet - put on the Missing PDFs list""")
    print(f"Missing PDFs list: {open_n} open ({new_chase} new today) -> {path}")
    assert tot["failed"] == 0, f"{tot['failed']} download(s) failed - see [FAIL] lines above, then re-run"
    print(f'Next: python build_master_workbook.py "{csv_path}"')


if __name__ == "__main__":
    main()

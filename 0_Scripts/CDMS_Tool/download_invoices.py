r"""CDMS check of the Missing PDFs list (5_NetSuite_Booking\MISSING_PDFS.xlsx, tab Chase).

download_hul_pdfs.py runs this FIRST by itself (when token.txt is still valid), so normally you
never run it on its own. Stand-alone:
    python download_invoices.py                   # every open invoice on the list
    python download_invoices.py 9633128587 ...    # or only these invoice numbers

For each invoice it asks the CDMS portal for the signed-invoice PDF:
  - PDF found (and not the same file that was flagged wrong) -> written to
    5_NetSuite_Booking\CDMS_Recovered.csv; download_hul_pdfs.py + build_master_workbook.py read
    that file, so the invoice is downloaded and listed in the normal Master_Entry.xlsx.
  - PDF not attached / not found / same wrong file -> status updated; it stays on the list.
Token: paste it by hand into token.txt (HOW_TO_RUN.txt); an expired token stops this cleanly.
"""
import csv, json, os, sys, urllib.error, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))          # 0_Scripts
import missing_tracker as mt
from build_entry_workbook import HUBS, hub_key
from cdms_token import HEADERS, REFRESH, get_token

API = "https://api-cdms.ripplr.in/api/champ/grn/list"
FIELDS = ["fc_name", "brand_name", "invoice_no", "invoice_date", "brand_grn_no", "s3_file_link"]


def api_list(inv, token):
    body = json.dumps({
        "sort": {"sort_column": "createdAt", "sort_direction": "DESC"},
        "filter": {"user_fcs_only": True, "user_brands_only": True,
                   "invoice_no": inv, "invoice_pending": False},
        "page": {"limit": 10, "offset": 0},
    }).encode()
    req = urllib.request.Request(API, data=body, method="POST",
                                 headers={**HEADERS, "authorization": "Bearer " + token})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode())


def pick_pdf(row):
    files = row.get("GrnSignedInvoices") or []
    if not files:
        return None
    for f in files:
        if f.get("is_active"):
            return f
    for f in files:
        if f.get("is_verified"):
            return f
    return files[0]


def main(argv=None, verbose=True):
    """Returns how many PDFs were found. Exits (SystemExit) if the token is missing/expired."""
    argv = sys.argv[1:] if argv is None else argv
    mt.assert_writable()
    chase = mt.load()
    invoices = argv or list(chase)
    if not invoices:
        print("CDMS check: Missing PDFs list is empty - nothing to look up.")
        return 0
    token, valid = get_token()

    found, stats, stopped = [], {}, ""
    for i, inv in enumerate(invoices, 1):
        try:
            data = api_list(inv, token).get("data", {})
        except urllib.error.HTTPError as e:
            if e.code == 401:                       # token refused -> stop, keep what we have
                stopped = f"CDMS refused the token (401) after {i - 1} lookups - {REFRESH}"
                break
            data, err = None, f"CDMS: lookup error {e.code}"
        except Exception as e:                      # one bad lookup must not stop the rest
            data, err = None, f"CDMS: lookup error {e}"[:120]
        if data is None:
            status, kind = err, "error"
            mt.upsert(chase, inv, status=status)
        elif not data.get("count"):
            status, kind = "CDMS: not found in portal - check the invoice number", "not found"
            mt.upsert(chase, inv, status=status)
        else:
            r0 = data["rows"][0]
            fc = (r0.get("FC") or {}).get("name", "")
            k = hub_key(fc)
            info = dict(hub=HUBS[k]["Location"] if k in HUBS else fc,
                        brand=(r0.get("Brand") or {}).get("name", ""),
                        invoice_date=(r0.get("invoice_date") or "")[:10])
            url = (pick_pdf(r0) or {}).get("s3_file_link") or ""
            fn = os.path.basename(url)
            if not url:
                status, kind = "CDMS: PDF not attached yet - hub must upload it", "not attached"
            elif fn == chase.get(inv, {}).get("bad_file"):
                status, kind = "CDMS: still the same wrong file - ask hub to re-upload", "same wrong file"
            else:
                status, kind = "CDMS: PDF found - goes into the next Master_Entry", "found"
                found.append({"fc_name": fc, "brand_name": info["brand"], "invoice_no": inv,
                              "invoice_date": info["invoice_date"],
                              "brand_grn_no": r0.get("brand_grn_no", ""), "s3_file_link": url})
            mt.upsert(chase, inv, status=status, **info)
        stats[kind] = stats.get(kind, 0) + 1
        if verbose or kind == "found":
            print(f"  [{i}/{len(invoices)}] {inv}: {status}")

    if stopped and not stats:
        sys.exit(stopped)                           # refused at once = expired; change nothing
    with open(mt.RECOVERED, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(found)
    mt.save(chase)
    print(f"CDMS check (token {valid}): {len(invoices)} looked up -> "
          + ", ".join(f"{v} {k}" for k, v in stats.items()))
    if stopped:
        print("  " + stopped)
    return len(found)


if __name__ == "__main__":
    main()

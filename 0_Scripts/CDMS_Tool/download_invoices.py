r"""
CDMS invoice-PDF downloader (reusable).

HOW TO RUN each month:
  1. Put a FRESH login token in  token.txt   (see HOW_TO_RUN.txt - token expires every 24h)
  2. Put the invoice numbers (one per line) in  invoices.txt
  3. Double-click run.bat  OR  run:  python download_invoices.py

Output:
  - PDFs   -> Desktop\CDMS_Invoices
  - Report -> Desktop\CDMS_PDF_Result.xlsx  (status + file name for every invoice)
"""
import os, json, urllib.request, urllib.error
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(HERE, "token.txt")
LIST_FILE = os.path.join(HERE, "invoices.txt")
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
PDF_DIR = os.path.join(DESKTOP, "CDMS_Invoices")
OUT = os.path.join(DESKTOP, "CDMS_PDF_Result.xlsx")
API = "https://api-cdms.ripplr.in/api/champ/grn/list"

with open(TOKEN_FILE) as f:
    TOKEN = f.read().strip().replace("Bearer ", "").strip()
with open(LIST_FILE) as f:
    invoices = [x.strip() for x in f if x.strip()]

os.makedirs(PDF_DIR, exist_ok=True)
HEADERS = {
    "authorization": "Bearer " + TOKEN,
    "content-type": "application/json",
    "accept": "application/json, text/plain, */*",
    "origin": "https://cdms.ripplr.in",
    "referer": "https://cdms.ripplr.in/",
    "user-agent": "Mozilla/5.0",
}


def api_list(inv):
    body = json.dumps({
        "sort": {"sort_column": "createdAt", "sort_direction": "DESC"},
        "filter": {"user_fcs_only": True, "user_brands_only": True,
                   "invoice_no": inv, "invoice_pending": False},
        "page": {"limit": 10, "offset": 0},
    }).encode()
    req = urllib.request.Request(API, data=body, headers=HEADERS, method="POST")
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


def download(url, dest):
    req = urllib.request.Request(url, headers={"user-agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r, open(dest, "wb") as out:
        out.write(r.read())
    return os.path.getsize(dest)


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "PDF Result"
ws.append(["Invoice No", "FC / HUB", "Brand", "GRN No", "Invoice Date",
           "PDF Status", "PDF File Name", "S3 Link"])   # S3 Link = the permanent public PDF URL

stats = {"downloaded": 0, "no_pdf": 0, "not_found": 0, "error": 0}
for i, inv in enumerate(invoices, 1):
    fc = brand = grn = idate = fname = ""
    try:
        resp = api_list(inv)
    except urllib.error.HTTPError as e:
        if e.code == 401:
            print("AUTH FAILED (401) - token expired. Refresh token.txt and re-run.")
            ws.append([inv, "", "", "", "", "TOKEN EXPIRED - not processed", ""])
            wb.save(OUT)
            raise SystemExit(2)
        stats["error"] += 1
        ws.append([inv, fc, brand, grn, idate, "API error %d" % e.code, fname])
        print("[%d/%d] %s: API ERROR %d" % (i, len(invoices), inv, e.code))
        continue
    except Exception as e:
        stats["error"] += 1
        ws.append([inv, fc, brand, grn, idate, "Error: %s" % e, fname])
        print("[%d/%d] %s: ERROR %s" % (i, len(invoices), inv, e))
        continue

    data = resp.get("data", {})
    if data.get("count", 0) == 0:
        stats["not_found"] += 1
        ws.append([inv, fc, brand, grn, idate, "Not found in portal", fname])
        print("[%d/%d] %s: NOT FOUND" % (i, len(invoices), inv))
        continue

    row0 = data["rows"][0]
    fc = (row0.get("FC") or {}).get("name", "")
    brand = (row0.get("Brand") or {}).get("name", "")
    grn = row0.get("brand_grn_no", "")
    idate = (row0.get("invoice_date") or "")[:10]

    pdf = pick_pdf(row0)
    if not pdf:
        stats["no_pdf"] += 1
        ws.append([inv, fc, brand, grn, idate, "PDF Not attached", fname])
        print("[%d/%d] %s: PDF Not attached" % (i, len(invoices), inv))
        continue

    fname = os.path.basename(pdf.get("file_name") or (inv + ".pdf"))
    dest = os.path.join(PDF_DIR, fname)
    if os.path.exists(dest):
        base, ext = os.path.splitext(fname)
        k = 1
        while os.path.exists(os.path.join(PDF_DIR, "%s(%d)%s" % (base, k, ext))):
            k += 1
        dest = os.path.join(PDF_DIR, "%s(%d)%s" % (base, k, ext))
        fname = os.path.basename(dest)
    try:
        sz = download(pdf["s3_file_link"], dest)
        stats["downloaded"] += 1
        ws.append([inv, fc, brand, grn, idate, "Downloaded", fname, pdf.get("s3_file_link", "")])
        print("[%d/%d] %s: OK -> %s (%d bytes)" % (i, len(invoices), inv, fname, sz))
    except Exception as e:
        stats["error"] += 1
        ws.append([inv, fc, brand, grn, idate, "Download error: %s" % e, fname])
        print("[%d/%d] %s: DOWNLOAD ERROR %s" % (i, len(invoices), inv, e))

for col, w in zip("ABCDEFGH", [14, 12, 14, 12, 12, 22, 55, 70]):
    ws.column_dimensions[col].width = w
wb.save(OUT)
print("---")
print("TOTAL:", len(invoices), "STATS:", stats)
print("Report:", OUT)
print("PDFs  :", PDF_DIR)

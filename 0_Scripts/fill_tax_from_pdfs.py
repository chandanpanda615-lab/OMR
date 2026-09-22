"""OPTION B: auto-fill Master_Entry.xlsx tax cells from the PDFs, using Gemini.

Flow:  download_hul_pdfs -> build_master_workbook -> Master_Entry.xlsx -> [fill] -> generate_master_csvs
At [fill] you decide per row, BEFORE running this script:
    type the amounts yourself          -> row left alone (your numbers win), no Gemini call
    type x in "No Gemini (type x)"     -> row left alone, no Gemini call
    leave both blank                   -> this script reads it

Every Gemini answer is SAVED FIRST in 5_NetSuite_Booking\\Gemini\\raw\\<md5>.json (one file per
PDF, never overwritten) and the Master is then filled FROM those saved answers. So:
    - the same PDF is never paid for twice (a rebuild or re-run reuses the saved answer, free);
    - build_master_workbook re-applies saved answers whenever it builds a Master;
    - Gemini\\Gemini_Results.xlsx lists every answer next to what was finally put in the import
      file, so you can see how often Gemini was right (refreshed here and by generate_master_csvs).

Gemini only READS: 5% taxable, 18% taxable, IGST yes/no, printed invoice no (on 2 pages),
invoice date, printed total tax. All tax math stays in Python.
SAFETY: a row is filled ONLY if the computed tax matches the printed total tax (within Rs 1).
Otherwise it stays blank, marked red REVIEW, with the printed tax kept in "Printed Tax (Gemini)" -
and generate_master_csvs refuses a row whose typed amounts don't match that printed tax.

Usage:
    python fill_tax_from_pdfs.py             # needs GEMINI_API_KEY in OMR\\.env (else: saved answers only)
    python fill_tax_from_pdfs.py --view      # only refresh Gemini\\Gemini_Results.xlsx
    python fill_tax_from_pdfs.py --selftest  # offline logic check, no API, no files
"""
import glob
import hashlib
import json
import os
from datetime import datetime
import sys
import time
import urllib.request
from build_entry_workbook import NO_GEMINI, PRINTED_TAX, GEMINI_CHECK, GEMINI_NOTE

# ---- pricing / tuning ----
MODEL = "gemini-2.5-flash"
IN_PRICE = 0.30 / 1_000_000
OUT_PRICE = 2.50 / 1_000_000   # output AND thinking tokens
USD_INR = 84.0
THINKING_BUDGET = 512          # 0 mis-reads the slab split; 512 reads it (verified on real HUL scans)
CHECKSUM_TOLERANCE = 1.0       # rupees; absorbs per-line GST rounding
RETRIES = 3

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BOOK_ROOT = os.path.join(os.path.dirname(SCRIPT_DIR), "5_NetSuite_Booking")
MASTER = os.path.join(BOOK_ROOT, "Master_Entry.xlsx")
STORE = os.path.join(BOOK_ROOT, "Gemini")
RAW = os.path.join(STORE, "raw")                    # <md5>.json = what Gemini said about that PDF
VIEW = os.path.join(STORE, "Gemini_Results.xlsx")

COL_DATE, COL_FILE, COL_A5, COL_A18, COL_IGST, COL_HUB = 2, 4, 5, 6, 7, 15

PROMPT = (
    "This is a Hindustan Unilever (HUL) GST tax invoice document, possibly scanned and "
    "rotated/upside-down, and it may bundle SEVERAL separate tax invoices plus logistics "
    "pages (lorry receipt / loading slip). Identify EVERY distinct tax invoice in the file. "
    "For each invoice use ONLY these TWO pages and ignore every other page: "
    "(A) the FIRST page of the invoice, titled 'TAX INVOICE', which prints 'Invoice Number' at the "
    "top and 'Invoice Date' in the right-hand column; "
    "(B) the page that has the 'Tax Details' table (Description / Taxable Value INR / Tax Rate % / "
    "Tax on Net Taxable Turnover) and the totals box (Total Taxes, Invoice Amt); this page also "
    "prints 'Invoice Number' at the top. "
    "In the Tax Details table the GST slabs are printed as: 5% slab = CGST 2.5% + SGST 2.5% (or IGST 5%); "
    "18% slab = CGST 9% + SGST 9% (or IGST 18%). Amounts use Indian digit grouping (1,75,231.05 = 175231.05). "
    "For each invoice return (read printed values, do not compute): the Invoice Number from page A; "
    "the Invoice Date from page A exactly as printed (dd.mm.yyyy); the Invoice Number from page B; "
    "the total TAXABLE value (pre-tax base) of the 5% slab from page B (0 if none); the total TAXABLE "
    "value of the 18% slab from page B (0 if none); whether IGST is charged (yes/no); and the "
    "'Total Taxes' amount from page B (CGST+SGST, or IGST)."
)


# ---------- pure logic (unit-tested by --selftest, no API) ----------
def compute_tax(a5, a18):
    """Total GST from the two slab bases. Same total whether intra (2.5+2.5 / 9+9) or IGST (5 / 18)."""
    return round((a5 or 0) * 0.05 + (a18 or 0) * 0.18, 2)


def reconcile(a5, a18, printed_tax):
    """Return (status, note). status in {PASS, REVIEW}; PASS only if computed==printed within tol."""
    if (a5 or 0) == 0 and (a18 or 0) == 0:
        return "REVIEW", "no taxable amounts read"
    if not printed_tax or printed_tax <= 0:
        return "REVIEW", "no printed tax to verify against"
    diff = round(compute_tax(a5, a18) - printed_tax, 2)
    if abs(diff) <= CHECKSUM_TOLERANCE:
        return "PASS", ""
    return "REVIEW", f"checksum off by {diff:,.2f} (computed {compute_tax(a5, a18):,.2f} vs printed {printed_tax:,.2f})"


def to_ddmmyyyy(s):
    """'19.09.2026' (or 19/09/2026, 19-09-2026) -> '19/09/2026'; None if it is not a real date."""
    try:
        return datetime.strptime(str(s).strip().replace(".", "/").replace("-", "/"), "%d/%m/%Y").strftime("%d/%m/%Y")
    except ValueError:
        return None


def date_ok(d, fn):
    """Printed invoice date (dd/mm/yyyy) must be 0-10 days before the date in the PDF file name
    (catches misreads like 18/05 for 18/09). No date in the file name -> cannot check -> accept."""
    try:
        gap = (datetime.strptime(_iso_from(fn), "%Y-%m-%d") - datetime.strptime(d, "%d/%m/%Y")).days
    except ValueError:
        return True
    return 0 <= gap <= 10


def match_invoice(external_id, extracted):
    """Find the extracted invoice whose printed number == this row's External ID."""
    key = str(external_id).strip()
    for e in extracted:
        if str(e.get("external_id", "")).strip() == key:
            return e
    return None


def judge(ext, fn, invoices):
    """Decide one Master row from the invoices Gemini read in its PDF -> dict with
    status PASS/REVIEW, note, printed (tax), and a5/a18/igst/date when PASS."""
    m = match_invoice(ext, invoices)
    if m is None:
        found = ", ".join(sorted({str(e.get("external_id", "")).strip() for e in invoices} - {""})) or "none"
        return dict(status="REVIEW", printed=None,
                    note=f"invoice {ext} not printed in PDF (found: {found}) - wrong file?")
    if str(m.get("tax_page_invoice_no", "")).strip() != str(ext).strip():
        return dict(status="REVIEW", printed=None,
                    note=f"invoice no on Tax Details page ({m.get('tax_page_invoice_no')}) != {ext}")
    d = to_ddmmyyyy(m.get("invoice_date"))
    dnote = "" if d else "invoice date not read"
    if d and not date_ok(d, fn):
        d, dnote = None, f"date {d} looks wrong, sheet date kept"
    status, note = reconcile(m["amount_5"], m["amount_18"], m["printed_total_tax"])
    printed = round(m.get("printed_total_tax") or 0, 2) or None
    if status == "PASS":
        return dict(status="PASS", printed=printed, a5=round(m["amount_5"], 2), a18=round(m["amount_18"], 2),
                    igst=bool(m["is_igst"]), date=d,
                    note="; ".join(x for x in ("IGST" if m["is_igst"] else "", dnote) if x))
    return dict(status="REVIEW", printed=printed,
                note=f"{note}; read 5%={m['amount_5']:,.2f} 18%={m['amount_18']:,.2f}" + (f"; {dnote}" if dnote else ""))


def _selftest():
    assert compute_tax(100000, 0) == 5000.0
    assert compute_tax(0, 100000) == 18000.0
    assert reconcile(232635.50, 338581.45, 72576.34)[0] == "PASS"          # real BYTI row
    assert reconcile(151255.70, 0, 7562.84)[0] == "PASS"                    # real OMR row (pure 5%)
    assert reconcile(0, 505700.77, 72576.34)[0] == "REVIEW"                 # wrong split -> caught
    assert reconcile(0, 0, 9815.38)[0] == "REVIEW"                          # nothing read
    assert reconcile(100000, 0, 0)[0] == "REVIEW"                           # no printed tax
    ex = [{"external_id": "9629076836"}, {"external_id": "9629076837"}]
    assert match_invoice("9629076837", ex)["external_id"] == "9629076837"   # bundle match
    assert match_invoice(9629076838, ex) is None                            # not in bundle -> None
    assert to_ddmmyyyy("19.09.2026") == "19/09/2026" and to_ddmmyyyy("31.02.2026") is None and to_ddmmyyyy(None) is None
    f = "TLBL-HUL-19092026-9633128802-20260920100901.pdf"
    assert date_ok("18/09/2026", f) and not date_ok("18/05/2026", f) and not date_ok("20/09/2026", f)   # 18/05 = real misread
    inv = [{"external_id": "9633128802", "tax_page_invoice_no": "9633128802", "invoice_date": "18.09.2026",
            "amount_5": 151255.70, "amount_18": 0, "is_igst": False, "printed_total_tax": 7562.84}]
    j = judge("9633128802", f, inv)
    assert j["status"] == "PASS" and j["a5"] == 151255.70 and j["date"] == "18/09/2026" and j["printed"] == 7562.84
    assert judge("9633128803", f, inv)["note"].endswith("wrong file?")                    # not in this PDF
    inv[0]["amount_18"] = 1000
    j = judge("9633128802", f, inv)
    assert j["status"] == "REVIEW" and j["printed"] == 7562.84 and "a5" not in j          # checksum off -> not filled
    print("selftest OK")


# ---------- saved answers (Gemini\raw\<md5>.json) ----------
def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


def load_raw(md5):
    p = os.path.join(RAW, md5 + ".json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_raw(rec):
    """Written right after each paid call, before anything else can fail."""
    os.makedirs(RAW, exist_ok=True)
    p = os.path.join(RAW, rec["md5"] + ".json")
    with open(p + ".tmp", "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=1)
    os.replace(p + ".tmp", p)


# ---------- Gemini I/O ----------
def _extract(client, types, Extract, pdf_path):
    """One Gemini call -> list of invoice dicts + (in, out, thinking) tokens. Retries on transient errors."""
    last = None
    for attempt in range(1, RETRIES + 1):
        try:
            up = client.files.upload(file=pdf_path)
            r = client.models.generate_content(
                model=MODEL,
                contents=[up, PROMPT],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=list[Extract],
                    thinking_config=types.ThinkingConfig(thinking_budget=THINKING_BUDGET),
                ),
            )
            try:
                client.files.delete(name=up.name)
            except Exception:
                pass
            m = r.usage_metadata
            inv = [e.model_dump() for e in (r.parsed or [])]
            return inv, (m.prompt_token_count or 0, m.candidates_token_count or 0,
                         getattr(m, "thoughts_token_count", 0) or 0)
        except Exception as e:
            last = e
            print(f"    [retry {attempt}/{RETRIES}] {os.path.basename(pdf_path)}: {e}")
            time.sleep(2 * attempt)
    raise last


def make_caller(key):
    """-> call(path, md5, hub): asks Gemini about one PDF, SAVES the answer, returns it."""
    from google import genai
    from google.genai import types
    from pydantic import BaseModel, Field

    class Extract(BaseModel):
        external_id: str = Field(description="Invoice Number printed at the top of the TAX INVOICE first page (page A)")
        invoice_date: str = Field(description="Invoice Date from page A, exactly as printed, dd.mm.yyyy")
        tax_page_invoice_no: str = Field(description="Invoice Number printed at the top of the Tax Details page (page B)")
        amount_5: float = Field(description="Total TAXABLE value (pre-tax) of 5% GST items; 0 if none")
        amount_18: float = Field(description="Total TAXABLE value (pre-tax) of 18% GST items; 0 if none")
        is_igst: bool = Field(description="true if IGST charged (interstate), false if CGST+SGST")
        printed_total_tax: float = Field(description="Total tax printed on the invoice (CGST+SGST or IGST); verification only")

    client = genai.Client(api_key=key)

    def call(path, md5, hub):
        invoices, (ti, to, tt) = _extract(client, types, Extract, path)
        rec = {"md5": md5, "file": os.path.basename(path), "hub": hub,
               "read_on": datetime.now().strftime("%Y-%m-%d %H:%M"), "model": MODEL,
               "thinking_budget": THINKING_BUDGET, "tokens": {"in": ti, "out": to, "thinking": tt},
               "cost_rs": round((ti * IN_PRICE + (to + tt) * OUT_PRICE) * USD_INR, 2),
               "invoices": invoices}
        save_raw(rec)
        return rec
    return call


# ---------- fill a sheet from saved answers (used here AND by build_master_workbook) ----------
def _pdf_index(hubs, root):
    """Map every PDF basename under each present hub's PDFs folder (below `root`) -> full path."""
    idx = {}
    for loc in hubs:
        base = os.path.join(root, loc, "PDFs")
        if not os.path.isdir(base):
            continue
        for dp, _, files in os.walk(base):
            for fn in files:
                if fn.lower().endswith(".pdf"):
                    idx.setdefault(fn, os.path.join(dp, fn))
    return idx


def _iso_from(basename):
    """SUKA-HUL-17092026-...pdf -> 2026-09-17 (matches download_from_links folder convention)."""
    t = basename[:-4].split("-")
    if len(t) > 2 and len(t[2]) == 8 and t[2].isdigit():
        d = t[2]
        return f"{d[4:]}-{d[2:4]}-{d[0:2]}"
    return "_fetched"


def _download(url, hub, basename, root):
    """Download the PDF from its column-D link into the hub's PDFs folder under `root`
    (so the generate step's zip finds it too). Cached copies are reused, never re-fetched."""
    folder = os.path.join(root, hub or "_misc", "PDFs", _iso_from(basename))
    os.makedirs(folder, exist_ok=True)
    target = os.path.join(folder, basename)
    if not os.path.exists(target):
        urllib.request.urlretrieve(url, target)
    return target


def gemini_cols(ws):
    """Header name -> column number; adds the 4 Gemini columns if an older sheet lacks them."""
    col = {c.value: c.column for c in ws[1] if c.value}
    for name in (NO_GEMINI, PRINTED_TAX, GEMINI_CHECK, GEMINI_NOTE):
        if name not in col:
            col[name] = ws.max_column + 1
            ws.cell(1, col[name]).value = name
    return col


def fill_sheet(ws, root, call=None):
    """Fill every row that is not typed by you and not marked x, from SAVED Gemini answers.
    call=None (build_master_workbook, or no API key): saved answers only - never costs money.
    call given: a PDF with no saved answer is sent to Gemini once and the answer saved first.
    Returns a dict of counts."""
    from openpyxl.styles import PatternFill
    green, red = PatternFill("solid", fgColor="C6EFCE"), PatternFill("solid", fgColor="FFC7CE")
    col = gemini_cols(ws)
    n = dict(rows=0, marked=0, typed=0, saved=0, new_calls=0, cost_rs=0.0, passed=0, review=0, not_read=0)
    todo = []
    for row in range(2, ws.max_row + 1):
        ext, fn = ws.cell(row, 1).value, ws.cell(row, COL_FILE).value
        if ext is None or not fn:
            continue
        n["rows"] += 1
        if str(ws.cell(row, col[NO_GEMINI]).value or "").strip():
            n["marked"] += 1
            continue
        if ws.cell(row, COL_A5).value not in (None, "") or ws.cell(row, COL_A18).value not in (None, ""):
            n["typed"] += 1
            continue
        d = ws.cell(row, COL_FILE)
        todo.append((row, str(ext).strip(), str(fn).strip(), d.hyperlink.target if d.hyperlink else None,
                     str(ws.cell(row, COL_HUB).value or "").strip()))
    if not todo or (call is None and not glob.glob(os.path.join(RAW, "*.json"))):
        n["not_read"] = len(todo)
        return n

    def mark(row, printed, status, note):
        ws.cell(row, col[PRINTED_TAX]).value = printed or None
        c = ws.cell(row, col[GEMINI_CHECK])
        c.value, c.fill = status, (green if status == "PASS" else red)
        ws.cell(row, col[GEMINI_NOTE]).value = note or None

    cache = _pdf_index({t[4] for t in todo if t[4]}, root)
    groups = {}
    for row, ext, fn, url, hub in todo:
        path = cache.get(fn)
        if not path and url and call:
            try:
                path = cache[fn] = _download(url, hub, fn, root)
            except Exception as e:
                print(f"    [download failed] {fn}: {e}")
        if not path:
            if call:
                mark(row, None, "REVIEW", "PDF unavailable (no local copy and no/failed link in column D)")
                n["review"] += 1
            else:
                n["not_read"] += 1
            continue
        md5 = _md5(path)
        groups.setdefault(md5, {"path": path, "hub": hub, "rows": []})["rows"].append((row, ext, fn))

    for i, (md5, g) in enumerate(groups.items(), 1):
        rec = load_raw(md5)
        if rec is not None:
            n["saved"] += len(g["rows"])
        elif call is None:
            n["not_read"] += len(g["rows"])
            continue
        else:
            print(f"  [{i}/{len(groups)}] Gemini reads {os.path.basename(g['path'])}  ({len(g['rows'])} row(s))")
            try:
                rec = call(g["path"], md5, g["hub"])
            except Exception as e:
                for row, _, _ in g["rows"]:
                    mark(row, None, "REVIEW", f"Gemini failed: {e}")
                    n["review"] += 1
                continue
            n["new_calls"] += 1
            n["cost_rs"] += rec["cost_rs"]
        for row, ext, fn in g["rows"]:
            j = judge(ext, fn, rec["invoices"])
            if j["status"] == "PASS":
                ws.cell(row, COL_A5).value = j["a5"]
                ws.cell(row, COL_A18).value = j["a18"]
                ws.cell(row, COL_IGST).value = "y" if j["igst"] else None
                if j["date"]:
                    ws.cell(row, COL_DATE).value = j["date"]      # printed date beats the CSV date
                n["passed"] += 1
            else:
                n["review"] += 1
            mark(row, j["printed"], j["status"], j["note"])
    return n


# ---------- Gemini_Results.xlsx: every answer vs what was finally entered ----------
def write_view():
    """Rebuild Gemini\\Gemini_Results.xlsx from Gemini\\raw\\*.json and every *Tax_Verification.csv
    (what generate_master_csvs put in the import files). Read-only for everything else."""
    import csv
    from openpyxl import Workbook
    from openpyxl.styles import Font
    recs = []
    for p in sorted(glob.glob(os.path.join(RAW, "*.json"))):
        with open(p, encoding="utf-8") as f:
            recs.append(json.load(f))
    if not recs:
        return None
    entered = {}   # invoice -> (taxable 5%, taxable 18%) from the newest import file that has it
    for p in sorted(glob.glob(os.path.join(BOOK_ROOT, "**", "*Tax_Verification.csv"), recursive=True),
                    key=os.path.getmtime):
        with open(p, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                try:
                    entered[str(r["Invoice"]).strip()] = (float(r["Taxable5%"] or 0), float(r["Taxable18%"] or 0))
                except (KeyError, ValueError):
                    pass
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(["Read on", "Hub", "PDF file", "Invoice (printed)", "Gemini 5%", "Gemini 18%", "IGST",
               "Printed tax", "Gemini check", "Entered 5%", "Entered 18%", "Gemini vs entered"])
    s = dict(files=len(recs), cost=sum(r.get("cost_rs", 0) for r in recs), answers=0, passed=0, same=0, diff=0)
    for rec in sorted(recs, key=lambda r: r["read_on"]):
        for e in rec["invoices"] or [{}]:
            inv = str(e.get("external_id", "")).strip()
            chk = reconcile(e.get("amount_5"), e.get("amount_18"), e.get("printed_total_tax"))[0] if e else "nothing read"
            t = entered.get(inv)
            verdict = ""
            if t:
                same = abs(t[0] - (e.get("amount_5") or 0)) <= 1 and abs(t[1] - (e.get("amount_18") or 0)) <= 1
                verdict = "SAME" if same else "DIFFERENT"
                if chk == "PASS":
                    s["same" if same else "diff"] += 1
            s["answers"] += bool(e)
            s["passed"] += chk == "PASS"
            ws.append([rec["read_on"], rec.get("hub", ""), rec["file"], inv, e.get("amount_5"), e.get("amount_18"),
                       "y" if e.get("is_igst") else "", e.get("printed_total_tax"), chk,
                       t[0] if t else None, t[1] if t else None, verdict])
    for i, w in enumerate([16, 13, 52, 14, 13, 13, 5, 12, 12, 13, 13, 16], start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    sm = wb.create_sheet("Summary", 0)
    for label, v in [("PDFs read by Gemini (each paid once)", s["files"]),
                     ("Total cost so far (Rs, approx)", round(s["cost"], 2)),
                     ("Invoices read", s["answers"]),
                     ("PASS (tax math matched the printed tax)", s["passed"]),
                     ("PASS answers later in an import file", s["same"] + s["diff"]),
                     ("   ... same as Gemini", s["same"]),
                     ("   ... DIFFERENT (you changed it - check why)", s["diff"])]:
        sm.append([label, v])
        sm.cell(sm.max_row, 1).font = Font(bold=True)
    sm.column_dimensions["A"].width = 46
    try:
        wb.save(VIEW)
    except PermissionError:
        print(f"  (close {os.path.basename(VIEW)} in Excel to refresh it - the answers are safe in Gemini\\raw)")
        return None
    return VIEW


def main():
    if "--selftest" in sys.argv:
        return _selftest()
    if "--view" in sys.argv:
        print("Refreshed:", write_view() or "nothing saved yet")
        return

    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    master = args[0] if args else MASTER
    root = os.path.dirname(os.path.abspath(master))   # PDFs live under <root>\<Hub>\PDFs
    if not os.path.exists(master):
        sys.exit(f'No entry workbook: {master}\nBuild it first: python build_master_workbook.py "<csv>"')
    try:
        open(master, "r+b").close()
    except PermissionError:
        sys.exit(f"Close {os.path.basename(master)} in Excel first - nothing was sent to Gemini.")

    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(SCRIPT_DIR), ".env"))   # OMR\.env
    key = os.environ.get("GEMINI_API_KEY")
    call = make_caller(key) if key else None
    if not key:
        print("No GEMINI_API_KEY -> using SAVED Gemini answers only (no new calls). Fill the rest by hand.")

    from openpyxl import load_workbook
    wb = load_workbook(master)               # keep formulas (H..N), dropdown, hyperlinks intact
    n = fill_sheet(wb["Entry"], root, call)
    wb.save(master)
    view = write_view()

    print("-" * 64)
    print(f"Rows in Master_Entry ............... {n['rows']}")
    print(f"  marked x (No Gemini) - skipped ... {n['marked']}")
    print(f"  already filled (you/Gemini) - skip {n['typed']}")
    print(f"  from saved answers (free) ........ {n['saved']}")
    print(f"  new Gemini calls ................. {n['new_calls']} PDF(s), about Rs {n['cost_rs']:.2f}")
    print(f"  not read (no key / no PDF) ....... {n['not_read']}")
    print(f"Filled (checksum PASS) ............. {n['passed']}")
    print(f"Left for you (red REVIEW) .......... {n['review']}")
    print(f"Saved: {master}")
    if view:
        print(f"Gemini answers vs entered: {view}")
    print("Open it, fill any red REVIEW rows by hand (match 'Printed Tax (Gemini)'), then: python generate_master_csvs.py")


if __name__ == "__main__":
    main()

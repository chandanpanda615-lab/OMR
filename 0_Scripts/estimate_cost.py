"""Estimate Gemini 2.5 Flash cost for a folder of invoice PDFs BEFORE calling the API.

Pure local: counts pages, dedups by MD5, checks if scanned. Makes NO API call.
Usage:  python estimate_cost.py "C:\\path\\to\\folder"
        (no arg -> uses the Byrathi/Booked folder)
"""
import hashlib, sys
from pathlib import Path
import fitz  # pymupdf

# ---- 2.5 Flash rates (verified 2026-09-20). Change here if Google changes them. ----
IN_USD_PER_1M = 0.30      # input (document modality = 258 tokens/page)
OUT_USD_PER_1M = 2.50     # output
USD_TO_INR = 88.0
TOK_PER_PAGE = 258        # Gemini document tokenization
PROMPT_TOK = 245          # our fixed instruction text (measured)
OUT_TOK_PER_PAGE = 30     # ponytail: rough, from measured 601 out / 22 pages; output cost is tiny anyway

DEFAULT = r"C:\Users\chandan.p\Desktop\OMR\5_NetSuite_Booking\Byrathi\Booked"


def cost_inr(pages):
    in_tok = pages * TOK_PER_PAGE + PROMPT_TOK
    out_tok = pages * OUT_TOK_PER_PAGE
    usd = in_tok * IN_USD_PER_1M / 1e6 + out_tok * OUT_USD_PER_1M / 1e6
    return usd * USD_TO_INR, in_tok


def main(folder):
    folder = Path(folder)
    pdfs = sorted(folder.rglob("*.pdf"))
    if not pdfs:
        print("No PDFs found in", folder); return

    seen = {}          # md5 -> first filename
    rows = []          # (name, pages, scanned, md5, is_dup)
    for p in pdfs:
        data = p.read_bytes()
        h = hashlib.md5(data).hexdigest()
        doc = fitz.open(p)
        pages = len(doc)
        text = "".join(doc[i].get_text() for i in range(pages)).strip()
        doc.close()
        is_dup = h in seen
        if not is_dup:
            seen[h] = p.name
        rows.append((p.name, pages, "scan" if len(text) < 50 else "text", h[:8], is_dup))

    print(f"Folder: {folder}\n")
    print(f"{'PDF':52} {'pages':>5} {'type':>5} {'md5':>9}  dup")
    print("-" * 82)
    for name, pages, typ, h8, dup in rows:
        print(f"{name[:52]:52} {pages:>5} {typ:>5} {h8:>9}  {'DUP' if dup else ''}")

    total_files = len(rows)
    dup_files = sum(1 for r in rows if r[4])
    uniq_files = total_files - dup_files
    total_pages = sum(r[1] for r in rows)
    uniq_pages = sum(r[1] for r in rows if not r[4])

    raw_inr, _ = cost_inr(total_pages)
    dedup_inr, _ = cost_inr(uniq_pages)

    print("\n" + "=" * 40)
    print(f"Files          : {total_files}  ({dup_files} duplicates -> skipped)")
    print(f"Unique files   : {uniq_files}")
    print(f"Total pages    : {total_pages}   (unique: {uniq_pages})")
    print(f"Cost if all sent   : Rs {raw_inr:.2f}")
    print(f"Cost after dedup   : Rs {dedup_inr:.2f}   <-- what you actually pay")
    print(f"Saved by dedup     : Rs {raw_inr - dedup_inr:.2f}")
    if uniq_files:
        print(f"Avg per unique file: Rs {dedup_inr / uniq_files:.3f}")
        print(f"Projection @40/day : Rs {dedup_inr / uniq_files * 40:.1f}/day  |  Rs {dedup_inr / uniq_files * 40 * 30:.0f}/month")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else DEFAULT)

"""File CDMS_Sorted\\<Hub>\\PDFs\\<date>\\*.pdf into the real hub folders.

Booked invoices (in _booked.csv) -> <Hub>\\Booked\\        (done)
Not-yet-booked                    -> <Hub>\\PDFs\\<date>\\   (waiting)
A file already present at the destination is skipped, never overwritten.

    python sort_cdms.py            # DRY RUN - shows the plan, moves nothing
    python sort_cdms.py --go       # actually move
    python sort_cdms.py --go Foo   # source folder other than CDMS_Sorted
"""
import os, shutil, sys
from mark_booked import OUT_ROOT, load_booked

NAME_MAP = {"Mysore": "Mysore Road", "YPR": "Yeshwantpura"}   # CDMS name -> real hub folder


def invoice_of(fn):
    for part in fn.split("-"):
        if part.isdigit() and 9 <= len(part) <= 11:   # 8=date, 14=timestamp, 9-11=invoice
            return part
    return None


def main():
    go = "--go" in sys.argv
    src_name = next((a for a in sys.argv[1:] if not a.startswith("-")), "CDMS_Sorted")
    src = os.path.join(OUT_ROOT, src_name)
    booked = load_booked()
    stats, problems = {}, []
    for hub in sorted(d for d in os.listdir(src) if os.path.isdir(os.path.join(src, d, "PDFs"))):
        dest_hub = NAME_MAP.get(hub, hub)
        if not os.path.isdir(os.path.join(OUT_ROOT, dest_hub)):
            problems.append(f"no main folder for hub '{hub}' -> '{dest_hub}'")
            continue
        s = stats.setdefault(dest_hub, {"booked": 0, "pending": 0, "exists": 0})
        for dp, _, files in os.walk(os.path.join(src, hub, "PDFs")):
            date = os.path.basename(dp)
            for fn in files:
                if not fn.lower().endswith(".pdf"):
                    continue
                inv = invoice_of(fn)
                if not inv:
                    problems.append(f"cannot read invoice no from {fn}")
                    continue
                if inv in booked:
                    dest_dir, bucket = os.path.join(OUT_ROOT, dest_hub, "Booked"), "booked"
                else:
                    dest_dir, bucket = os.path.join(OUT_ROOT, dest_hub, "PDFs", date), "pending"
                dest = os.path.join(dest_dir, fn)
                if os.path.exists(dest):
                    s["exists"] += 1
                    continue
                s[bucket] += 1
                if go:
                    os.makedirs(dest_dir, exist_ok=True)
                    shutil.move(os.path.join(dp, fn), dest)

    print("MODE:", "MOVE" if go else "DRY RUN (nothing moved)")
    tb = tp = te = 0
    for hub in sorted(stats):
        s = stats[hub]
        tb += s["booked"]; tp += s["pending"]; te += s["exists"]
        print(f"  {hub:<14} -> Booked {s['booked']:>3} | PDFs(pending) {s['pending']:>3} | already-there {s['exists']:>3}")
    print(f"  {'TOTAL':<14} -> Booked {tb:>3} | PDFs(pending) {tp:>3} | already-there {te:>3}")
    for p in problems:
        print("PROBLEM:", p)


if __name__ == "__main__":
    main()

r"""Show every downloaded invoice PDF in plain words -- hub, invoice no, date, brand,
and BOOKED-or-not -- so you never have to read cryptic PDF filenames by hand.

    python status.py            # print a per-hub summary + write 5_NetSuite_Booking\STATUS.csv
    python status.py Byrathi    # only that hub

_booked.csv (written by mark_booked.py) is the single source of truth for "booked".
Read-only: this script never changes, moves, or deletes anything.
"""
import csv, os, sys
from download_from_links import parse            # filename -> (hub, date, invoice, brand)
from mark_booked import OUT_ROOT, load_booked, LEDGER


def marked_on():
    m = {}
    if os.path.exists(LEDGER):
        with open(LEDGER, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                m[str(r["invoice_no"]).strip()] = r.get("marked_on", "")
    return m


def nice_date(d):
    p = (d or "").split("-")                      # parse() gives YYYY-MM-DD
    return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 and p[0].isdigit() else (d or "")


def scan(only=None):
    booked, mon, rows = load_booked(), marked_on(), []
    for hub in sorted(os.listdir(OUT_ROOT)):
        hub_dir = os.path.join(OUT_ROOT, hub)
        if not os.path.isdir(hub_dir):
            continue
        if only and hub.lower() != only.lower():
            continue
        for folder in ("PDFs", "Booked"):        # PDFs = pending pile, Booked = done pile
            for dp, _, files in os.walk(os.path.join(hub_dir, folder)):
                for fn in sorted(files):
                    if not fn.lower().endswith(".pdf"):
                        continue
                    _, date, inv, brand = parse(fn)
                    is_bk = inv in booked
                    rows.append({"Hub": hub, "Invoice": inv, "Date": nice_date(date),
                                 "Brand": brand, "Status": "BOOKED" if is_bk else "PENDING",
                                 "MarkedOn": mon.get(inv, ""), "Folder": folder, "File": fn})
    return rows


def main():
    only = sys.argv[1] if len(sys.argv) > 1 else None
    rows = scan(only)

    out = os.path.join(OUT_ROOT, "STATUS.csv")
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["Hub", "Invoice", "Date", "Brand",
                                          "Status", "MarkedOn", "Folder", "File"])
        w.writeheader(); w.writerows(rows)

    # terminal = quick glance per hub; STATUS.csv = full detail to open in Excel
    hubs = {}
    for r in rows:
        b, p = hubs.get(r["Hub"], (0, 0))
        hubs[r["Hub"]] = (b + (r["Status"] == "BOOKED"), p + (r["Status"] == "PENDING"))
    print(f"{'Hub':<16}{'BOOKED':>8}{'PENDING':>9}")
    tb = tp = 0
    for h in sorted(hubs):
        b, p = hubs[h]; tb += b; tp += p
        print(f"{h:<16}{b:>8}{p:>9}")
    print(f"{'-'*33}\n{'TOTAL':<16}{tb:>8}{tp:>9}")
    print(f"\nPENDING = still to book (PDF in <Hub>\\PDFs). BOOKED = done (moved to <Hub>\\Booked).")
    print(f"Full list (open in Excel): {out}")


if __name__ == "__main__":
    main()

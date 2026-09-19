"""Combine several hubs' NetSuite import CSVs into ONE Header + ONE Expenses file
so the whole day imports in a single job.

Usage:
    python combine_import_csvs.py "Mysore Road" Byrathi Chrompet

Writes: 5_NetSuite_Booking\\Combined_<today>\\1_Combined_Header.csv
                                            \\2_Combined_Expenses.csv

Safe because all hubs share identical columns and invoice numbers (External ID)
are unique per invoice. The two asserts below ARE the self-check: they run on the
real files every time and stop the combine if columns differ or an External ID
repeats across hubs (which would double-book in one import).
"""
import csv, os, sys, zipfile
from datetime import date

# Project root = parent of 0_Scripts; self-locating so the top folder can move.
OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    assert rows, f"empty file: {path}"
    return rows[0], rows[1:]   # header, data rows


def combine(prefix, name, hubs):
    header, rows = None, []
    for h in hubs:
        p = os.path.join(OUT_ROOT, h, f"{prefix}_{h}_Bills_{name}.csv")
        assert os.path.exists(p), f"missing: {p} — run generate_import_csvs.py {h} first"
        hdr, rws = read_csv(p)
        if header is None:
            header = hdr
        assert hdr == header, f"{name} columns differ in '{h}':\n {hdr}\n vs\n {header}"
        rows += rws
    return header, rows


def main():
    hubs = sys.argv[1:]
    assert hubs, 'give hub names, e.g. python combine_import_csvs.py "Mysore Road" Byrathi Chrompet'

    h_header, h_rows = combine("1", "Header", hubs)
    e_header, e_rows = combine("2", "Expenses", hubs)

    # External IDs must be unique across hubs, else one import double-books.
    ext = h_header.index("External ID")
    seen, dups = set(), []
    for r in h_rows:
        (dups.append(r[ext]) if r[ext] in seen else seen.add(r[ext]))
    assert not dups, f"duplicate External IDs across hubs (would double-book): {sorted(set(dups))}"

    out_dir = os.path.join(OUT_ROOT, f"Combined_{date.today().isoformat()}")
    os.makedirs(out_dir, exist_ok=True)
    h_path = os.path.join(out_dir, "1_Combined_Header.csv")
    e_path = os.path.join(out_dir, "2_Combined_Expenses.csv")
    with open(h_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(h_header); w.writerows(h_rows)
    with open(e_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(e_header); w.writerows(e_rows)

    # bundle every referenced PDF into ONE zip -> a single File Cabinet upload for the day
    att = h_header.index("Attached file")
    wanted = [r[att] for r in h_rows if r[att]]
    found = {}
    for h in hubs:
        for dp, _, files in os.walk(os.path.join(OUT_ROOT, h, "PDFs")):
            for fn in files:
                if fn in wanted and fn not in found:
                    found[fn] = os.path.join(dp, fn)
    z_path = os.path.join(out_dir, "Combined_Invoices.zip")
    with zipfile.ZipFile(z_path, "w", zipfile.ZIP_DEFLATED) as z:
        for fn in wanted:
            if fn in found:
                z.write(found[fn], fn)
    missing = [fn for fn in wanted if fn not in found]

    print(f"Header  : {h_path}  ({len(h_rows)} bills)")
    print(f"Expenses: {e_path}  ({len(e_rows)} lines)")
    print(f"Zip     : {z_path}  ({len(wanted) - len(missing)}/{len(wanted)} PDFs)")
    print(f"Hubs    : {', '.join(hubs)}")
    for fn in missing:
        print("WARN: PDF not found:", fn)


if __name__ == "__main__":
    main()

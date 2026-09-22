"""Build a human-entry workbook for one hub from the grn_copy CSV.

Usage:
    python build_entry_workbook.py OMR
    python build_entry_workbook.py Yeshwantpura "C:\\path\\to\\some.csv"

Output: <OUT_ROOT>\\<LOCATION>_Entry.xlsx  with two sheets:
  - Config : per-hub values, resolved from HUBS below (edit if a hub is missing)
  - Entry  : one row per invoice that has a PDF; you fill Amount_5% / Amount_18%,
             and tick IGST? only for the rare inter-state invoice.
Only HUL / HUL SAMADHAN rows with an S3 link are included.
"""
import csv, os, sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

DEFAULT_CSV = r"C:\Users\chandan.p\Downloads\grn_copy_upload_data (3).csv"
# Project root = parent of 0_Scripts; self-locating so the top folder can be renamed/moved.
OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
BRANDS = {"HUL", "HUL SAMADHAN"}

YELLOW = PatternFill("solid", fgColor="FFF2CC")   # you fill these
BOLD = Font(bold=True)

# Values common to every hub.
COMMON = {
    "Department": "Sales Cost", "Class": "Distribution", "Currency": "INR",
    "Exchange Rate": "1.00", "Posting Period": "Sep 2026", "Approval Status": "Approved",
    "Account": "50020 Cost of Goods Sold : Purchase",
    "INDIA TAX SECTION CODE": "194Q TDS on Purchases", "India Tax Nature": "Goods",
    "HSN 5%": "4090000", "HSN 18%": "33059011",
    "Memo": "Being Purchase for the month of Sep'26",
}

# Per-hub config, keyed by fc_name (lowercased). GSTIN/address have an intra + optional IGST variant.
HUBS = {
    "omr": {
        "Vendor": "RPP-DIST-0396 HINDUSTAN UNILEVER LIMITED_HUL_OMR",
        "Location": "OMR", "BRANDS / PROJECTS": "HUL_OMR", "Place of Supply": "33-Tamil Nadu",
        "GSTIN intra": "33AAACH1004N1Z1",
        "Address intra": "101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,",
        "GSTIN igst": "", "Address igst": "",
    },
    "chrompet": {
        "Vendor": "RPP-DIST-0388 HINDUSTAN UNILEVER LIMITED_HUL_Chrompet",
        "Location": "Chrompet", "BRANDS / PROJECTS": "HUL_Chrompet", "Place of Supply": "33-Tamil Nadu",
        "GSTIN intra": "33AAACH1004N1Z1",
        "Address intra": "101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,",
        "GSTIN igst": "", "Address igst": "",
    },
    "yeshwantpura": {
        "Vendor": "RPP-DIST-0014 HINDUSTAN UNILIVER LIMITED_HUL_YSP",
        "Location": "Yeshwantpura", "BRANDS / PROJECTS": "HUL_YPR", "Place of Supply": "29-Karnataka",
        "GSTIN intra": "29AAACH1004N1ZQ",
        # NOTE: NetSuite's label misspells "LOGITICS" (not LOGISTICS) — must match exactly.
        "Address intra": "HINDUSTAN UNILEVER LIMITED C/O LINFOX LOGITICS INDIA PVT LTD S NO. 74/2, GUVALAKANAHALLI NH NO.7. CHIKKABALLAPUR",
        "GSTIN igst": "33AAACH1004N1Z1",
        "Address igst": "Hindustan Unilever Ltd Ponds House, No:101 Santhome High Road",
    },
    "soukya": {
        "Vendor": "RPP-DIST-0250 Hindustan Uniliver Limited_Hul_Soukya",
        "Location": "Soukya", "BRANDS / PROJECTS": "HUL_Soukya", "Place of Supply": "29-Karnataka",
        "GSTIN intra": "29AAACH1004N1ZQ",
        "Address intra": "C/O LINFOX LOGITICS INDIA PVT LTD",
        "GSTIN igst": "33AAACH1004N1Z1",
        "Address igst": "101, PONDS HOUSE, SANTHOME HIGH ROAD, ,",
    },
    "byrathi": {
        "Vendor": "RPP-DIST-0370 HINDUSTAN UNILIVER LIMITED_HUL_Byrathi",  # UNILIVER misspelled, verbatim
        "Location": "Byrathi", "BRANDS / PROJECTS": "HUL_ Byrathi", "Place of Supply": "29-Karnataka",  # note space after HUL_
        "GSTIN intra": "29AAACH1004N1ZQ",  # bill-verified Vendor Tax Reg. Number
        # NOTE: label misspells "LOGITICS" (not LOGISTICS) — must match exactly. Same label as YSP.
        "Address intra": "HINDUSTAN UNILEVER LIMITED C/O LINFOX LOGITICS INDIA PVT LTD S NO. 74/2, GUVALAKANAHALLI NH NO.7. CHIKKABALLAPUR",
        "GSTIN igst": "33AAACH1004N1Z1",
        "Address igst": "No.101, Santhome High Road",
    },
    "mysore road": {  # key MUST match LOC value lowercased ("TLBL"->"Mysore Road"->"mysore road")
        "Vendor": "RPP-DIST-0017 HINDUSTAN UNILIVER LIMITED_HUL_MYSORE ROAD",  # UNILIVER misspelled, verbatim
        "Location": "Mysore Road", "BRANDS / PROJECTS": "HUL_Mysore Road", "Place of Supply": "29-Karnataka",
        "GSTIN intra": "29AAACH1004N1ZQ",  # bill-verified Vendor Tax Reg. Number
        # NOTE: this hub's label differs from Byrathi/YSP — no company prefix, no CHIKKABALLAPUR, ends with period.
        "Address intra": "C/O LINFOX LOGITICS INDIA PVT LTD S NO. 74/2, GUVALAKANAHALLI NH NO.7.",
        "GSTIN igst": "33AAACH1004N1Z1",
        "Address igst": "No.101, Santhome High Road,",  # trailing comma (Byrathi's has none) — exact
    },
    "gouribidanur": {  # Location spelled OU (Gouribidanur); vendor/brand spelled AU (Gauribidanur) — verbatim
        "Vendor": "RPP-DIST-0170 HINDUSTAN UNILEVER LTD._HUL_GAURIBIDANUR",  # UNILEVER LTD. (correct spelling here)
        "Location": "Gouribidanur", "BRANDS / PROJECTS": "HUL_Gauribidanur", "Place of Supply": "29-Karnataka",
        "GSTIN intra": "29AAACH1004N1ZQ",  # bill- & vendor-record-verified
        # This label uses correct "LOGISTICS" (with S) + "C/O." with period + trailing comma — unique to this hub.
        "Address intra": "C/O. LINFOX LOGISTICS INDIA PVT LTD,S.NO.74/2,",
        "GSTIN igst": "", "Address igst": "",  # vendor record has ONLY the KA address — no Chennai/IGST row
    },
    "hosur": {  # REVERSED orientation: POS is 33-TN, so "intra"(CGST+SGST)=Chennai/33, "igst"=Karnataka/29
        "Vendor": "RPP-DIST-0162 HINDUSTAN UNILEVER LTD._HUL_HOSUR",  # UNILEVER LTD. (correct), _HUL_HOSUR
        "Location": "Hosur", "BRANDS / PROJECTS": "HUL_ Hosur", "Place of Supply": "33-Tamil Nadu",  # space after HUL_
        # intra = GSTIN state == POS state (33==33 -> CGST+SGST). Vendor record's DEFAULT billing = this TN address.
        "GSTIN intra": "33AAACH1004N1Z1",  # bill-verified (Vendor Tax Reg. Number) + vendor Default Tax Reg
        "Address intra": "Ponda House, No:101 Santhome High",  # LABEL verbatim: "Ponda" (not Ponds), colon, truncated
        # igst = Karnataka GSTIN 29 vs POS 33 -> IGST. Second address row on the vendor record.
        "GSTIN igst": "29AAACH1004N1ZQ",
        "Address igst": "CFA LINFOX LOGISTICS INDIA PVT LTD, S NO 74-2",  # LABEL verbatim: CFA (not C/O), "74-2" hyphen
    },
    "tumkur": {  # key = LOC["TMKR"] lowercased. Vendor has NO "_HUL_" segment — verbatim, correct UNILEVER spelling
        "Vendor": "RPP-DIST-0346 Hindustan Unilever Limited_Tumkur",
        "Location": "Tumkur", "BRANDS / PROJECTS": "HUL_ Tumkur", "Place of Supply": "29-Karnataka",  # space after HUL_
        "GSTIN intra": "29AAACH1004N1ZQ",  # bill- & vendor-record-verified (Default Tax Reg / Vendor Tax Reg. Number)
        # Unique label: mixed case, "C/o" no period, comma after linfox, "S No," with comma, trailing comma — exact.
        "Address intra": "C/o linfox, Logistics india pvt ltd S No, 74/2,",
        "GSTIN igst": "", "Address igst": "",  # vendor record has ONLY the KA address — no Chennai/IGST row
    },
    "pondicherry": {  # ALWAYS IGST: POS is 34-Pondicherry but vendor's only GSTIN is 33-TN (Chennai) -> inter-state.
        "Vendor": "RPP-DIST-0084 HINDUSTAN UNILIVER LIMITED_HUL_Pondicherry",  # UNILIVER misspelled, verbatim; LOC code PNCY
        "Location": "Pondicherry", "BRANDS / PROJECTS": "HUL_Pondicherry", "Place of Supply": "34-Pondicherry",
        # Vendor's only tax reg (TN/Chennai). 33 != POS 34 -> generator books IGST automatically, no flag needed.
        "GSTIN intra": "33AAACH1004N1Z1",
        "Address intra": "101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI",  # vendor Default billing LABEL (verify trailing punctuation)
        "GSTIN igst": "", "Address igst": "",  # no second registration; single-GSTIN vendor
    },
}

# Other spellings of a hub seen in the GRN CSV / CDMS portal -> HUBS key. Every script resolves
# fc_name through hub_key(), so a new spelling is fixed here once ("Chromepet" was silently dropped).
ALIASES = {"chromepet": "chrompet", "mysore": "mysore road", "ypr": "yeshwantpura"}


def hub_key(name):
    k = (name or "").strip().lower()
    return ALIASES.get(k, k)


def to_ddmmyyyy(iso):  # 2026-09-13 -> 13/09/2026
    p = (iso or "").split("-")
    return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 else iso


def config_rows(hub):
    h = HUBS[hub]
    rows = [
        ("Vendor", h["Vendor"]),
        ("Vendor Tax Reg. Number", h["GSTIN intra"]),
        ("Vendor Select (address label)", h["Address intra"]),
        ("Vendor Tax Reg. Number (IGST)", h["GSTIN igst"]),
        ("Vendor Select (IGST address label)", h["Address igst"]),
        ("Location", h["Location"]),
        ("BRANDS / PROJECTS", h["BRANDS / PROJECTS"]),
        ("Place of Supply", h["Place of Supply"]),
    ]
    rows += list(COMMON.items())
    return rows


# Entry-sheet layout, shared by the per-hub builder and build_master_workbook.
ENTRY_HEADERS = ["External ID", "Invoice Date", "Brand", "Attached file",
                 "Amount_5%", "Amount_18%", "IGST? (y=IGST)",
                 "CGST 2.5%", "SGST 2.5%", "CGST 9%", "SGST 9%", "IGST",
                 "Total GST", "Grand Total (Taxable+GST)"]
ENTRY_WIDTHS = [15, 12, 13, 42, 12, 12, 13, 11, 11, 11, 11, 11, 11, 20]
# Master_Entry columns after N. Hub must stay first (generate/mark_booked read it by position);
# the rest are looked up by NAME. "No Gemini": type x -> fill_tax_from_pdfs never reads that row.
NO_GEMINI, PRINTED_TAX, GEMINI_CHECK, GEMINI_NOTE = ("No Gemini (type x)", "Printed Tax (Gemini)",
                                                     "Gemini Check", "Gemini note")
MASTER_EXTRA = ["Hub", "Remark", NO_GEMINI, PRINTED_TAX, GEMINI_CHECK, GEMINI_NOTE]
MASTER_EXTRA_WIDTHS = [16, 30, 11, 14, 12, 60]


def add_entry_row(ws, inv, idate, brand, pdf, igst_editable):
    """Append one invoice row + its rate-wise tax cross-check formulas. Columns A..N are
    fixed so the formulas are stable; a master sheet may add extra columns AFTER N."""
    ws.append([inv, idate, brand, pdf, None, None, None])  # E,F,G are the fill-in cells
    row = ws.max_row
    ws.cell(row, 5).fill = YELLOW       # Amount_5%
    ws.cell(row, 6).fill = YELLOW       # Amount_18%
    if igst_editable:
        ws.cell(row, 7).fill = YELLOW   # IGST? flag (blank=intra CGST+SGST, y=IGST)
    ig = f'LOWER(TRIM(G{row}))="y"'
    ws.cell(row, 8).value  = f"=IF({ig},0,ROUND(E{row}*0.025,2))"             # CGST 2.5%
    ws.cell(row, 9).value  = f"=H{row}"                                       # SGST 2.5%
    ws.cell(row, 10).value = f"=IF({ig},0,ROUND(F{row}*0.09,2))"              # CGST 9%
    ws.cell(row, 11).value = f"=J{row}"                                       # SGST 9%
    ws.cell(row, 12).value = f"=IF({ig},ROUND(E{row}*0.05+F{row}*0.18,2),0)"  # IGST
    ws.cell(row, 13).value = f"=H{row}+I{row}+J{row}+K{row}+L{row}"           # Total GST
    ws.cell(row, 14).value = f"=E{row}+F{row}+M{row}"                         # Grand total
    for col in range(5, 15):
        ws.cell(row, col).number_format = "#,##0.00"
    return row


def main():
    key = (sys.argv[1] if len(sys.argv) > 1 else "OMR").strip()
    csv_path = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_CSV
    hub = key.lower()
    if hub not in HUBS:
        sys.exit(f"Hub '{key}' not in HUBS. Add its config to build_entry_workbook.py first. "
                 f"Known: {', '.join(HUBS)}")
    location = HUBS[hub]["Location"]

    rows = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if (r.get("brand_name") or "").strip().upper() not in BRANDS:
                continue
            if hub_key(r.get("fc_name")) != hub:
                continue
            if not (r.get("s3_file_link") or "").strip().startswith("https://cdms-signed-invoice.s3"):
                continue
            rows.append(r)
    if not rows:
        sys.exit(f"No HUL/HULS invoices with a PDF link found for '{key}'.")

    wb = Workbook()
    cfg = wb.active
    cfg.title = "Config"
    cfg["A1"], cfg["B1"] = "Field", "Value"
    cfg["A1"].font = cfg["B1"].font = BOLD
    for i, (field, value) in enumerate(config_rows(hub), start=2):
        cfg[f"A{i}"] = field
        cfg[f"B{i}"] = value
    cfg.column_dimensions["A"].width = 34
    cfg.column_dimensions["B"].width = 70

    ent = wb.create_sheet("Entry")
    ent.append(ENTRY_HEADERS)
    for c in range(1, len(ENTRY_HEADERS) + 1):
        ent.cell(1, c).font = BOLD
    has_igst = bool(HUBS[hub]["GSTIN igst"])
    for r in rows:
        add_entry_row(ent,
                      (r.get("invoice_no") or "").strip(),
                      to_ddmmyyyy(r.get("invoice_date")),
                      (r.get("brand_name") or "").strip(),
                      os.path.basename((r.get("s3_file_link") or "").strip()),
                      igst_editable=has_igst)
    for i, w in enumerate(ENTRY_WIDTHS, start=1):
        ent.column_dimensions[chr(64 + i)].width = w
    ent.freeze_panes = "E2"   # keep ID/date/brand/file visible while typing amounts

    hub_dir = os.path.join(OUT_ROOT, location)   # each hub keeps its own folder
    os.makedirs(hub_dir, exist_ok=True)
    out = os.path.join(hub_dir, f"{location}_Entry.xlsx")
    wb.save(out)
    print(f"Built: {out}")
    print(f"Hub: {key} -> Location {location}, POS {HUBS[hub]['Place of Supply']}, "
          f"{'has IGST variant' if has_igst else 'intra-state only'}")
    print(f"Invoices to enter: {len(rows)}")
    print("Fill YELLOW cells: Amount_5%, Amount_18%, IGST? (rare).")
    print("Cross-check: CGST 2.5% / SGST 2.5% / CGST 9% / SGST 9% shown rate-wise to match the invoice.")


if __name__ == "__main__":
    main()

"""Single 5_NetSuite_Booking\\MISSING_PDFS.xlsx with one tab per source.

    Missing_Links     - written by download_hul_pdfs.py (invoices whose S3 link is blank)
    Wrong_File_Chase  - written by mark_booked.py       (listed but not booked + your Remark)

write_sheet() replaces ONLY its own tab and keeps the other, so the two writers never
clobber each other. Both tabs carry an invoice_no column -> copy those numbers into
CDMS_Tool\\invoices.txt to pull the PDFs from the CDMS portal.
"""
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

OUT_ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "5_NetSuite_Booking")
XLSX = os.path.join(OUT_ROOT, "MISSING_PDFS.xlsx")
BOLD = Font(bold=True)


def write_sheet(sheet_name, headers, rows, widths=None):
    """Upsert one tab into MISSING_PDFS.xlsx. Returns the file path."""
    os.makedirs(OUT_ROOT, exist_ok=True)
    if os.path.exists(XLSX):
        wb = load_workbook(XLSX)
    else:
        wb = Workbook()
        wb.remove(wb.active)            # drop the default empty sheet
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])       # replace our tab, leave the other one alone
    ws = wb.create_sheet(sheet_name)
    ws.append(list(headers))
    for c in range(1, len(headers) + 1):
        ws.cell(1, c).font = BOLD
    for r in rows:
        ws.append(list(r))
    for i, w in enumerate(widths or [16] * len(headers), start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"
    wb.save(XLSX)
    return XLSX

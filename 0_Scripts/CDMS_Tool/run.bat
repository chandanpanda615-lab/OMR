@echo off
cd /d "%~dp0"
python download_invoices.py
echo.
echo Done. PDFs are on your Desktop in CDMS_Invoices, report is CDMS_PDF_Result.xlsx
pause

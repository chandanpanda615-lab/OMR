@echo off
cd /d "%~dp0"
python download_invoices.py
echo.
echo Done. New links are in 5_NetSuite_Booking\CDMS_Recovered.csv - the next download_hul_pdfs / build_master_workbook run picks them up.
pause

@echo off
echo Generating NetSuite Upload CSV from OMR_Purchase_Invoices_Entry.xlsx...
python "%~dp0export_netsuite_csv.py"
echo.
pause

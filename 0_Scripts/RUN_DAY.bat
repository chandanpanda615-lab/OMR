@echo off
rem Morning, in one go:   RUN_DAY.bat "C:\path\to\grn.csv"
rem   STEP 1  download_hul_pdfs.py     (CDMS check of the Missing PDFs list first, if token.txt is valid)
rem   STEP 2  build_master_workbook.py (Master_Entry.xlsx, pre-filled from saved Gemini answers,
rem           + e-invoice QR check: QR Check column, Invoice Date from the QR, WRONG FILE Remark)
cd /d "%~dp0"
if "%~1"=="" (
  echo Usage: RUN_DAY.bat "C:\path\to\grn.csv"
  pause
  exit /b 1
)
python download_hul_pdfs.py "%~1" || goto :fail
echo.
python build_master_workbook.py "%~1" || goto :fail
echo.
echo Done. Open 5_NetSuite_Booking\Master_Entry.xlsx - type amounts yourself and/or x in "No Gemini",
echo then optionally: python fill_tax_from_pdfs.py    then: python generate_master_csvs.py
echo Rows you type: "3-Way Match" (right side) must say MATCH. WRONG FILE rows already have their Remark.
pause
exit /b 0
:fail
echo.
echo STOPPED - read the message above.
pause
exit /b 1

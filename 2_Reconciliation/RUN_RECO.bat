@echo off
REM Double-click, then pick which folder to reconcile from the list it shows.
REM First: drop all hub GRN files + the NS export (renamed to exactly NS.xlsx) into that folder.
cd /d "%~dp0"
python reconcile.py
echo.
pause

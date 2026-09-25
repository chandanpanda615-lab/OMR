# Workflow at a glance

One lane every day. Every step prints plain counts and updates the **Missing PDFs list**
(`5_NetSuite_Booking\MISSING_PDFS.xlsx`, tab `Chase`). Full command reference: **README.md**.

```
 (CDMS token refreshes itself from refresh.txt - see CDMS_Tool\HOW_TO_RUN.txt)

 STEP 1  python download_hul_pdfs.py "<grn.csv>"           ─► Missing PDFs list: CDMS re-check + no-link rows
           a. CDMS check of the Missing PDFs list (only if token valid; else "SKIPPED")
           b. download GRN + CDMS-found PDFs → <Hub>\PDFs\<yyyy-mm-dd>\
           c. prints per hub: Downloaded now | Already on disk | Booked (skip) | No link (list)

 STEP 2  python build_master_workbook.py "<grn.csv>"       ─► Missing PDFs list: unknown hub rows
           → ONE Master_Entry.xlsx (pre-filled from SAVED Gemini answers - free)
           (RUN_DAY.bat "<grn.csv>" = steps 1 + 2)

 STEP 3  DEFAULT: python fill_tax_from_pdfs.py   (Gemini reads every untouched row and fills it)
           answers saved first → 5_NetSuite_Booking\Gemini\raw\<md5>.json (never paid twice)
           text PDFs (Chrompet) are read locally first - free, 5 math checks; only scans go to Gemini
           → Master filled from those; unsure rows = red REVIEW with the printed tax shown
           YOU then cross-check: spot-check PASS rows, properly verify every REVIEW row
         manual override (skip Gemini for a row, or all): in Master_Entry.xlsx, row by row,
           type amounts yourself  |  x in "No Gemini"  |  Remark = can't book (wrong file)

 STEP 4  python generate_master_csvs.py                    ─► Missing PDFs list: rows left out of the batch
           MONEY CHECK: stops if typed amounts ≠ printed tax on the PDF
           → 1_ALL_*.csv + ALL_Invoices.zip, and Gemini\Gemini_Results.xlsx (Gemini vs entered)

 STEP 5  NetSuite import (zip FIRST)
 STEP 6  python mark_booked.py        (only after NetSuite OK) ─► booked rows LEAVE the list
 STEP 7  python archive_batch.py
```

## What the step-1 counts mean
| Column | Meaning |
|---|---|
| **Downloaded now** | new PDF saved today |
| **Already on disk** | downloaded in an earlier run — not downloaded again (nothing lost, nothing skipped from booking) |
| **Booked (skip)** | already in NetSuite (`_booked.csv`) — ignored so it can never be booked twice |
| **No link (list)** | the GRN has no PDF link yet — put on the Missing PDFs list; step 1 re-checks it in CDMS next time |

## What your entry input triggers
| In `Master_Entry.xlsx` you… | Result | Missing PDFs list |
|---|---|---|
| fill `Amount_5%` / `Amount_18%` | **booked** at step 6 | leaves the list |
| type `x` in `No Gemini` | Gemini never reads that row — you fill it | — |
| type a `Remark`, amounts blank | **not booked** | "wrong file"; that PDF is never listed again, a new one comes back with "VERIFY" |
| leave blank | not booked | "left blank" (+ Gemini's note); listed again next day |

## Safety guarantees
- Nothing goes into `_booked.csv` until NetSuite accepts (only step 6 writes it).
- Step 4 stops when a row's amounts don't match the tax printed on its PDF (as read by Gemini).
  If you checked the PDF and Gemini misread it: `python generate_master_csvs.py --accept=<invoice>`.
- A Gemini answer is saved before it is used; the same PDF is never sent (paid) twice.
- `build_master_workbook.py` refuses to overwrite a `Master_Entry.xlsx` that already has amounts.
- Any script that writes `MISSING_PDFS.xlsx` or `Master_Entry.xlsx` stops **before** changing anything
  if Excel has the file open.

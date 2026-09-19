# Workflow at a glance

Two lanes, linked by `MISSING_PDFS.xlsx`. Full command reference: **README.md**.

## The big picture — two lanes, linked by `MISSING_PDFS.xlsx`

```
 DAILY FLOW (top level)                         CDMS RECOVERY (CDMS_Sorted\, quarantined)
 ─────────────────────                          ─────────────────────────────────────────
 download_hul_pdfs.py                            CDMS_Tool\ (token.txt + invoices.txt)
        │  blank-link → Missing_Links                    │  download_invoices.py
        ▼                                                ▼
 build_master_workbook.py  ──► Master_Entry.xlsx  build_cdms_entry.py ──► CDMS_Entry.xlsx
        │  (fill amounts / Remark)                       │  (fill amounts / Remark)
        ▼                                                ▼
 generate_master_csvs.py                         cdms_generate.py     (CSV+zip, NO tracker change)
        ▼                                                ▼
 ══ NetSuite import ══                           ══ NetSuite import ══
        ▼                                                ▼
 mark_booked.py                                  cdms_mark_booked.py   ◄── only after NetSuite OK
        ▼                                                ▼
 archive_batch.py                                archive_cdms.py
        │                                                │
        └──────────►  MISSING_PDFS.xlsx  ◄───────────────┘
              Missing_Links (no PDF) · Wrong_File_Chase (wrong/DO-NOT-BOOK)
                              │
                              └──► feeds invoices.txt → CDMS_Tool  (the loop closes)
```

## Daily flow — step by step
| # | Command | Produces | Tracker effect |
|---|---|---|---|
| 1 | `download_hul_pdfs.py "<grn>.csv"` | PDFs → `<Hub>\PDFs\<date>\` | blank-link invoices → `Missing_Links` |
| 2 | `build_master_workbook.py "<grn>.csv"` | `Master_Entry.xlsx` (D cell = PDF link) | — |
| 3 | *you fill* amounts / `Remark` | — | — |
| 4 | `generate_master_csvs.py` | `1_ALL_*`, `2_ALL_*`, verify, zip | **none** |
| 5 | NetSuite import | booked bills | — |
| 6 | `mark_booked.py` *(after OK)* | ledger + PDFs → `Booked\` | booked→drop `Missing_Links`; `Remark`→`Wrong_File_Chase` |
| 7 | `archive_batch.py` | `Archive\Master_<date>\` | — |

## CDMS recovery — step by step
| # | Command | Produces | Tracker effect |
|---|---|---|---|
| 1 | `CDMS_Tool\download_invoices.py` | PDFs → `Desktop\CDMS_Invoices` + result xlsx | — |
| 2 | `build_cdms_entry.py` | `CDMS_Entry.xlsx`; PDFs → `CDMS_Sorted\<Hub>\PDFs\` | 5 wrong ones pre-flagged |
| 3 | *you fill* amounts / `Remark` | — | — |
| 4 | `cdms_generate.py` | `1_CDMS_*`, `2_CDMS_*`, verify, zip | **none** (safe to re-run) |
| 5 | NetSuite import | booked bills | — |
| 6 | `cdms_mark_booked.py` *(after OK)* | ledger + PDFs → `<Hub>\Booked\` | booked→drop both tabs; `Remark`→`Wrong_File_Chase` |
| 7 | `archive_cdms.py` | `CDMS_Sorted\Archive\CDMS_<date>\` | — |

## What your entry input triggers (the cascade — fires at step 6 only)
| In the entry sheet you… | Result | Where it goes |
|---|---|---|
| fill `Amount_5%` / `Amount_18%` | **booked** | ledger `_booked.csv`; PDF → `Booked\`; dropped from trackers |
| type a `Remark`, amounts blank | **DO NOT BOOK** | `MISSING_PDFS.xlsx → Wrong_File_Chase` |
| leave blank, no remark | untouched | stays to work next time |

**Key guarantee:** nothing leaves the trackers until NetSuite accepts (step 4 changes no tracker; only step 6 does). If NetSuite rejects, fix and re-run step 4 — the safety net stays intact.

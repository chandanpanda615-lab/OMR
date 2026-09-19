# The `--skip` flag explained (generate_import_csvs.py)

`--skip` tells the script: **"leave out these invoice numbers, book only the rest."**
Use it when some invoices are already booked and you don't want to import them again.

## What happens inside
The number after `--skip=` is the **invoice number** (= External ID = column A in your Excel).
The script reads your Excel and, for every row, checks:
> "Is this invoice number in the skip list?"  → **Yes = ignore it. No = include it.**

Skipped invoices are left out of the CSVs **and** the zip.

## Simple example
Your OMR Excel has **4 invoices**:
```
9633119086
9629071933
9629071934
9629071935
```
You already booked **9633119086** and **9629071934** yesterday. Today you want only the other 2.

Run:
```
python generate_import_csvs.py OMR --skip=9633119086,9629071934
```
Result — the CSVs + zip now contain **only 2 invoices**:
```
9629071933
9629071935
```
The printout will say `2 bills` instead of `4`.

## Why use it (instead of deleting rows from the Excel)
- You keep your full Excel record — nothing deleted.
- Your NetSuite import is clean — no error rows for already-booked invoices.

> Note: even without `--skip`, NetSuite will **not** create a duplicate — the External ID blocks it,
> so a re-import of a booked invoice just fails safely. `--skip` only keeps the file tidy.

## Rules for typing it
- Separate numbers with a **comma**, no spaces needed: `--skip=111,222,333`
- Use the **invoice number** exactly as it appears in column A of the Excel.

## Real case
We used this for Yeshwantpura: 1 of 4 invoices was already booked, so we skipped that one
and regenerated the other 3 for import.

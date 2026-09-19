# MASTER HANDOVER: ORACLE NETSUITE PURCHASE BILL AUTOMATION (HUL FMCG)

**Author / Maintainer:** Antigravity AI Pair Programmer  
**Client / Entity:** Intelligent Retail Pvt Ltd (Ripplr) — NetSuite Account `7100839`  
**User / Role:** Vignesh S (`Accounts Payable_Ripplr`)  
**Scope:** Daily FMCG Vendor Bill Accounting for Hindustan Unilever Limited (HUL) across 10 Distribution Hubs  
**Status:** **PROVEN & VERIFIED** (15 of 15 bills successfully imported with automated GST, TDS, PDF attachments, and GL posting).

---

## 1. Executive Summary & Core Objectives

### The Business Need
Intelligent Retail Pvt Ltd (Ripplr) receives 40–50 scanned FMCG purchase invoices daily from Hindustan Unilever Limited (HUL) across 8–10 retail hubs in Karnataka, Tamil Nadu, and Puducherry. Manual entry into Oracle NetSuite was slow, error-prone, and caused bottlenecks in GST input credit matching and vendor reconciliation.

### The Solution Built
A standardized, 2-file CSV automated import pipeline that:
1. Automatically retrieves signed PDF invoices from Amazon S3.
2. Ingests invoice data and splits line items across GST slabs (5% HSN `4090000` and 18% HSN `33059011`).
3. Uploads packaged PDF files into NetSuite File Cabinet (`Ganesh Folder`) via Advanced Add.
4. Generates two linked CSV files (`Primary File - Vendor Bill` and `Linked File - Expenses`) linked by `External ID`.
5. Triggers NetSuite's India Localization SuiteTax server-side scripts during CSV import to automatically calculate CGST, SGST, IGST, and TDS 194Q.
6. Automatically attaches the PDF invoice to the Bill record's `Attached file` (`custbody11`) field with native preview and download capability.

---

## 2. Tested & Verified Milestone: OMR Hub (September 2026)

### Scope of Test Batch
* **Location:** OMR Hub (Chennai, Tamil Nadu)
* **Invoices in Batch:** 18 total invoices across 5th, 6th, 8th, and 10th September 2026.
* **Manual Exclusion:** 3 invoices from 8th September (`9633120170`, `9633120171`, `9633120172`) were entered manually by the user earlier. **Rule:** These 3 must never be re-imported to avoid duplicates.
* **Unbilled Invoices Imported (15 Total):**
  * **5th September (6 invoices):** `9633118905`, `9633118906`, `9633118907`, `9633118994`, `9633118995`, `9633118996`
  * **6th September (6 invoices):** `9633119242`, `9633119243`, `9633119244`, `9633119272`, `9633119273`, `9633119274`
  * **10th September (3 invoices):** `9633122377`, `9633122378`, `9633122379`

### Verification Result
* **NetSuite Job Status:** `Complete` | `100.0%` | **`15 of 15 records imported successfully`**
* **PDF Attachment:** 100% attached without error (`custbody11` matched File Cabinet).
* **Tax Engine:** Confirmed! Server script executed:
  * CGST @ 2.5% calculated.
  * SGST @ 2.5% calculated.
  * TDS Section 194Q @ 0.1% deducted.
* **GL Impact Posting:** Confirmed!
  * Debit `50020 Cost of Goods Sold : Purchase` (Taxable Amount)
  * Debit `70010 GST Input : Input Tax - CGST` (CGST Amount)
  * Debit `70020 GST Input : Input Tax - SGST` (SGST Amount)
  * Credit `22030 Other Current Liabilities : TDS Payable` (0.1% TDS Amount)
  * Credit `20010 Accounts Payable : Trade Creditors Distribution` (Net Payable Amount)

---

## 3. Critical Technical Pitfalls & Exactly How They Were Solved

### Pitfall 1: Why 1-File CSV Import Fails & 2-File Import is Mandatory
* **Problem:** If you try to combine Header and Expense line items into a single CSV file, sublist columns (`Amount`, `Department`, `Class`, `Location`, `Memo`) have identical names to header fields. NetSuite requires manual drag-and-drop mapping. When you map line-level `Amount`, NetSuite displays a fatal warning: `The field Amount (Req) has been removed from mapping`. If ignored, the transaction is rejected or values are corrupted.
* **Solution:** Use **Multiple Files to Upload**:
  * **Primary File:** `1_OMR_Bills_Header.csv` (Mapped to `Vendor Bill`)
  * **Linked File:** `2_OMR_Bills_Expenses.csv` (Mapped to `Vendor Bill - Expenses`)
  * **Link Key:** `External ID` (Invoice Number)
  * **Benefit:** NetSuite auto-maps 100% of all fields automatically. Zero manual drag-and-drop required!

### Pitfall 2: `Invalid custbody11 reference key <filename>.pdf`
* **Problem:** In NetSuite, `custbody11` is the custom field `Attached file`. If you pass a filename in the CSV before the file exists in NetSuite, the import fails with `Invalid custbody11 reference key`.
* **Solution:**
  1. Package all PDF invoices for the batch into a single `.zip` file (e.g. `15_Invoices.zip`).
  2. In NetSuite, navigate to: `Documents > File Cabinet > Ganesh Folder` (or appropriate hub folder).
  3. Click the blue **`Advanced Add`** button.
  4. Select the `.zip` file, check **`Unzip files`** `[✔]` and **`Overwrite files with same name`** `[✔]`.
  5. Click **`Add`**. NetSuite extracts all PDFs into the File Cabinet in ~10 seconds.
  6. Now run the CSV import. NetSuite matches the filename string and creates the clickable preview/download links.

### Pitfall 3: Missing GST (CGST/SGST/IGST) & Missing TDS Deduction
* **Problem:** In our first test run, 3 bills imported but had `Tax = 0.00` and no TDS.
* **Root Cause:** In Step 2 of the NetSuite CSV Import assistant (`Import Options`), under `Advanced Options`, the checkbox **`RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS`** was unchecked by default.
* **Solution:** **MUST ALWAYS CHECK** 👉 **`[✔] RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS`**.  
  This forces NetSuite's India Localization SuiteTax scripts to run on record creation, calculating GST based on HSN and TDS based on Section 194Q.

### Pitfall 4: `Invalid billaddresslist reference key ... for entity 24039`
* **Problem:** In our second run, all 15 invoices failed with:  
  `Invalid billaddresslist reference key 101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI, Chennai, 33-TN 600028 India for entity 24039.`
* **Root Cause:**  
  1. NetSuite's `billaddresslist` (`Vendor Select` dropdown on Billing tab) is a list field that expects the **exact text in the `LABEL` column** of the Vendor's Address Book, NOT the full 5-line address text!
  2. The user's role is `Accounts Payable_Ripplr`, which has View-only permissions on Vendor records (no `Edit` button on Vendor master), so the user cannot manually check `DEFAULT BILLING` on the Vendor.
  3. By viewing the Vendor record (`id=24039`), we uncovered the exact `LABEL`:  
     `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` *(note the trailing comma!)*.
* **Solution:**
  * **Option A (Safe Default):** Do NOT include `Vendor Select` in the Header CSV. The bill imports cleanly with 100% success, and all taxes calculate perfectly!
  * **Option B (Populated Address):** Pass the exact string from the Vendor's `LABEL` column: `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,`.

### Pitfall 5: How NetSuite Decides Intra-State (CGST/SGST) vs Inter-State (IGST)
* **Mechanism:** NetSuite India Localization SuiteTax does **NOT** determine GST from the billing address text! It determines GST by comparing the first 2 digits of the GSTIN:
  * **Intra-State:** If `Vendor Tax Reg. Number` starts with `33` and `Place of Supply` is `33-Tamil Nadu` $\rightarrow$ NetSuite applies **CGST (2.5%) + SGST (2.5%)**.
  * **Inter-State:** If `Vendor Tax Reg. Number` starts with `29` (Karnataka) and `Place of Supply` is `33-Tamil Nadu` (or vice versa) $\rightarrow$ NetSuite automatically applies **IGST (5% or 18%)**!
* **Rule:** Always provide `Place of Supply` (e.g. `33-Tamil Nadu`) and `Vendor Tax Reg. Number` in `1_Header.csv`.

---

## 4. Master Hub Configuration Matrix (10 Active HUL Hubs)

Sourced directly from the company's master tracker: `C:\Users\chandan.p\Downloads\Purchase - Tracker (2).xlsx` (Sheet: `Aug'26`):

| # | Hub Name | State | NetSuite Location | Place of Supply | NetSuite Brand | NetSuite Vendor Ledger Name | Vendor GSTIN | Address Label (`Vendor Select`) |
|---|---|---|---|---|---|---|---|---|
| 1 | **OMR** | Tamil Nadu | `OMR` | `33-Tamil Nadu` | `HUL_OMR` | `RPP-DIST-0396 HINDUSTAN UNILEVER LIMITED_HUL_OMR` | `33AAACH1004N1Z1` | `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` |
| 2 | **Chrompet** | Tamil Nadu | `Chrompet` | `33-Tamil Nadu` | `HUL_Chrompet` | `RPP-DIST-0388 HINDUSTAN UNILEVER LIMITED_HUL_Chrompet` | `33AAACH1004N1Z1` | `101, PONDS HOUSE, SANTHOME HIGH ROAD, SANTHOME CHENNAI,` |
| 3 | **Hosur** | Tamil Nadu | `Hosur` | `33-Tamil Nadu` | `HUL_ Hosur` | `RPP-DIST-0162 HINDUSTAN UNILEVER LTD._HUL_HOSUR` | `33AAACH1004N1Z1` | Check Vendor `id` |
| 4 | **Yeshwantpura (YPR)** | Karnataka | `YSPR` | `29-Karnataka` | `HUL_YPR` | `RPP-DIST-0014 HINDUSTAN UNILIVER LIMITED_HUL_YSP` | `29AAACH1004N1Z3` | Check Vendor `id` |
| 5 | **Mysore Road** | Karnataka | `Mysore Road` | `29-Karnataka` | `HUL_Mysore Road` | `RPP-DIST-0017 HINDUSTAN UNILIVER LIMITED_HUL_MYSORE ROAD` | `29AAACH1004N1Z3` | Check Vendor `id` |
| 6 | **Soukya** | Karnataka | `Soukya` | `29-Karnataka` | `HUL_Soukya` | `RPP-DIST-0250 Hindustan Uniliver Limited_Hul_Soukya` | `29AAACH1004N1Z3` | Check Vendor `id` |
| 7 | **Byrathi** | Karnataka | `Byrathi` | `29-Karnataka` | `HUL_ Byrathi` | `RPP-DIST-0370 HINDUSTAN UNILIVER LIMITED_HUL_Byrathi` | `29AAACH1004N1Z3` | Check Vendor `id` |
| 8 | **Gouribidanur** | Karnataka | `Gouribidanur` | `29-Karnataka` | `HUL_Gauribidanur` | `RPP-DIST-0170 HINDUSTAN UNILEVER LTD._HUL_GAURIBIDANUR` | `29AAACH1004N1Z3` | Check Vendor `id` |
| 9 | **Tumkur** | Karnataka | `Tumkur` | `29-Karnataka` | `HUL_ Tumkur` | `RPP-DIST-0346 Hindustan Unilever Limited_Tumkur` | `29AAACH1004N1Z3` | Check Vendor `id` |
| 10 | **Pondicherry** | Puducherry | `Pondicherry` | `34-Puducherry` | `HUL_Pondicherry` | `RPP-DIST-0084 HINDUSTAN UNILIVER LIMITED_HUL_Pondicherry` | `34AAACH1004N1..` | Check Vendor `id` |

*Note on Closed / Inactive Hubs:*  
* `RPP-DIST-0095 HINDUSTAN UNILEVER LIMITED-HUL_Pune` (Phursungi, Maharashtra) $\rightarrow$ Marked **CLOSED**.  
* `RPP-DIST-0015 HINDUSTAN UNILIVER LIMITED_HUL-HYD` (Kompally, Telangana) $\rightarrow$ Marked **CLOSED**.  
* `RPP-DIST-0302 HINDUSTAN UNILEVER LTD._HUL _Warangal HYD` (Warangal, Telangana) $\rightarrow$ Marked **NO PURCHASE**.

---

## 5. CSV Data Specifications & NetSuite Mapping

### File 1: Header CSV (`1_<Hub>_Bills_Header.csv`)
Mapped to NetSuite Record: **`Vendor Bill`**

| Column Name | Sample Value | NetSuite Destination Field | Comments |
|---|---|---|---|
| `External ID` | `9633122377` | `External ID` | Unique Key matching Invoice Number |
| `Reference No.` | `9633122377` | `Reference No.` | Invoice Number |
| `Vendor` | `RPP-DIST-0396 HINDUSTAN UNILEVER LIMITED_HUL_OMR` | `Vendor` | Full NetSuite Vendor Name |
| `Vendor Tax Reg. Number` | `33AAACH1004N1Z1` | `Vendor Tax Reg. Number` | Vendor GSTIN |
| `Date` | `10/09/2026` | `Date` | Format: `DD/MM/YYYY` |
| `Posting Period` | `Sep 2026` | `Posting Period` | Format: `Mmm YYYY` |
| `Invoice Date` | `10/09/2026` | `Invoice Date` | Same as Date |
| `Invoice Receipt Date`| `10/09/2026` | `Invoice Receipt Date` | Same as Date |
| `Due Date` | `10/09/2026` | `Due Date` | Same as Date |
| `Tax Point Date` | `10/09/2026` | `Tax Point Date` | Crucial for GST calculation |
| `Approval Status` | `Approved` | `Approval Status` | Directly Approved |
| `Currency` | `INR` | `Currency` | Indian Rupee |
| `Exchange Rate` | `1.00` | `Exchange Rate` | Always 1.00 for domestic |
| `Department` | `Sales Cost` | `Department` | Standard Classification |
| `Class` | `Distribution` | `Class` | Standard Classification |
| `Location` | `OMR` | `Location` | NetSuite Hub Location |
| `BRANDS / PROJECTS` | `HUL_OMR` | `BRANDS / PROJECTS` | Custom Segment |
| `Place of Supply` | `33-Tamil Nadu` | `Place of Supply` | Custom Field (`custbody_ind_gst_pos`) |
| `Attached file` | `OMR-HULS-10092026-9633122377-20260911153446.pdf` | `Attached file` (`custbody11`) | Exact PDF name in File Cabinet |
| `Memo` | `Being Purchase for the month of Sep'26` | `Memo` | Header description |

### File 2: Expenses CSV (`2_<Hub>_Bills_Expenses.csv`)
Mapped to NetSuite Sublist: **`Vendor Bill - Expenses`**

| Column Name | Sample Value | NetSuite Destination Field | Comments |
|---|---|---|---|
| `External ID` | `9633122377` | `External ID` | Links to Header `External ID` |
| `Line` | `1` | (Sublist Line sequence) | `1` for 5%, `2` for 18% |
| `Account` | `50020 Cost of Goods Sold : Purchase` | `Account` | Standard Purchase COGS account |
| `Amount` | `237338.28` | `Amount` | **Net Taxable Amount** (Excluding GST) |
| `Memo` | `Being Purchase for the month of Sep'26` | `Memo` | Line-item description |
| `Department` | `Sales Cost` | `Department` | Inherited from header |
| `Class` | `Distribution` | `Class` | Inherited from header |
| `Location` | `OMR` | `Location` | Inherited from header |
| `BRANDS / PROJECTS` | `HUL_OMR` | `BRANDS / PROJECTS` | Inherited from header |
| `INDIA TAX SECTION CODE` | `194Q TDS on Purchases` | `INDIA TAX SECTION CODE` | Triggers 0.1% TDS deduction |
| `INDIA TAX HSN OR SAC CODE` | `4090000` (for 5%) or `33059011` (for 18%) | `INDIA TAX HSN OR SAC CODE` | Triggers India GST SuiteTax calculation |
| `India Tax Nature` | `Goods` | `India Tax Nature` | Specifies goods vs services |

---

## 6. S3 Signed PDF Retrieval Pipeline

The file `C:\Users\chandan.p\Desktop\OMR\4_Master_Reports\Purchase GRN Data.xlsx` (Sheet: `Sheet1`) contains the master record of all FMCG deliveries, GRNs, and direct Amazon S3 download links for signed PDF invoices.

### Python S3 Bulk Download Pattern:
```python
import os, urllib.request, zipfile, openpyxl

excel_path = r"C:\Users\chandan.p\Desktop\OMR\4_Master_Reports\Purchase GRN Data.xlsx"
wb = openpyxl.load_workbook(excel_path, data_only=True)
ws = wb["Sheet1"]

output_dir = r"C:\Users\chandan.p\Desktop\OMR\Batch_9th_10th\PDFs"
os.makedirs(output_dir, exist_ok=True)

# Filter by fc_name == 'OMR' and relevant invoice numbers
# Column 3 = invoice_no, Column 11 = s3_file_link
for r in range(2, ws.max_row + 1):
    inv = str(ws.cell(r, 3).value or "").strip()
    s3_url = str(ws.cell(r, 11).value or "").strip()
    if s3_url.startswith("https://cdms-signed-invoice.s3"):
        filename = os.path.basename(s3_url)
        target_path = os.path.join(output_dir, filename)
        if not os.path.exists(target_path):
            urllib.request.urlretrieve(s3_url, target_path)

# Package into Zip for NetSuite File Cabinet
zip_path = r"C:\Users\chandan.p\Desktop\OMR\Batch_9th_10th\12_OMR_Invoices.zip"
with zipfile.ZipFile(zip_path, "w") as zf:
    for f in os.listdir(output_dir):
        if f.endswith(".pdf"):
            zf.write(os.path.join(output_dir, f), arcname=f)
```

---

## 7. The Immediate Next Task: 12 More Invoices for OMR (9th & 10th Sep)

The user has highlighted the next 12 unbilled invoices for OMR:

### Invoice List:
1. `9633121142` (09-Sep-26) $\rightarrow$ `OMR-HULS-09092026-9633121142-20260910150949.pdf`
2. `9633121143` (09-Sep-26) $\rightarrow$ `OMR-HULS-09092026-9633121143-20260910151431.pdf`
3. `9633121144` (09-Sep-26) $\rightarrow$ `OMR-HULS-09092026-9633121144-20260910150845.pdf`
4. `9633121217` (09-Sep-26) $\rightarrow$ `OMR-HULS-09092026-9633121217-20260910150901.pdf`
5. `9633121218` (09-Sep-26) $\rightarrow$ `OMR-HULS-09092026-9633121218-20260910150917.pdf`
6. `9633121219` (09-Sep-26) $\rightarrow$ `OMR-HULS-09092026-9633121219-20260910151002.pdf`
7. `9633122278` (10-Sep-26) $\rightarrow$ `OMR-HULS-10092026-9633122278-20260911153310.pdf`
8. `9633122279` (10-Sep-26) $\rightarrow$ `OMR-HULS-10092026-9633122279-20260911153323.pdf`
9. `9633122280` (10-Sep-26) $\rightarrow$ `OMR-HULS-10092026-9633122280-20260911153340.pdf`
10. `9633122287` (10-Sep-26) $\rightarrow$ `OMR-HULS-10092026-9633122287-20260911153353.pdf`
11. `9633122288` (10-Sep-26) $\rightarrow$ `OMR-HULS-10092026-9633122288-20260911153407.pdf`
12. `9633122289` (10-Sep-26) $\rightarrow$ `OMR-HULS-10092026-9633122289-20260911153424.pdf`

### Execution Runbook for the Next Agent:
1. **Download & Zip:** Execute the Python S3 retrieval script above to save all 12 PDFs into `C:\Users\chandan.p\Desktop\OMR\Batch_9th_10th\PDFs` and zip them to `12_OMR_Invoices_9th_10th.zip`.
2. **Pre-populate Entry Workbook:** Generate `OMR_Entry_9th_10th.xlsx` with pre-filled headers, locations, and PDF filenames so the user only enters 5% / 18% amounts (or OCR if enabled).
3. **Generate Import CSVs:** Run the generator script to create `1_OMR_9th_10th_Header.csv` and `2_OMR_9th_10th_Expenses.csv`.
4. **NetSuite Upload:**
   * User uploads `12_OMR_Invoices_9th_10th.zip` to NetSuite File Cabinet (`Ganesh Folder`) via Advanced Add.
   * User runs CSV Import using saved template `HUL Purchase Bill Import`.
   * Check `RUN SERVER SUITESCRIPT AND TRIGGER WORKFLOWS` [✔].
   * Click Save & Run.

---

## 8. Directory & File Inventory

* **Master Reports:**
  * `C:\Users\chandan.p\Desktop\OMR\4_Master_Reports\Purchase GRN Data.xlsx`: Master S3 invoice links and delivery status.
  * `C:\Users\chandan.p\Downloads\Purchase - Tracker (2).xlsx`: Master hub tracker across all states and brands.
* **Working Directory (OMR):**
  * `C:\Users\chandan.p\Desktop\OMR\`: Primary working folder for OMR.
  * `C:\Users\chandan.p\Desktop\OMR\OMR_Purchase_Invoices_Entry.xlsx`: Master entry workbook for OMR.
  * `C:\Users\chandan.p\Desktop\OMR\1_OMR_Bills_Header.csv`: Successfully imported header file.
  * `C:\Users\chandan.p\Desktop\OMR\2_OMR_Bills_Expenses.csv`: Successfully imported expense line file.
  * `C:\Users\chandan.p\Desktop\OMR\15_Invoices.zip`: 15 PDFs uploaded to NetSuite File Cabinet.
  * `C:\Users\chandan.p\Desktop\OMR\Batch_9th_10th\`: New folder designated for the next 12 invoices.

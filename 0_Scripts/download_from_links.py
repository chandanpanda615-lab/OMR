"""Case 2 (no grn CSV, only copied rows/links): paste -> download -> build entry Excel.

Paste anything between the PASTE markers -- full tab rows or bare URLs, both work.
Everything is read from the link filename itself, e.g.
    YSPR-HUL-07092026-9629071933-20260909080204.pdf
    ^loc ^brand ^DDMMYYYY  ^invoice
so extra columns, tabs or spaces you paste are ignored.

Run:  python download_from_links.py            # download PDFs + build <Hub>_Entry.xlsx
      python download_from_links.py --selftest # parse check, no download
Then fill the YELLOW cells in the Excel and run:  python generate_import_csvs.py <Location>
"""
import csv, os, re, subprocess, sys, urllib.request

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Project root = parent of 0_Scripts; self-locating so the top folder can be renamed/moved.
BOOK_ROOT = os.path.join(os.path.dirname(SCRIPT_DIR), "5_NetSuite_Booking")   # per-hub: <Location>\PDFs\<date>

# Link location-code -> nice folder/hub name. Unknown code -> used as-is.
# ponytail: only the codes we've actually seen; add a line when a new hub shows up.
LOC = {"YSPR": "Yeshwantpura", "OMR": "OMR", "CRMP": "Chrompet", "SUKA": "Soukya", "BYTI": "Byrathi", "TLBL": "Mysore Road", "GBDR": "Gouribidanur", "TMKR": "Tumkur", "HSRA": "Hosur", "PNCY": "Pondicherry"}
BRAND = {"HUL": "HUL", "HULS": "HUL SAMADHAN"}

URL_RE = re.compile(r"https://cdms-signed-invoice\.s3[^\s\"']+?\.pdf", re.I)

# ---- PASTE YOUR ROWS/LINKS BETWEEN THE TRIPLE QUOTES ----
PASTE = r"""
Mysore Road	HUL	9629075332	15/09/2026	15/09/2026	PUR08882			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-15092026-9629075332-20260917065124.pdf
Mysore Road	HUL	9629075333	15/09/2026	15/09/2026	PUR08883			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-15092026-9629075333-20260916064235.pdf
Mysore Road	HUL	9629075334	15/09/2026	15/09/2026	PUR08884			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-15092026-9629075334-20260916064252.pdf
Soukya	HUL	9629075335	16/09/2026	16/09/2026	GRN00765	PO-RP-BLR-SUKA04904	ASN-SUKA-000009806	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/SUKA-HUL-16092026-9629075335-20260916163011.pdf
Soukya	HUL	9629075336	16/09/2026	16/09/2026	GRN00766	PO-RP-BLR-SUKA04906	ASN-SUKA-000009808	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/SUKA-HUL-16092026-9629075336-20260916163028.pdf
Soukya	HUL	9629075337	16/09/2026	16/09/2026	GRN00767	PO-RP-BLR-SUKA04905	ASN-SUKA-000009807	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/SUKA-HUL-16092026-9629075337-20260916163047.pdf
Soukya	HUL	9629075338	16/09/2026	16/09/2026	GRN00768	PO-RP-BLR-SUKA04903	ASN-SUKA-000009805	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/SUKA-HUL-16092026-9629075338-20260916163102.pdf
Byrathi	HUL	9629075346	15/09/2026	15/09/2026	GRN05883	PO-RP-BLR-HMR11237	PO-ASN-HMR-000010866	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-15092026-9629075346-20260915181627.pdf
Byrathi	HUL	9629075347	15/09/2026	15/09/2026	GRN05884	PO-RP-BLR-HMR11235	PO-ASN-HMR-000010864	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-15092026-9629075347-20260915181922.pdf
Byrathi	HUL	9629075348	15/09/2026	15/09/2026	GRN05885	PO-RP-BLR-HMR11238	PO-ASN-HMR-000010867	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-15092026-9629075348-20260915182041.pdf
Byrathi	HUL	9629075349	15/09/2026	15/09/2026	GRN05886	PO-RP-BLR-HMR11236	PO-ASN-HMR-000010865	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-15092026-9629075349-20260915182237.pdf
Yeshwantpura	HUL	9629075350	15/09/2026	15/09/2026	GIR06874			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-15092026-9629075350-20260915205726.pdf
Yeshwantpura	HUL	9629075351	15/09/2026	15/09/2026	GIR06876			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-15092026-9629075351-20260915210512.pdf
Yeshwantpura	HUL	9629075352	15/09/2026	15/09/2026	GIR06875			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-15092026-9629075352-20260915205854.pdf
Gouribidanur	HUL	9629075404	15/09/2026	15/09/2026	PUR02736	PO-RP-BLR-GBDR02493	PO-ASN-GBDR-000004519	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/GBDR-HUL-15092026-9629075404-20260915161034.pdf
Gouribidanur	HUL	9629075405	15/09/2026	15/09/2026	PUR02737	PO-RP-BLR-GBDR02495	PO-ASN-GBDR-000004521	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/GBDR-HUL-15092026-9629075405-20260915161049.pdf
Gouribidanur	HUL	9629075406	15/09/2026	15/09/2026	PUR02738	PO-RP-BLR-GBDR02494	PO-ASN-GBDR-000004520	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/GBDR-HUL-15092026-9629075406-20260915161105.pdf
Tumkur	HUL	9629075517	16/09/2026	16/09/2026	PUR00520	PO-RP-BLR-TMKR02126	PO-ASN-TMKR-000002219	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TMKR-HUL-16092026-9629075517-20260916164833.pdf
Tumkur	HUL	9629075518	16/09/2026	16/09/2026	PUR00521	PO-RP-BLR-TMKR02125	PO-ASN-TMKR-000002218	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TMKR-HUL-16092026-9629075518-20260916164848.pdf
Byrathi	HUL	9629075792	16/09/2026	16/09/2026	GRN05895	PO-RP-BLR-HMR11245	PO-ASN-HMR-000010874	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-16092026-9629075792-20260916192429.pdf
Byrathi	HUL	9629075793	16/09/2026	16/09/2026	GRN05896	PO-RP-BLR-HMR11243	PO-ASN-HMR-000010872	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-16092026-9629075793-20260916192632.pdf
Byrathi	HUL	9629075794	16/09/2026	16/09/2026	GRN05897	PO-RP-BLR-HMR11247	PO-ASN-HMR-000010876	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-16092026-9629075794-20260916192750.pdf
Mysore Road	HUL	9629075803	16/09/2026	16/09/2026	PUR08887			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-16092026-9629075803-20260917065016.pdf
Mysore Road	HUL	9629075804	16/09/2026	16/09/2026	PUR08888			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-16092026-9629075804-20260917065030.pdf
Mysore Road	HUL	9629075805	16/09/2026	16/09/2026	PUR08889			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-16092026-9629075805-20260917065047.pdf
Byrathi	HUL	9629075812	16/09/2026	16/09/2026	GRN05889	PO-RP-BLR-HMR11248	PO-ASN-HMR-000010877	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-16092026-9629075812-20260916192951.pdf
Byrathi	HUL	9629075813	16/09/2026	16/09/2026	GRN05890	PO-RP-BLR-HMR11246	PO-ASN-HMR-000010875	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-16092026-9629075813-20260916193141.pdf
Byrathi	HUL	9629075814	16/09/2026	16/09/2026	GRN05891	PO-RP-BLR-HMR11244	PO-ASN-HMR-000010873	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/BYTI-HUL-16092026-9629075814-20260916193256.pdf
Yeshwantpura	HUL	9629075822	16/09/2026	16/09/2026	GIR06879			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-16092026-9629075822-20260917080716.pdf
Yeshwantpura	HUL	9629075823	16/09/2026	16/09/2026	GIR06880			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-16092026-9629075823-20260917080753.pdf
Yeshwantpura	HUL	9629075824	16/09/2026	16/09/2026	GIR06878			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-16092026-9629075824-20260917080815.pdf
Pondicherry	HUL	9633123460	15/09/2026	15/09/2026	PUR06779			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/PNCY-HUL-15092026-9633123460-20260916091017.pdf
Yeshwantpura	HUL	9633124384	15/09/2026	15/09/2026	GIR06873			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/YSPR-HUL-15092026-9633124384-20260915210549.pdf
Hosur	HUL	9633124470	16/09/2026	16/09/2026	PRB00368	PO-RP-BLR-HSRA04455	PO-ASN-HSRA-000007223	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/HSRA-HUL-16092026-9633124470-20260916132546.pdf
Hosur	HUL	9633124471	16/09/2026	16/09/2026	PRB00369	PO-RP-BLR-HSRA04453	PO-ASN-HSRA-000007221	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/HSRA-HUL-16092026-9633124471-20260916132727.pdf
Hosur	HUL	9633124472	16/09/2026	16/09/2026	PRB00370	PO-RP-BLR-HSRA04454	PO-ASN-HSRA-000007222	pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/HSRA-HUL-16092026-9633124472-20260916132604.pdf
Pondicherry	HUL	9633124575	15/09/2026	15/09/2026	PUR06782			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/PNCY-HUL-15092026-9633124575-20260916091125.pdf
Pondicherry	HUL	9633124576	15/09/2026	15/09/2026	PUR06781			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/PNCY-HUL-15092026-9633124576-20260916091141.pdf
Pondicherry	HUL	9633124577	15/09/2026	15/09/2026	PUR06783			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/PNCY-HUL-15092026-9633124577-20260916091102.pdf
Pondicherry	HUL	9633124578	15/09/2026	15/09/2026	PUR06780			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/PNCY-HUL-15092026-9633124578-20260916091156.pdf
Chromepet	HUL SAMADHAN	9633124897	13/09/2026	13/09/2026	PUR02889			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124897-20260915170109.pdf
Chromepet	HUL SAMADHAN	9633124898	13/09/2026	13/09/2026	PUR02890			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124898-20260915170053.pdf
Chromepet	HUL SAMADHAN	9633124899	13/09/2026	13/09/2026	PUR02891			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124899-20260915170034.pdf
Chromepet	HUL SAMADHAN	9633124903	13/09/2026	13/09/2026	PUR02892			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124903-20260915170011.pdf
Chromepet	HUL SAMADHAN	9633124904	13/09/2026	13/09/2026	PUR02893			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124904-20260915170243.pdf
Chromepet	HUL SAMADHAN	9633124905	13/09/2026	13/09/2026	PUR02894			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124905-20260915170224.pdf
Chromepet	HUL SAMADHAN	9633124949	13/09/2026	13/09/2026	PUR02895			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124949-20260915170206.pdf
Chromepet	HUL SAMADHAN	9633124950	13/09/2026	13/09/2026	PUR02896			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124950-20260915170149.pdf
Chromepet	HUL SAMADHAN	9633124951	13/09/2026	13/09/2026	PUR02897			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-13092026-9633124951-20260915170131.pdf
OMR	HUL SAMADHAN	9633124967	13/09/2026	13/09/2026	PUR00932			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/OMR-HULS-13092026-9633124967-20260915145149.pdf
OMR	HUL SAMADHAN	9633124968	13/09/2026	13/09/2026	PUR00933			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/OMR-HULS-13092026-9633124968-20260915145204.pdf
OMR	HUL SAMADHAN	9633124969	13/09/2026	13/09/2026	PUR00934			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/OMR-HULS-13092026-9633124969-20260915145232.pdf
OMR	HUL SAMADHAN	9633124976	13/09/2026	13/09/2026	PUR00935			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/OMR-HULS-13092026-9633124976-20260915145247.pdf
OMR	HUL SAMADHAN	9633124977	13/09/2026	13/09/2026	PUR00936			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/OMR-HULS-13092026-9633124977-20260915145259.pdf
OMR	HUL SAMADHAN	9633124978	13/09/2026	13/09/2026	PUR00937			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/OMR-HULS-13092026-9633124978-20260915145311.pdf
OMR	HUL SAMADHAN	9633125695	15/09/2026	15/09/2026	PUR00938				
OMR	HUL SAMADHAN	9633125696	15/09/2026	15/09/2026	PUR00939				
OMR	HUL SAMADHAN	9633125697	15/09/2026	15/09/2026	PUR00940				
Chromepet	HUL SAMADHAN	9633125710	15/09/2026	15/09/2026	PUR02898			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-15092026-9633125710-20260916145007.pdf
Chromepet	HUL SAMADHAN	9633125711	15/09/2026	15/09/2026	PUR02899			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-15092026-9633125711-20260916144949.pdf
Chromepet	HUL SAMADHAN	9633125712	15/09/2026	15/09/2026	PUR02900			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-15092026-9633125712-20260916144931.pdf
Chromepet	HUL SAMADHAN	9633125716	15/09/2026	15/09/2026	PUR02901			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-15092026-9633125716-20260916144850.pdf
Chromepet	HUL SAMADHAN	9633125717	15/09/2026	15/09/2026	PUR02902			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-15092026-9633125717-20260916144834.pdf
Chromepet	HUL SAMADHAN	9633125718	15/09/2026	15/09/2026	PUR02903			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/CRMP-HULS-15092026-9633125718-20260916144737.pdf
OMR	HUL SAMADHAN	9633125767	15/09/2026	15/09/2026	PUR00941				
OMR	HUL SAMADHAN	9633125768	15/09/2026	15/09/2026	PUR00942				
OMR	HUL SAMADHAN	9633125769	15/09/2026	15/09/2026	PUR00943				
Mysore Road	HUL	9633126062	16/09/2026	16/09/2026	PUR08886			pdf	https://cdms-signed-invoice.s3.ap-south-1.amazonaws.com/prod/epod/grn/TLBL-HUL-16092026-9633126062-20260917064906.pdf
Chromepet	HUL SAMADHAN	9633126422	16/09/2026	16/09/2026	PUR02904				
Chromepet	HUL SAMADHAN	9633126423	16/09/2026	16/09/2026	PUR02905				
Chromepet	HUL SAMADHAN	9633126424	16/09/2026	16/09/2026	PUR02906				
Chromepet	HUL SAMADHAN	9633126428	16/09/2026	16/09/2026	PUR02907				
Chromepet	HUL SAMADHAN	9633126429	16/09/2026	16/09/2026	PUR02908				
Chromepet	HUL SAMADHAN	9633126430	16/09/2026	16/09/2026	PUR02909				
Chromepet	HUL SAMADHAN	9633126434	16/09/2026	16/09/2026	PUR02910				
Chromepet	HUL SAMADHAN	9633126435	16/09/2026	16/09/2026	PUR02911				
Chromepet	HUL SAMADHAN	9633126436	16/09/2026	16/09/2026	PUR02912				
OMR	HUL SAMADHAN	9633126455	16/09/2026	16/09/2026	PUR00947				
OMR	HUL SAMADHAN	9633126456	16/09/2026	16/09/2026	PUR00948				
OMR	HUL SAMADHAN	9633126457	16/09/2026	16/09/2026	PUR00949				
OMR	HUL SAMADHAN	9633126479	16/09/2026	16/09/2026	PUR00944				
OMR	HUL SAMADHAN	9633126480	16/09/2026	16/09/2026	PUR00945				
OMR	HUL SAMADHAN	9633126481	16/09/2026	16/09/2026	PUR00946				
OMR	HUL SAMADHAN	9633126594	16/09/2026	16/09/2026	PUR00950				
OMR	HUL SAMADHAN	9633126595	16/09/2026	16/09/2026	PUR00951				
OMR	HUL SAMADHAN	9633126596	16/09/2026	16/09/2026	PUR00952				
"""
# ---------------------------------------------------------


def parse(url):
    # filename: <LOC>-<BRAND>-<DDMMYYYY>-<INVOICE>-<timestamp>.pdf
    p = os.path.basename(url).rsplit(".", 1)[0].split("-")
    loc = LOC.get(p[0].upper(), p[0]) if p else "Unknown"
    brand = BRAND.get(p[1].upper(), p[1]) if len(p) > 1 else "HUL"
    d = p[2] if len(p) > 2 else ""
    date = f"{d[4:]}-{d[2:4]}-{d[0:2]}" if len(d) == 8 and d.isdigit() else "unknown-date"
    inv = p[3] if len(p) > 3 else os.path.basename(url)
    return loc, date, inv, brand


def urls():
    return list(dict.fromkeys(URL_RE.findall(PASTE)))  # de-dupe, keep order


def load_booked():
    # invoices already booked in NetSuite (written by mark_booked.py) -> never re-download.
    p = os.path.join(BOOK_ROOT, "_booked.csv")
    booked = set()
    if os.path.exists(p):
        with open(p, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                booked.add(str(row["invoice_no"]).strip())
    return booked


def write_source_and_build(loc, rows):
    """Write <loc>_source.csv and build the fillable <loc>_Entry.xlsx."""
    os.makedirs(os.path.join(BOOK_ROOT, loc), exist_ok=True)
    csv_path = os.path.join(BOOK_ROOT, loc, f"{loc}_source.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["fc_name", "brand_name", "invoice_no", "invoice_date",
                    "brand_grn_date", "brand_grn_no", "wms_po_number",
                    "wms_asn_number", "file_type", "s3_file_link"])
        for loc_, date, inv, brand, url in rows:
            w.writerow([loc_, brand, inv, date, date, "", "", "", "pdf", url])
    print(f"[CSV]  {csv_path}")
    r = subprocess.run([sys.executable, "build_entry_workbook.py", loc, csv_path],
                       cwd=SCRIPT_DIR)
    if r.returncode:
        print(f"[WARN] Could not build Excel for '{loc}'. Add its config to "
              f"build_entry_workbook.py (HUBS), then rerun. PDFs are already downloaded.")


def selftest():
    loc, date, inv, brand = parse("https://x/YSPR-HUL-07092026-9629071933-20260909080204.pdf")
    assert (loc, date, inv, brand) == ("Yeshwantpura", "2026-09-07", "9629071933", "HUL"), \
        (loc, date, inv, brand)
    assert len(urls()) == 4, urls()
    print("selftest OK:", loc, date, inv, brand, "| links found:", len(urls()))


def main():
    if "--selftest" in sys.argv:
        return selftest()
    found = urls()
    assert found, "No links found - paste rows between the PASTE markers first."
    booked = load_booked()
    ok = skipped = failed = already = 0
    by_loc = {}   # loc -> list of (loc, date, inv, brand, url) for the workbook
    for url in found:
        loc, date, inv, brand = parse(url)
        if str(inv).strip() in booked:      # already imported to NetSuite -> never again
            already += 1
            print(f"[ALREADY BOOKED] {loc} / {date} / {inv} - skipped")
            continue
        by_loc.setdefault(loc, []).append((loc, date, inv, brand, url))
        folder = os.path.join(BOOK_ROOT, loc, "PDFs", date)
        os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, os.path.basename(url))
        if os.path.exists(target):
            skipped += 1
            continue
        try:
            urllib.request.urlretrieve(url, target)
            ok += 1
            print(f"[OK]  {loc} / {date} / {os.path.basename(url)}")
        except Exception as e:
            failed += 1
            print(f"[FAIL] {url}\n       {e}")
    print(f"\nDownloaded={ok}  Skipped(existing)={skipped}  AlreadyBooked={already}  Failed={failed}")
    print(f"Saved under: {BOOK_ROOT}\\<Location>\\PDFs\\<date>\n")
    assert failed == 0, f"{failed} download(s) failed - see [FAIL] lines above"
    for loc, rows in by_loc.items():
        write_source_and_build(loc, rows)


if __name__ == "__main__":
    main()

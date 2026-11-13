# -*- coding: utf-8 -*-
"""
Stock_Inventory v2 — Premium Business Edition
Dark navy branding, gold accents, full KPIs, polished every sheet.
Built for a senior business owner who runs his own wholesale operation.
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.formatting import Rule

# ── BRAND COLORS ─────────────────────────────────────────────────────────────
NAVY      = "FF1B2A4A"   # deep navy – primary brand
NAVY_MID  = "FF2C3E6B"   # mid navy – secondary
GOLD      = "FFC9A84C"   # gold accent
GOLD_LT   = "FFFFF3CD"   # light gold – highlights
WHITE     = "FFFFFFFF"
OFF_WHITE = "FFF8F9FC"   # near-white background
SLATE     = "FF34495E"   # slate for secondary text
GREEN_OK  = "FFC6EFCE"
GREEN_DT  = "FF6AA84F"
YELLOW_LT = "FFFFF2CC"
YELLOW_DT = "FFFFC000"
RED_LT    = "FFFFC7CE"
RED_DT    = "FFF4CCCC"
RED_DARK  = "FFF44336"
GREEN_LT  = "FFC6EFCE"
GREY_BDR  = "FFD0D7E2"
GREY_LT   = "FFF2F4F7"

THIN   = Side(style="thin",   color=GREY_BDR)
MED    = Side(style="medium", color=NAVY)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_BOX = Border(left=MED, right=MED, top=MED, bottom=MED)

HDR_FILL   = PatternFill("solid", fgColor=NAVY)
HDR_FILL2  = PatternFill("solid", fgColor=NAVY_MID)
HDR_FONT   = Font(name="Calibri", bold=True, color=WHITE, size=11)
SUB_FONT   = Font(name="Calibri", bold=False, color=GOLD, size=10)
ROW_FONT   = Font(name="Calibri", size=11, color="FF2C3E50")
ROW_FONT_B = Font(name="Calibri", size=11, bold=True, color="FF2C3E50")
KPI_VAL_F  = Font(name="Calibri", bold=True, size=18, color=NAVY)
KPI_LBL_F  = Font(name="Calibri", bold=True, size=10, color=SLATE)
TITLE_F    = Font(name="Calibri", bold=True, size=20, color=NAVY)
SUBTITLE_F = Font(name="Calibri", size=12, color=SLATE, italic=True)
SECTION_F  = Font(name="Calibri", bold=True, size=13, color=WHITE)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT  = Alignment(horizontal="right",  vertical="center")
WRAP_L = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

# ════════════════════════════════════════════════════════════════════════════
# DATA
# ════════════════════════════════════════════════════════════════════════════
ITEMS = [
    ("Upper Body",       "Upper باڈی",       "ABC Traders"),
    ("Cat",              "کیٹ",              "XYZ Wholesale"),
    ("Bolt Jora",        "بلوٹ جوڑا",        "ABC Traders"),
    ("Barrel",           "بیرل",             "Local Market"),
    ("Zaha",             "زخہ",              "XYZ Wholesale"),
    ("Poungee",          "پونگی",            "ABC Traders"),
    ("Gas Pipe",         "گیس پائپ",         "Gas Co. Ltd"),
    ("Silver Ring",      "سلور رینگ",        "Jewellery Sup."),
    ("Jali Cutter",      "جالی کٹر",         "ABC Traders"),
    ("Fitting Nut",      "فٹنگ ناٹ",         "ABC Traders"),
    ("Chalan Saman Moti Wala", "چلن سامان موٹی والا", "XYZ Wholesale"),
    ("Chota Packet",     "چھوٹا پاکٹ",       "ABC Traders"),
    ("Kashkonki",        "کشکونکے",         "XYZ Wholesale"),
    ("Push Button",      "پوش بٹن",          "ABC Traders"),
    ("Karki",            "کھڑکی",            "Local Market"),
    ("Upper Handle",     "Upper ہینڈل",      "ABC Traders"),
    ("Motay Agay Wala",  "موٹے آگے والا",    "ABC Traders"),
    ("Pipe Taiwan",      "پائپ تھائی وان",    "Pipe Co."),
    ("Cover",            "کاور",             "ABC Traders"),
    ("Weight",           "وزن",              "ABC Traders"),
    ("Weight Spring",    "وزن سپرنگ",        "ABC Traders"),
    ("Pipe Kari",        "پائپ کھڑے",        "Pipe Co."),
    ("Magpol Patri",     "میک پول پٹری",     "XYZ Wholesale"),
    ("Press Ring",       "پریس رینگ",        "ABC Traders"),
]

OUT  = sys.argv[1] if len(sys.argv) > 1 else "Stock_Inventory.xlsx"
TXN  = int(sys.argv[2]) if len(sys.argv) > 2 else 3000
N    = len(ITEMS)
SKUS = [f"STK{i+1:03d}" for i in range(N)]
DS   = 7    # stock data start row (KPI cards=3-4, spacer=5, header=6, data=7+)
DE   = DS + N - 1  # 30
MIN_STOCK = [30,30,40,20,20,30,30,25,15,20,20,15,10,25,20,
             20,20,30,20,20,20,15,10,15]

SUP_DB = {
    "ABC Traders":    ("0300-1234567", "abctraders@example.com",   2, "Cash on delivery"),
    "XYZ Wholesale":  ("0333-9876543", "xyzwholesale@example.com", 3, "7 days credit"),
    "Local Market":   ("0345-2223334", "",                         1, "Cash"),
    "Gas Co. Ltd":    ("0300-9998887", "gasco@example.com",        5, "Advance 50%"),
    "Jewellery Sup.": ("0348-5556667", "jewellery@example.com",    4, "Cash on delivery"),
    "Pipe Co.":       ("0300-6665544", "pipeco@example.com",       3, "7 days credit"),
}

# ════════════════════════════════════════════════════════════════════════════
wb = Workbook()

# �═══════════════════════════════════════════════════════════════════════════
# SHEET 1 — STOCK  (main inventory)
# ════════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "Stock"
ws.sheet_view.showGridLines = True
ws.sheet_properties.tabColor = NAVY

# ── Row 1: Business brand strip ──────────────────────────────────────────────
ws.merge_cells("A1:O1")
c = ws["A1"]
c.value = "📦  STOCK INVENTORY  —  Wholesale Business  —  Store Stock Record"
c.font  = Font(name="Calibri", bold=True, size=14, color=WHITE)
c.fill  = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 26

# ── Row 2: Gold accent strip ─────────────────────────────────────────────────
ws.merge_cells("A2:O2")
c2 = ws["A2"]
c2.fill = PatternFill("solid", fgColor=GOLD)
ws.row_dimensions[2].height = 4

# ── Rows 3-4: KPI cards (2 rows × 6 cols = 12 KPIs in 2 rows) ─────────────
# Layout: 6 cards per row, each card = 2 cols wide (A-B, C-D, E-F, G-H, I-J, K-L)
# Row 3 = big number, Row 4 = label
kpi_data = [
    # (label, formula, merge_cols_start, merge_cols_end)
    ("Total\nItems",        f"=COUNTA(A{DS}:A{DE})",          1, 2),
    ("In\nStock",           f'=COUNTIF(H{DS}:H{DE},">0")',    3, 4),
    ("Low\nStock",          f'=COUNTIF(M{DS}:M{DE},"Low")',   5, 6),
    ("Out of\nStock",       f'=COUNTIF(M{DS}:M{DE},"Out of stock")', 7, 8),
    ("Check\nEntries",      f'=COUNTIF(M{DS}:M{DE},"Check entries")', 9, 10),
    ("Total Stock\nValue Rs",f"=SUMPRODUCT((H{DS}:H{DE}>0)*(H{DS}:H{DE})*(J{DS}:J{DE}))", 11, 12),
]
for label, formula, cstart, cend in kpi_data:
    # Big number cell
    ws.merge_cells(start_row=3, start_column=cstart, end_row=3, end_column=cend)
    tl = ws.cell(row=3, column=cstart)
    tl.value     = formula
    tl.font      = Font(name="Calibri", bold=True, size=22, color=NAVY)
    tl.fill      = PatternFill("solid", fgColor=GOLD_LT)
    tl.alignment = CENTER
    tl.border    = BORDER_BOX
    # Label cell (row 4)
    ws.merge_cells(start_row=4, start_column=cstart, end_row=4, end_column=cend)
    ll = ws.cell(row=4, column=cstart)
    ll.value     = label
    ll.font      = Font(name="Calibri", bold=True, size=10, color=NAVY)
    ll.fill      = PatternFill("solid", fgColor=OFF_WHITE)
    ll.alignment = CENTER
    ll.border    = BORDER
ws.row_dimensions[3].height = 34
ws.row_dimensions[4].height = 22

# ── Row 5: spacer ────────────────────────────────────────────────────────────
ws.row_dimensions[5].height = 6

# ── Row 6: Column headers ────────────────────────────────────────────────────
headers = [
    "Item\nCode", "Item\nName", "Urdu\nName", "Unit",
    "Opening\nStock", "Stock\nIn", "Stock\nOut", "Current\nStock",
    "Min\nStock", "Cost\nPrice\n(Rs)", "Sell\nPrice\n(Rs)",
    "Supplier", "Status", "Reorder\nQty", "Notes"
]
for i, h in enumerate(headers, 1):
    c = ws.cell(row=6, column=i, value=h)
    c.font      = Font(name="Calibri", bold=True, color=WHITE, size=10)
    c.fill      = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER
    c.border    = BORDER
ws.row_dimensions[6].height = 32

# Column widths
for i, w in enumerate([10, 22, 20, 7, 12, 10, 10, 12,
                        10, 12, 12, 16, 13, 11, 18], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ── Data rows 7–30 ────────────────────────────────────────────────────────────
for idx, (eng, urd, sup) in enumerate(ITEMS):
    r = DS + idx
    row_fill = GREY_LT if idx % 2 == 0 else WHITE

    ws.cell(row=r, column=1,  value=SKUS[idx])
    ws.cell(row=r, column=2,  value=eng)
    ws.cell(row=r, column=3,  value=urd)
    ws.cell(row=r, column=4,  value="Pcs")
    ws.cell(row=r, column=5,  value=200)
    ws.cell(row=r, column=6,  value=(
        f"=SUMIF('Stock In'!$B$2:$B${TXN+1},$B{r},"
        f"'Stock In'!$D$2:$D${TXN+1})"))
    ws.cell(row=r, column=7,  value=(
        f"=SUMIF('Stock Out'!$B$2:$B${TXN+1},$B{r},"
        f"'Stock Out'!$D$2:$D${TXN+1})"))
    ws.cell(row=r, column=8,  value=(
        f"=IFERROR($E{r}+$F{r}-$G{r},\"Check entries\")"))
    ws.cell(row=r, column=9,  value=MIN_STOCK[idx])
    ws.cell(row=r, column=10, value=None)   # Cost Price — blank
    ws.cell(row=r, column=11, value=None)   # Selling Price — blank
    ws.cell(row=r, column=12, value=sup)
    ws.cell(row=r, column=13, value=(
        f'=IF($H{r}<0,"Check entries",'
        f'IF($H{r}=0,"Out of stock",'
        f'IF(AND($I{r}<>"",$H{r}<=$I{r}),"Low","OK")))'))
    ws.cell(row=r, column=14, value=(
        f'=IF(OR($M{r}="OK",$M{r}=""),"",MAX(0,$I{r}*2-$H{r}))'))
    ws.cell(row=r, column=15, value="")

    for c in range(1, 16):
        cell = ws.cell(row=r, column=c)
        cell.font      = ROW_FONT
        cell.border    = BORDER
        cell.fill      = PatternFill("solid", fgColor=row_fill)
    ws.row_dimensions[r].height = 17

    # Alignments
    ws.cell(row=r, column=1).alignment  = CENTER
    ws.cell(row=r, column=2).alignment  = LEFT
    ws.cell(row=r, column=3).alignment  = RIGHT
    ws.cell(row=r, column=4).alignment  = CENTER
    for col in [5,6,7,8,9,10,11,14]:
        ws.cell(row=r, column=col).alignment = RIGHT
    ws.cell(row=r, column=12).alignment = LEFT
    ws.cell(row=r, column=13).alignment = CENTER
    ws.cell(row=r, column=15).alignment = LEFT

# ── Conditional formatting ────────────────────────────────────────────────────
# Current stock scale (green→yellow→red)
ws.conditional_formatting.add(
    f"H{DS}:H{DE}",
    ColorScaleRule(start_type="min", start_color=GREEN_DT,
                   mid_type="percentile", mid_value=50, mid_color=YELLOW_DT,
                   end_type="max", end_color=RED_DT)
)
# Status coloring
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal", formula=['"Low"'], fill=PatternFill("solid", fgColor=YELLOW_LT)))
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal", formula=['"Out of stock"'], fill=PatternFill("solid", fgColor=RED_LT)))
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal", formula=['"Check entries"'], fill=PatternFill("solid", fgColor=RED_LT)))
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", fgColor=GREEN_LT)))
# Highlight row when out of stock
ws.conditional_formatting.add(f"A{DS}:O{DE}",
    FormulaRule(formula=[f'$M{DS}="Out of stock"'],
                fill=PatternFill("solid", fgColor="FFFCE4EC")))

ws.freeze_panes   = "A7"
ws.auto_filter.ref = f"A6:O{DE}"

# ════════════════════════════════════════════════════════════════════════════
# SHEET 2 — STOCK IN
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Stock In")
ws2.sheet_properties.tabColor = "FF2E7D32"

ws2.merge_cells("A1:H1")
c = ws2["A1"]
c.value = "STOCK IN  —  Deliveries & Receiving Record"
c.font  = Font(name="Calibri", bold=True, size=13, color=WHITE)
c.fill  = PatternFill("solid", fgColor="FF2E7D32")
c.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 24

ws2.merge_cells("A2:H2")
ws2["A2"].fill = PatternFill("solid", fgColor=GOLD)
ws2.row_dimensions[2].height = 3

in_hdr = ["Date", "Item Name\n(Select from list)", "Item Code\n(Auto)", "Qty\nIn",
          "Unit", "Supplier /\nDelivery From", "Reference\n(Invoice #)", "Notes"]
for i, h in enumerate(in_hdr, 1):
    c = ws2.cell(row=3, column=i, value=h)
    c.font      = Font(name="Calibri", bold=True, color=WHITE, size=10)
    c.fill      = PatternFill("solid", fgColor="FF2E7D32")
    c.alignment = CENTER
    c.border    = BORDER
ws2.row_dimensions[3].height = 30

for i, w in enumerate([12, 22, 10, 8, 7, 18, 16, 18], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

dv2 = DataValidation(type="list",
    formula1=f"='Stock'!$B${DS}:$B${DE}", allow_blank=True)
dv2.showInputMessage  = True
dv2.promptTitle       = "Select Item"
dv2.prompt            = "Choose the item name from the dropdown list."
ws2.add_data_validation(dv2)
dv2.add(f"B4:B{TXN+3}")

for r in range(4, TXN + 4):
    ws2.cell(row=r, column=3,
             value=(f'=IF($B{r}="","",'
                    f'INDEX(\'Stock\'!$A${DS}:$A${DE},'
                    f'MATCH($B{r},\'Stock\'!$B${DS}:$B${DE},0)))'))
    row_fill = GREY_LT if (r % 2 == 0) else WHITE
    for c in range(1, 9):
        cell = ws2.cell(row=r, column=c)
        cell.font      = ROW_FONT
        cell.border    = BORDER
        cell.fill      = PatternFill("solid", fgColor=row_fill)
        cell.alignment = LEFT
    ws2.cell(row=r, column=5, value="Pcs")
    ws2.cell(row=r, column=5).alignment = CENTER
    ws2.cell(row=r, column=3).alignment = CENTER
    ws2.row_dimensions[r].height = 16

_sample_in = [
    ("2026-11-03","Upper Body",  50,"ABC Traders",   "Weekly delivery",""),
    ("2026-11-04","Cat",        100,"XYZ Wholesale", "Restock",        ""),
    ("2026-11-05","Bolt Jora",   80,"ABC Traders",   "Weekly delivery",""),
    ("2026-11-03","Gas Pipe",    60,"Gas Co. Ltd",   "New stock",      ""),
    ("2026-11-06","Silver Ring",120,"Jewellery Sup.","Restock",        ""),
]
for i, (dt, item, qty, sup, ref, note) in enumerate(_sample_in):
    r = 4 + i
    ws2.cell(row=r, column=1, value=dt)
    ws2.cell(row=r, column=2, value=item)
    ws2.cell(row=r, column=4, value=qty)
    ws2.cell(row=r, column=6, value=sup)
    ws2.cell(row=r, column=7, value=ref)
    ws2.cell(row=r, column=8, value=note)
    for c in range(1, 9):
        ws2.cell(row=r, column=c).font      = ROW_FONT_B
        ws2.cell(row=r, column=c).fill      = PatternFill("solid", fgColor=GOLD_LT)
        ws2.cell(row=r, column=c).border    = BORDER
        ws2.cell(row=r, column=c).alignment = LEFT
    ws2.cell(row=r, column=1).alignment = CENTER
    ws2.cell(row=r, column=3).alignment = CENTER
    ws2.cell(row=r, column=4).alignment = RIGHT
    ws2.row_dimensions[r].height = 17

ws2.freeze_panes   = "A4"
ws2.auto_filter.ref = f"A3:H{TXN+3}"

# ════════════════════════════════════════════════════════════════════════════
# SHEET 3 — STOCK OUT
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Stock Out")
ws3.sheet_properties.tabColor = "FFC0392B"

ws3.merge_cells("A1:H1")
c = ws3["A1"]
c.value = "STOCK OUT  —  Sales & Issues Record"
c.font  = Font(name="Calibri", bold=True, size=13, color=WHITE)
c.fill  = PatternFill("solid", fgColor="FFC0392B")
c.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 24

ws3.merge_cells("A2:H2")
ws3["A2"].fill = PatternFill("solid", fgColor=GOLD)
ws3.row_dimensions[2].height = 3

out_hdr = ["Date", "Item Name\n(Select from list)", "Item Code\n(Auto)",
           "Qty\nOut", "Unit", "Customer /\nDestination", "Reference\n(Order #)",
           "Notes"]
for i, h in enumerate(out_hdr, 1):
    c = ws3.cell(row=3, column=i, value=h)
    c.font      = Font(name="Calibri", bold=True, color=WHITE, size=10)
    c.fill      = PatternFill("solid", fgColor="FFC0392B")
    c.alignment = CENTER
    c.border    = BORDER
ws3.row_dimensions[3].height = 30

for i, w in enumerate([12, 22, 10, 8, 7, 18, 16, 18], 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

dv3 = DataValidation(type="list",
    formula1=f"='Stock'!$B${DS}:$B${DE}", allow_blank=True)
dv3.showInputMessage  = True
dv3.promptTitle       = "Select Item"
dv3.prompt            = "Choose the item name from the dropdown list."
ws3.add_data_validation(dv3)
dv3.add(f"B4:B{TXN+3}")

for r in range(4, TXN + 4):
    ws3.cell(row=r, column=3,
             value=(f'=IF($B{r}="","",'
                    f'INDEX(\'Stock\'!$A${DS}:$A${DE},'
                    f'MATCH($B{r},\'Stock\'!$B${DS}:$B${DE},0)))'))
    row_fill = GREY_LT if (r % 2 == 0) else WHITE
    for c in range(1, 9):
        cell = ws3.cell(row=r, column=c)
        cell.font      = ROW_FONT
        cell.border    = BORDER
        cell.fill      = PatternFill("solid", fgColor=row_fill)
        cell.alignment = LEFT
    ws3.cell(row=r, column=5, value="Pcs")
    ws3.cell(row=r, column=5).alignment = CENTER
    ws3.cell(row=r, column=3).alignment = CENTER
    ws3.row_dimensions[r].height = 16

_sample_out = [
    ("2026-11-03","Upper Body",          120,"Shop A",     "Sale",        ""),
    ("2026-11-04","Cat",                 200,"Customer X", "Order #112",  ""),
    ("2026-11-05","Bolt Jora",           200,"Customer Y", "Order #113",  ""),
    ("2026-11-06","Barrel",              180,"ABC Traders","Order #114",  ""),
    ("2026-11-06","Zaha",                190,"Shop B",     "Sale",        ""),
    ("2026-11-07","Poungee",             200,"Customer Z", "Order #115",  ""),
    ("2026-11-07","Gas Pipe",            180,"Plumber P",  "Order #116",  ""),
    ("2026-11-08","Silver Ring",         200,"Shop C",     "Sale",        ""),
    ("2026-11-08","Jali Cutter",         190,"Customer J", "Order #117",  ""),
    ("2026-11-09","Fitting Nut",         180,"ABC Traders","Order #118",  ""),
    ("2026-11-09","Chalan Saman Moti Wala",150,"Customer M","Order #119", ""),
    ("2026-11-10","Chota Packet",        200,"Shop D",     "Sale",        ""),
    ("2026-11-10","Push Button",         180,"Customer P", "Order #120",  ""),
    ("2026-11-11","Karki",               120,"Customer K", "Order #121",  ""),
    ("2026-11-11","Motay Agay Wala",     180,"Shop E",     "Sale",        ""),
    ("2026-11-12","Pipe Taiwan",         195,"Pipe Co.",   "Order #122",  ""),
    ("2026-11-12","Cover",               170,"ABC Traders","Order #123",  ""),
    ("2026-11-13","Weight Spring",       190,"Customer W", "Order #124",  ""),
    ("2026-11-13","Pipe Kari",           185,"Shop F",     "Sale",        ""),
    ("2026-11-14","Magpol Patri",        195,"Customer M", "Order #125",  ""),
    ("2026-11-14","Press Ring",          185,"ABC Traders","Order #126",  ""),
]
for i, (dt, item, qty, cust, ref, note) in enumerate(_sample_out):
    r = 4 + i
    ws3.cell(row=r, column=1, value=dt)
    ws3.cell(row=r, column=2, value=item)
    ws3.cell(row=r, column=4, value=qty)
    ws3.cell(row=r, column=6, value=cust)
    ws3.cell(row=r, column=7, value=ref)
    ws3.cell(row=r, column=8, value=note)
    for c in range(1, 9):
        ws3.cell(row=r, column=c).font      = ROW_FONT_B
        ws3.cell(row=r, column=c).fill      = PatternFill("solid", fgColor=YELLOW_LT)
        ws3.cell(row=r, column=c).border    = BORDER
        ws3.cell(row=r, column=c).alignment = LEFT
    ws3.cell(row=r, column=1).alignment = CENTER
    ws3.cell(row=r, column=3).alignment = CENTER
    ws3.cell(row=r, column=4).alignment = RIGHT
    ws3.row_dimensions[r].height = 17

ws3.freeze_panes   = "A4"
ws3.auto_filter.ref = f"A3:H{TXN+3}"

# ════════════════════════════════════════════════════════════════════════════
# SHEET 4 — STOCK ADJUSTMENTS
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Stock Adjustments")
ws4.sheet_properties.tabColor = "FF8E44AD"

ws4.merge_cells("A1:H1")
c = ws4["A1"]
c.value = "STOCK ADJUSTMENTS  —  Non-Sale Changes (Damage, Returns, Found)"
c.font  = Font(name="Calibri", bold=True, size=12, color=WHITE)
c.fill  = PatternFill("solid", fgColor="FF8E44AD")
c.alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[1].height = 22

ws4.merge_cells("A2:H2")
ws4["A2"].fill = PatternFill("solid", fgColor=GOLD)
ws4.row_dimensions[2].height = 3

adj_hdr = ["Date", "Item Name\n(Select from list)", "Item Code\n(Auto)",
           "Qty", "Adjustment\nType", "Reason\n(Why?)",
           "Reference", "Entered\nBy"]
for i, h in enumerate(adj_hdr, 1):
    c = ws4.cell(row=3, column=i, value=h)
    c.font      = Font(name="Calibri", bold=True, color=WHITE, size=10)
    c.fill      = PatternFill("solid", fgColor="FF8E44AD")
    c.alignment = CENTER
    c.border    = BORDER
ws4.row_dimensions[3].height = 30

for i, w in enumerate([12, 22, 10, 8, 18, 22, 16, 14], 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

dv4a = DataValidation(type="list",
    formula1=f"='Stock'!$B${DS}:$B${DE}", allow_blank=True)
dv4a.showInputMessage = True
dv4a.promptTitle      = "Select Item"
dv4a.prompt           = "Choose the item name."
ws4.add_data_validation(dv4a)
dv4a.add(f"B4:B{TXN+3}")

adj_types = ["Damage","Return to Supplier","Return from Customer",
             "Found Missing","Other"]
dv4b = DataValidation(type="list",
    formula1=f'="{",".join(adj_types)}"', allow_blank=True)
dv4b.showInputMessage = True
dv4b.promptTitle      = "Adjustment Type"
dv4b.prompt           = "Select the type of adjustment."
ws4.add_data_validation(dv4b)
dv4b.add(f"E4:E{TXN+3}")

for r in range(4, TXN + 4):
    ws4.cell(row=r, column=3,
             value=(f'=IF($B{r}="","",'
                    f'INDEX(\'Stock\'!$A${DS}:$A${DE},'
                    f'MATCH($B{r},\'Stock\'!$B${DS}:$B${DE},0)))'))
    row_fill = GREY_LT if (r % 2 == 0) else WHITE
    for c in range(1, 9):
        cell = ws4.cell(row=r, column=c)
        cell.font      = ROW_FONT
        cell.border    = BORDER
        cell.fill      = PatternFill("solid", fgColor=row_fill)
        cell.alignment = LEFT
    ws4.cell(row=r, column=4, value="Pcs")
    ws4.cell(row=r, column=4).alignment = CENTER
    ws4.cell(row=r, column=3).alignment = CENTER
    ws4.row_dimensions[r].height = 16

ws4.freeze_panes   = "A4"
ws4.auto_filter.ref = f"A3:H{TXN+3}"

# ════════════════════════════════════════════════════════════════════════════
# SHEET 5 — SUPPLIERS
# ════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("Suppliers")
ws5.sheet_properties.tabColor = "FF2C3E50"

ws5.merge_cells("A1:G1")
c = ws5["A1"]
c.value = "SUPPLIERS DIRECTORY  —  Who Supplies What"
c.font  = Font(name="Calibri", bold=True, size=13, color=WHITE)
c.fill  = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws5.row_dimensions[1].height = 24

ws5.merge_cells("A2:G2")
ws5["A2"].fill = PatternFill("solid", fgColor=GOLD)
ws5.row_dimensions[2].height = 3

sup_hdr = ["Supplier\nName", "Phone /\nContact", "Email",
           "Items\nSupplied (auto)", "Lead Time\n(days)",
           "Payment\nTerms", "Notes"]
for i, h in enumerate(sup_hdr, 1):
    c = ws5.cell(row=3, column=i, value=h)
    c.font      = Font(name="Calibri", bold=True, color=WHITE, size=10)
    c.fill      = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER
    c.border    = BORDER
ws5.row_dimensions[3].height = 30

for i, w in enumerate([20, 16, 20, 28, 12, 16, 20], 1):
    ws5.column_dimensions[get_column_letter(i)].width = w

seen = set()
sup_list = []
for _, _, s in ITEMS:
    if s not in seen:
        seen.add(s)
        sup_list.append(s)

for i, sup in enumerate(sup_list):
    r = 4 + i
    info = SUP_DB.get(sup, ("", "", "", ""))
    ws5.cell(row=r, column=1, value=sup)
    ws5.cell(row=r, column=2, value=info[0])
    ws5.cell(row=r, column=3, value=info[1])
    ws5.cell(row=r, column=4,
             value=(f'=TEXTJOIN(", ",TRUE,'
                    f'FILTER(\'Stock\'!$B${DS}:$B${DE},'
                    f'\'Stock\'!$L${DS}:$L${DE}=$A{r},""))'))
    ws5.cell(row=r, column=5, value=info[2])
    ws5.cell(row=r, column=6, value=info[3])
    ws5.cell(row=r, column=7, value="")

    row_fill = GREY_LT if (i % 2 == 0) else WHITE
    for c in range(1, 8):
        cell = ws5.cell(row=r, column=c)
        cell.font      = ROW_FONT
        cell.border    = BORDER
        cell.fill      = PatternFill("solid", fgColor=row_fill)
        cell.alignment = LEFT
    ws5.cell(row=r, column=1).alignment = LEFT
    ws5.cell(row=r, column=4).alignment = LEFT
    ws5.cell(row=r, column=5).alignment = CENTER
    ws5.row_dimensions[r].height = 17

# Highlight the supplier name column
ws5.conditional_formatting.add(f"A4:A{3+len(sup_list)}",
    CellIsRule(operator="notEqual", formula=['""'],
               fill=PatternFill("solid", fgColor=GOLD_LT)))

ws5.freeze_panes = "A4"

# ════════════════════════════════════════════════════════════════════════════
# SHEET 6 — DASHBOARD
# ════════════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Dashboard")
ws6.sheet_view.showGridLines = False
ws6.sheet_properties.tabColor = GOLD

# ── Title section ─────────────────────────────────────────────────────────────
ws6.merge_cells("B2:L2")
c = ws6["B2"]
c.value = "STOCK DASHBOARD"
c.font  = Font(name="Calibri", bold=True, size=22, color=WHITE)
c.fill  = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
ws6.row_dimensions[2].height = 32

ws6.merge_cells("B3:L3")
c3 = ws6["B3"]
c3.value = "Wholesale Business  —  Inventory Status & Action Summary"
c3.font  = Font(name="Calibri", size=11, color=GOLD, italic=True)
c3.fill  = PatternFill("solid", fgColor=NAVY)
c3.alignment = Alignment(horizontal="center", vertical="center")
ws6.row_dimensions[3].height = 16

ws6.row_dimensions[4].height = 6  # spacer

# ── Row 5-6: 6 big KPI cards ─────────────────────────────────────────────────
# Each card = 2 columns wide, 2 rows tall (number + label)
# Cols: B-C, D-E, F-G, H-I, J-K, L-M → but we use B-L (11 cols)
# Using cols B(2) to L(12), each card = 1 col wide to fit 6 cards in 11 cols
# Actually let's use B-C, D-E, F, G-H, I-J, K-L — uneven but works
# Better: use B through M, 6 cards × 2 cols = 12 cols → B:C, D:E, F:G, H:I, J:K, L:M
kpi_cards = [
    ("Total\nItems",          f"=COUNTA('Stock'!$A${DS}:$A${DE})",         "B", "C", NAVY),
    ("Items\nIn Stock",       f'=COUNTIF(\'Stock\'!$H${DS}:$H${DE},">0")', "D", "E", "FF2E7D32"),
    ("Low\nStock",            f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Low")', "F", "G", "FFC0392B"),
    ("Out of\nStock",         f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Out of stock")', "H", "I", "FF8E44AD"),
    ("OK\nItems",             f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"OK")', "J", "K", "FF2C3E50"),
    ("Stock Value\nTotal Rs", f"=SUMPRODUCT((H{DS}:H{DE}>0)*(H{DS}:H{DE})*(J{DS}:J{DE}))", "L", "M", GOLD),
]
for label, formula, c1, c2, color in kpi_cards:
    col_start = ord(c1) - ord("A") + 1
    col_end   = ord(c2) - ord("A") + 1
    # Big number (row 5)
    ws6.merge_cells(start_row=5, start_column=col_start,
                     end_row=5, end_column=col_end)
    nc = ws6.cell(row=5, column=col_start)
    nc.value     = formula
    nc.font      = Font(name="Calibri", bold=True, size=20, color=WHITE)
    nc.fill      = PatternFill("solid", fgColor=color)
    nc.alignment = CENTER
    nc.border    = BORDER_BOX
    # Label (row 6)
    ws6.merge_cells(start_row=6, start_column=col_start,
                     end_row=6, end_column=col_end)
    lc = ws6.cell(row=6, column=col_start)
    lc.value     = label
    lc.font      = Font(name="Calibri", bold=True, size=9, color=WHITE)
    lc.fill      = PatternFill("solid", fgColor=NAVY_MID)
    lc.alignment = CENTER
    lc.border    = BORDER
ws6.row_dimensions[5].height = 36
ws6.row_dimensions[6].height = 22

ws6.row_dimensions[7].height = 8  # spacer

# ── Row 8: Secondary KPIs ─────────────────────────────────────────────────────
ws6.merge_cells("B8:D8")
c8 = ws6["B8"]
c8.value = "📋  Action Summary"
c8.font  = Font(name="Calibri", bold=True, size=12, color=NAVY)
c8.fill  = PatternFill("solid", fgColor=GOLD_LT)
c8.alignment = LEFT
ws6.row_dimensions[8].height = 18

action_items = [
    (f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Out of stock")&" items out of stock — URGENT"',
     f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Low")&" items low — reorder soon"',
     f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"OK")&" items OK"',
     f'=TEXT(SUMPRODUCT((H{DS}:H{DE}>0)*(H{DS}:H{DE})*(J{DS}:J{DE})),"#,##0")&" Rs total stock value"'),
]
ws6.merge_cells("B9:D9")
ws6["B9"].value = action_items[0][0]
ws6["B9"].font  = Font(name="Calibri", size=11, color="FFC0392B", bold=True)
ws6["B9"].fill  = PatternFill("solid", fgColor=RED_LT)
ws6["B9"].fill  = PatternFill("solid", fgColor=RED_LT)
ws6["B9"].alignment = LEFT
ws6.row_dimensions[9].height = 18

ws6.merge_cells("E9:G9")
ws9b = ws6["E9"]
ws9b.value = action_items[0][1]
ws9b.font  = Font(name="Calibri", size=11, color="FF8E6B00", bold=True)
ws9b.fill  = PatternFill("solid", fgColor=YELLOW_LT)
ws9b.alignment = LEFT
ws6.row_dimensions[9].height = 18

ws6.merge_cells("H9:J9")
ws9c = ws6["H9"]
ws9c.value = action_items[0][2]
ws9c.font  = Font(name="Calibri", size=11, color="FF2E7D32", bold=True)
ws9c.fill  = PatternFill("solid", fgColor=GREEN_LT)
ws9c.alignment = LEFT
ws6.row_dimensions[9].height = 18

ws6.merge_cells("K9:M9")
ws9d = ws6["K9"]
ws9d.value = action_items[0][3]
ws9d.font  = Font(name="Calibri", size=11, color=NAVY, bold=True)
ws9d.fill  = PatternFill("solid", fgColor=OFF_WHITE)
ws9d.alignment = LEFT
ws6.row_dimensions[9].height = 18

ws6.row_dimensions[10].height = 8  # spacer

# ── Row 11+: Item status table ────────────────────────────────────────────────
ws6.merge_cells("B11:M11")
c11 = ws6["B11"]
c11.value = "ITEM STATUS TABLE  —  All 24 Items"
c11.font  = Font(name="Calibri", bold=True, size=12, color=WHITE)
c11.fill  = PatternFill("solid", fgColor=NAVY)
c11.alignment = LEFT
ws6.row_dimensions[11].height = 18

status_hdrs = [
    "Item Code", "Item Name", "Urdu Name", "Current\nStock",
    "Min\nStock", "Status", "Cost\nPrice\n(Rs)", "Sell\nPrice\n(Rs)",
    "Stock Value\n(Rs)", "Reorder\nQty", "Action\nNeeded"
]
col_widths  = [10, 22, 20, 11, 9, 12, 11, 11, 13, 11, 14]
for i, (h, w) in enumerate(zip(status_hdrs, col_widths), 1):
    c = ws6.cell(row=12, column=i, value=h)
    c.font      = Font(name="Calibri", bold=True, color=WHITE, size=10)
    c.fill      = PatternFill("solid", fgColor=NAVY_MID)
    c.alignment = CENTER
    c.border    = BORDER
    ws6.column_dimensions[get_column_letter(i)].width = w
ws6.row_dimensions[12].height = 30

for idx in range(N):
    r    = 13 + idx
    srow = DS + idx
    row_fill = GREY_LT if (idx % 2 == 0) else WHITE

    ws6.cell(row=r, column=1,  value=f"='Stock'!$A${srow}")
    ws6.cell(row=r, column=2,  value=f"='Stock'!$B${srow}")
    ws6.cell(row=r, column=3,  value=f"='Stock'!$C${srow}")
    ws6.cell(row=r, column=4,  value=f"='Stock'!$H${srow}")
    ws6.cell(row=r, column=5,  value=f"='Stock'!$I${srow}")
    ws6.cell(row=r, column=6,  value=f"='Stock'!$M${srow}")
    ws6.cell(row=r, column=7,  value=f"='Stock'!$J${srow}")
    ws6.cell(row=r, column=8,  value=f"='Stock'!$K${srow}")
    ws6.cell(row=r, column=9,  value=f"=IFERROR($D{r}*$G{r},0)")
    ws6.cell(row=r, column=10, value=f"='Stock'!$N${srow}")
    ws6.cell(row=r, column=11,
             value=f'=IF($F{r}="OK","✓ All good",IF($F{r}="Out of stock","🚨 Order NOW",IF($F{r}="Low","⚠ Reorder","? Check")))')

    for c in range(1, 12):
        cell = ws6.cell(row=r, column=c)
        cell.font      = ROW_FONT
        cell.border    = BORDER
        cell.fill      = PatternFill("solid", fgColor=row_fill)
    ws6.row_dimensions[r].height = 16

    ws6.cell(row=r, column=1).alignment  = CENTER
    ws6.cell(row=r, column=2).alignment  = LEFT
    ws6.cell(row=r, column=3).alignment  = RIGHT
    ws6.cell(row=r, column=4).alignment  = CENTER
    ws6.cell(row=r, column=5).alignment  = CENTER
    ws6.cell(row=r, column=6).alignment  = CENTER
    ws6.cell(row=r, column=7).alignment  = RIGHT
    ws6.cell(row=r, column=8).alignment  = RIGHT
    ws6.cell(row=r, column=9).alignment  = RIGHT
    ws6.cell(row=r, column=10).alignment = CENTER
    ws6.cell(row=r, column=11).alignment = CENTER

DE_dash = 13 + N - 1  # 36

# Conditional formatting on dashboard
ws6.conditional_formatting.add(f"F13:F{DE_dash}",
    CellIsRule(operator="equal",    formula=['"Low"'],          fill=PatternFill("solid", fgColor=YELLOW_LT)))
ws6.conditional_formatting.add(f"F13:F{DE_dash}",
    CellIsRule(operator="equal",    formula=['"Out of stock"'], fill=PatternFill("solid", fgColor=RED_LT)))
ws6.conditional_formatting.add(f"F13:F{DE_dash}",
    CellIsRule(operator="equal",    formula=['"OK"'],           fill=PatternFill("solid", fgColor=GREEN_LT)))
ws6.conditional_formatting.add(f"K13:K{DE_dash}",
    CellIsRule(operator="containsText", formula=['"NOW"'],
               fill=PatternFill("solid", fgColor=RED_LT))
)
ws6.conditional_formatting.add(f"D13:D{DE_dash}",
    CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor=RED_LT))
)
ws6.conditional_formatting.add(f"D13:D{DE_dash}",
    CellIsRule(operator="equal",    formula=["0"], fill=PatternFill("solid", fgColor=RED_LT))
)

ws6.freeze_panes = "A13"

# ════════════════════════════════════════════════════════════════════════════
# SHEET 7 — HOW TO USE
# ════════════════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("How to Use")
ws7.sheet_view.showGridLines = False
ws7.sheet_properties.tabColor = SLATE

ws7.column_dimensions["A"].width = 3
ws7.column_dimensions["B"].width = 110

def gtitle(row, text):
    ws7.merge_cells(f"B{row}:B{row}")
    c = ws7.cell(row=row, column=2, value=text)
    c.font = Font(name="Calibri", bold=True, size=16, color=WHITE)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws7.row_dimensions[row].height = 24

def gsection(row, text):
    c = ws7.cell(row=row, column=2, value=text)
    c.font = Font(name="Calibri", bold=True, size=13, color=NAVY)
    c.fill = PatternFill("solid", fgColor=GOLD_LT)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws7.row_dimensions[row].height = 20

def gbody(row, text):
    c = ws7.cell(row=row, column=2, value=text)
    c.font = Font(name="Calibri", size=11, color="FF2C3E50")
    c.alignment = WRAP_L
    lines = max(1, len(text)//100 + text.count("\n") + 1)
    ws7.row_dimensions[row].height = 15 * lines

def gspacer(row):
    ws7.row_dimensions[row].height = 6

rr = 1
gtitle(rr, "  STOCK INVENTORY — PROFESSIONAL EDITION — HOW TO USE"); rr += 2

gsection(rr, "📦  1. STOCK SHEET — Your Main Inventory View"); rr += 1
gbody(rr, "This is where you see everything at a glance. The top shows 6 big KPI cards: Total Items, "
          "Items In Stock, Low Stock, Out of Stock, Check Entries, and Total Stock Value in Rs. "
          "Below that, every item in your store with all its details."); rr += 1
gbody(rr, "Item Code (STK001–STK024): A unique code for each item. Use this in Stock In, Stock Out, "
          "and Adjustments sheets. The code auto-fills when you select the item from the dropdown."); rr += 1
gbody(rr, "Item Name: English name of the part. Select from the dropdown in the other sheets."); rr += 1
gbody(rr, "Urdu Name: Urdu label for quick recognition when you're on mobile or checking the shelf."); rr += 1
gbody(rr, "Unit: Default 'Pcs'. Change to kg, metre, box, etc. if you sell in different units."); rr += 1
gbody(rr, "Opening Stock: Starting count. Currently 200 for all items — update this after your first "
          "real shelf count to whatever is actually on hand."); rr += 1
gbody(rr, "Stock In (Col F): Auto-calculated. Sums all deliveries for this item from the Stock In sheet. "
          "Do NOT type here — it updates automatically when you add a delivery."); rr += 1
gbody(rr, "Stock Out (Col G): Auto-calculated. Sums all sales/issues from the Stock Out sheet. "
          "Do NOT type here."); rr += 1
gbody(rr, "Current Stock (Col H): Auto = Opening Stock + Stock In − Stock Out. Updates instantly "
          "when you add entries. This is your live stock count."); rr += 1
gbody(rr, "Minimum Stock (Col I): The level at which you want to be alerted to reorder. "
          "Set a different minimum for each item based on how fast it sells."); rr += 1
gbody(rr, "Cost Price (Rs) (Col J): What you pay per piece from your supplier. Fill this in — "
          "it's used to calculate your Total Stock Value (Rs) on the Dashboard."); rr += 1
gbody(rr, "Selling Price (Rs) (Col K): What you sell each piece for. Fill this in for profit tracking."); rr += 1
gbody(rr, "Supplier (Col L): Who supplies this item. Already filled in for all items."); rr += 1
gbody(rr, "Status (Col M): Auto-calculated — OK (green), Low (yellow), Out of stock (red), "
          "Check entries (red, if stock goes negative for some reason)."); rr += 1
gbody(rr, "Reorder Qty (Col N): Auto-suggested quantity to order when stock is Low or Out. "
          "Formula: 2×Minimum Stock − Current Stock. Blank when status is OK."); rr += 1
gbody(rr, "Notes: Free text for anything about the item — shelf location, notes from supplier, etc."); rr += 2

gsection(rr, "📥  2. STOCK IN SHEET — Record Every Delivery"); rr += 1
gbody(rr, "Use this sheet every time goods arrive at your store. Date, Item Name (pick from dropdown — "
          "Item Code auto-fills), Qty In, Supplier, Reference (invoice number), and Notes."); rr += 1
gbody(rr, "The dropdown in Column B contains all 24 item names. When you select one, Column C "
          "automatically fills in the correct Item Code (STK001, STK002, etc.)."); rr += 1
gbody(rr, "The Stock sheet's 'Stock In' column for that item will automatically add up all entries "
          "for that item — no manual calculation needed."); rr += 1
gbody(rr, "Tip: Always add a Stock In entry on the day the goods arrive. This keeps your Current Stock "
          "accurate and your Dashboard correct."); rr += 2

gsection(rr, "📤  3. STOCK OUT SHEET — Record Every Sale or Issue"); rr += 1
gbody(rr, "Use this sheet every time stock leaves your store — sale, sample given, damaged and thrown, "
          "or any other reason. Same dropdown system as Stock In."); rr += 1
gbody(rr, "Columns: Date, Item Name (dropdown → auto Item Code), Qty Out, Customer/Destination, "
          "Reference (order number), Notes."); rr += 1
gbody(rr, "The 'Stock Out' column on the Stock sheet auto-sums all entries for each item, and "
          "Current Stock updates automatically."); rr += 1
gbody(rr, "Tip: Add a Stock Out entry the same day as the sale. Your stock count stays accurate "
          "and you can track daily sales easily."); rr += 2

gsection(rr, "🔧  4. STOCK ADJUSTMENTS SHEET — Non-Sale Changes"); rr += 1
gbody(rr, "Use this for changes that are NOT normal sales. This keeps your sales data clean and "
          "helps you see what's actually selling vs. what's being lost."); rr += 1
gbody(rr, "Types (pick from dropdown): Damage (broken/lost), Return to Supplier (sent back), "
          "Return from Customer (came back), Found Missing (counted more than the sheet showed), Other."); rr += 1
gbody(rr, "Reason: Explain why. Reference: note, invoice number, or anything to identify it. "
          "Entered By: your name or whoever made the entry."); rr += 1
gbody(rr, "Important: Adjustments do NOT affect Stock In or Stock Out — they are tracked separately "
          "so you can see your losses clearly at the end of the month."); rr += 2

gsection(rr, "🏭  5. SUPPLIERS SHEET — Who Supplies What"); rr += 1
gbody(rr, "A directory of all your suppliers. The 'Items Supplied' column automatically shows which "
          "items each supplier provides (filtered from the Stock sheet)."); rr += 1
gbody(rr, "Pre-filled for each supplier: Phone number, Email, Lead Time (how many days they take "
          "to deliver after you order), and Payment Terms (when you pay them)."); rr += 1
gbody(rr, "Lead Time is important: if a supplier takes 5 days and your item is going low, order NOW "
          "so it arrives before you run out. Use this to plan reorders smartly."); rr += 1
gbody(rr, "Update phone/email/payment terms as they change. Add new suppliers by typing a new row."); rr += 2

gsection(rr, "📊  6. DASHBOARD — Your Business at a Glance"); rr += 1
gbody(rr, "The Dashboard is your daily check-in page. Open it every morning to see:"); rr += 1
gbody(rr, "• 6 big KPI cards: Total Items, Items In Stock, Low Stock, Out of Stock, OK Items, "
          "and Total Stock Value in Rs."); rr += 1
gbody(rr, "• Action Summary row: how many items are out of stock (URGENT), low (reorder soon), "
          "OK, and the total stock value in Rs."); rr += 1
gbody(rr, "• Full item table below: all 24 items with Current Stock, Status, Cost Price, Sell Price, "
          "Stock Value (Rs) = Current Stock × Cost Price, Reorder Qty, and an Action Needed column "
          "that tells you exactly what to do."); rr += 1
gbody(rr, "The 'Action Needed' column uses emojis: ✓ All good (green), ⚠ Reorder (yellow), "
          "🚨 Order NOW (red). Quick visual scan tells you what needs attention."); rr += 1
gbody(rr, "Stock Value (Rs) = Current Stock × Cost Price. Fill in Cost Price for each item to see "
          "your total inventory value. This is important for business planning and tax records."); rr += 2

gsection(rr, "📅  7. DAILY / WEEKLY / MONTHLY ROUTINE"); rr += 1
gbody(rr, "DAILY (every day):"); rr += 1
gbody(rr, "  → Add every delivery to Stock In sheet (even if it's just one item)."); rr += 1
gbody(rr, "  → Add every sale to Stock Out sheet (even small sales, samples, or issues)."); rr += 1
gbody(rr, "  → Open Dashboard and check the Action Summary — act on any 'Order NOW' items."); rr += 1
gbody(rr, "WEEKLY (once a week, e.g. every Saturday):"); rr += 1
gbody(rr, "  → Count each item on the shelf (physical count)."); rr += 1
gbody(rr, "  → Compare with Current Stock on the Stock sheet. Fix any differences using "
          "Stock Adjustments sheet (Found Missing or Damage as appropriate)."); rr += 1
gbody(rr, "  → Review all 'Low' and 'Out of stock' items. Call suppliers and place orders "
          "using the Reorder Qty suggested on the Stock sheet."); rr += 1
gbody(rr, "  → Update the Weekly Count sheet if you use it for record-keeping."); rr += 1
gbody(rr, "MONTHLY (once a month):"); rr += 1
gbody(rr, "  → Fill in Cost Price (Col J) and Selling Price (Col K) for items you haven't done yet. "
          "The Dashboard Stock Value will update automatically."); rr += 1
gbody(rr, "  → Download a backup: File → Download → Microsoft Excel (.xlsx). Keep a dated copy."); rr += 1
gbody(rr, "  → Review total stock value on Dashboard. Note it down for your business records."); rr += 1
gbody(rr, "  → Review Stock Adjustments — how much stock did you lose to damage this month? "
          "Is there a pattern? Maybe you need better handling or different suppliers."); rr += 2

gsection(rr, "👥  8. HOW TO SHARE (Google Sheets)"); rr += 1
gbody(rr, "If you use the Google Sheet version (not the Excel file):"); rr += 1
gbody(rr, "  → Open the Google Sheet → click 'Share' (top right corner)."); rr += 1
gbody(rr, "  → Add your brother's email and your partner's email in the 'Add people and groups' box."); rr += 1
gbody(rr, "  → Set their permission to 'Viewer' (can see but cannot edit)."); rr += 1
gbody(rr, "  → You remain the only 'Editor'. They can open it on phone or PC and see stock levels, "
          "but cannot change anything."); rr += 1
gbody(rr, "  → If you want them to be able to add Stock In/Out entries, set them as 'Editor' — "
          "but only if you trust them to follow the process."); rr += 2

gsection(rr, "📝  9. NOTES & TIPS"); rr += 1
gbody(rr, "• Item codes STK001–STK024 are fixed. Do not change them — they are used in all the "
          "auto-calculations and lookups across sheets."); rr += 1
gbody(rr, "• The item name dropdown is your best friend. Always use it in Stock In, Stock Out, and "
          "Adjustments. Typing the name manually will break the auto-fill and the SUMIF totals."); rr += 1
gbody(rr, "• Cost Price and Selling Price are blank right now. Fill them in as you know them. "
          "The Dashboard will immediately show your total stock value in Rs."); rr += 1
gbody(rr, "• Reorder Qty formula: 2×Minimum Stock − Current Stock. This gives you enough to get "
          "back to 2× minimum when you reorder — a comfortable buffer."); rr += 1
gbody(rr, "• If Current Stock shows 'Check entries' — something is wrong. Usually this means a "
          "Stock In or Stock Out entry has a problem. Check the transaction sheets."); rr += 1
gbody(rr, "• Excel file = your offline backup. Google Sheet = your live online copy to share with "
          "your brother and partner. Keep both up to date."); rr += 1
gbody(rr, "• If you add a new item, add it to the Stock sheet first (new row, new SKU STK025 etc.), "
          "then the dropdown in other sheets will automatically include it."); rr += 1

ws7.freeze_panes = "A1"

# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
wb.save(OUT)
print(f"✅  Saved {OUT}")
print(f"    {N} items  |  SKU STK001–STK{N:03d}  |  {TXN} txn rows")
print(f"    7 sheets: Stock | Stock In | Stock Out | Stock Adjustments |")
print(f"              Suppliers | Dashboard | How to Use")
print(f"    Result: 11 OK | 11 Low | 2 Out of stock")

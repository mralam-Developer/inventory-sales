# -*- coding: utf-8 -*-
"""Build Stock_Inventory.xlsx — professional wholesale inventory."""
import sys
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ---- 24 items: (English, Urdu, Supplier) ----
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
SKUS = [f"STK{i+1:03d}" for i in range(len(ITEMS))]
N    = len(ITEMS)

# ---- styles ----
THIN   = Side(style="thin", color="FFB0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL  = PatternFill("solid", fgColor="FF1F4E79")
HDR_FONT  = Font(name="Calibri", bold=True, color="FFFFFFFF", size=11)
KPI_LBL_FILL = PatternFill("solid", fgColor="FFD6E4F0")
KPI_LBL_FONT = Font(name="Calibri", bold=True, size=11)
BIGKPI_FILL  = PatternFill("solid", fgColor="FF1F4E79")
BIGKPI_FONT  = Font(name="Calibri", bold=True, size=20, color="FFFFFFFF")
ROW_FONT  = Font(name="Calibri", size=11)
GOOD  = PatternFill("solid", fgColor="FFC6EFCE")
LOW   = PatternFill("solid", fgColor="FFFFEB9C")
ZERO  = PatternFill("solid", fgColor="FFFFC7CE")
NEG   = PatternFill("solid", fgColor="FFFF7C7C")
LIGHT = PatternFill("solid", fgColor="FFF2F7FB")
CENTER= Alignment(horizontal="center", vertical="center")
LEFT  = Alignment(horizontal="left",   vertical="center")
RIGHT = Alignment(horizontal="right",  vertical="center")
WRAP  = Alignment(horizontal="left",   vertical="top", wrap_text=True)


def hdr(ws, row, ncols, start=1):
    for c in range(start, start + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.alignment = CENTER; cell.border = BORDER


def base_row(ws, row, ncols, start=1):
    for c in range(start, start + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = ROW_FONT; cell.border = BORDER
        cell.alignment = LEFT


# ===================== Workbook =====================
wb = Workbook()

# ===================== Sheet 1: Stock =====================
ws = wb.active
ws.title = "Stock"
ws.sheet_view.showGridLines = True

headers = [
    "Item Code", "Item Name", "Urdu Name", "Unit",
    "Opening Stock", "Stock In", "Stock Out", "Current Stock",
    "Minimum Stock", "Cost Price (Rs)", "Selling Price (Rs)",
    "Supplier", "Status", "Reorder Qty", "Notes"
]
NC  = len(headers)
DS  = 5          # data starts at row 5
DE  = DS + N - 1 # row 28

# ---- Row 1: KPI strip ----
kpi_labels = ["Items", "In Stock", "Low", "Out of Stock", "Check", "Total Value (Rs)"]
kpi_cols   = [1, 2, 3, 4, 5, 6]  # A-F
for i, (lbl, col) in enumerate(zip(kpi_labels, kpi_cols), 1):
    c = ws.cell(row=1, column=col, value=lbl)
    c.fill = KPI_LBL_FILL; c.font = KPI_LBL_FONT
    c.alignment = CENTER; c.border = BORDER
ws.row_dimensions[1].height = 20

# ---- Row 2: KPI values ----
ws["A2"] = f"=COUNTA(A{DS}:A{DE})"
ws["B2"] = f'=COUNTIF(H{DS}:H{DE},">0")'
ws["C2"] = f'=COUNTIF(M{DS}:M{DE},"Low")'
ws["D2"] = f'=COUNTIF(M{DS}:M{DE},"Out of stock")'
ws["E2"] = f'=COUNTIF(M{DS}:M{DE},"Check entries")'
ws["F2"] = f"=SUMPRODUCT((H{DS}:H{DE}>0)*(H{DS}:H{DE})*(J{DS}:J{DE}))"
for c in range(1, 7):
    cell = ws.cell(row=2, column=c)
    cell.font = Font(name="Calibri", bold=True, size=14,
                     color="FF1F4E79")
    cell.alignment = CENTER
    cell.border = BORDER
ws.row_dimensions[2].height = 22

# ---- Row 3: spacer ----
ws.row_dimensions[3].height = 6

# ---- Row 4: header ----
for i, h in enumerate(headers, 1):
    ws.cell(row=4, column=i, value=h)
hdr(ws, 4, NC)
ws.row_dimensions[4].height = 24

# ---- Column widths ----
widths = [10, 24, 22, 7, 13, 11, 11, 13,
          13, 13, 13, 18, 14, 12, 20]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- Data rows ----
MIN_STOCK = [30, 30, 40, 20, 20, 30, 30, 25, 15, 20,
              20, 15, 10, 25, 20, 20, 20, 30, 20, 20,
              20, 15, 10, 15]

for idx, (eng, urd, sup) in enumerate(ITEMS):
    r = DS + idx

    ws.cell(row=r, column=1,  value=SKUS[idx])          # A: Item Code
    ws.cell(row=r, column=2,  value=eng)                 # B: Item Name
    ws.cell(row=r, column=3,  value=urd)                 # C: Urdu Name
    ws.cell(row=r, column=4,  value="Pcs")               # D: Unit
    ws.cell(row=r, column=5,  value=200)                 # E: Opening Stock
    ws.cell(row=r, column=6,  value=                     # F: Stock In
            f"=SUMIF('Stock In'!$B$2:$B${TXN+1},$B{r},"
            f"'Stock In'!$D$2:$D${TXN+1})")
    ws.cell(row=r, column=7,  value=                     # G: Stock Out
            f"=SUMIF('Stock Out'!$B$2:$B${TXN+1},$B{r},"
            f"'Stock Out'!$D$2:$D${TXN+1})")
    ws.cell(row=r, column=8,  value=                     # H: Current Stock
            f"=IFERROR($E{r}+$F{r}-$G{r},\"Check entries\")")
    ws.cell(row=r, column=9,  value=MIN_STOCK[idx])      # I: Minimum Stock
    ws.cell(row=r, column=12, value=sup)                 # L: Supplier
    ws.cell(row=r, column=13, value=(                    # M: Status
        f'=IF($H{r}<0,"Check entries",'
        f'IF($H{r}=0,"Out of stock",'
        f'IF(AND($I{r}<>"",$H{r}<=$I{r}),"Low","OK")))'))
    ws.cell(row=r, column=14, value=                     # N: Reorder Qty
        f'=IF(OR($M{r}="OK",$M{r}=""),"",MAX(0,$I{r}*2-$H{r}))')
    ws.cell(row=r, column=15, value="")                  # O: Notes

    base_row(ws, r, NC)
    ws.row_dimensions[r].height = 17
    ws.cell(row=r, column=3).alignment  = RIGHT  # Urdu
    ws.cell(row=r, column=5).alignment  = RIGHT  # Opening
    ws.cell(row=r, column=6).alignment  = RIGHT  # In
    ws.cell(row=r, column=7).alignment  = RIGHT  # Out
    ws.cell(row=r, column=8).alignment  = RIGHT  # Current
    ws.cell(row=r, column=9).alignment  = RIGHT  # Min Stock
    ws.cell(row=r, column=13).alignment = CENTER # Status
    ws.cell(row=r, column=14).alignment = CENTER # Reorder Qty

# ---- Conditional formatting ----
ws.conditional_formatting.add(f"H{DS}:H{DE}",
    CellIsRule(operator="lessThan", formula=["0"], fill=NEG))
ws.conditional_formatting.add(f"H{DS}:H{DE}",
    CellIsRule(operator="equal",    formula=["0"],  fill=ZERO))
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal",    formula=['"Low"'],          fill=LOW))
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal",    formula=['"Out of stock"'], fill=ZERO))
ws.conditional_formatting.add(f"M{DS}:M{DE}",
    CellIsRule(operator="equal",    formula=['"Check entries"'], fill=NEG))
ws.conditional_formatting.add(f"A{DS}:O{DE}",
    CellIsRule(operator="equal",    formula=['"Low"'],
               fill=PatternFill("solid", fgColor="FFFFFDE7")))
ws.conditional_formatting.add(f"A{DS}:O{DE}",
    CellIsRule(operator="equal",    formula=['"Out of stock"'],
               fill=PatternFill("solid", fgColor="FFFFC7CE")))

ws.freeze_panes   = "A5"
ws.auto_filter.ref = f"A4:O{DE}"

# ===================== Sheet 2: Stock In =====================
ws2 = wb.create_sheet("Stock In")
in_h = ["Date", "Item Name", "Item Code", "Qty In", "Unit",
        "Supplier", "Reference", "Notes"]
for i, h in enumerate(in_h, 1):
    ws2.cell(row=1, column=i, value=h)
hdr(ws2, 1, len(in_h))
ws2.row_dimensions[1].height = 22
for i, w in enumerate([13, 24, 10, 9, 7, 18, 20, 20], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

dv2 = DataValidation(type="list",
    formula1=f"='Stock'!$B${DS}:$B${DE}", allow_blank=True)
ws2.add_data_validation(dv2)
dv2.add(f"B2:B{TXN+1}")

for r in range(2, TXN + 2):
    ws2.cell(row=r, column=3,
             value=(f'=IF($B{r}="","",'
                    f'INDEX(\'Stock\'!$A${DS}:$A${DE},'
                    f'MATCH($B{r},\'Stock\'!$B${DS}:$B${DE},0)))'))
    for c in range(1, 9):
        ws2.cell(row=r, column=c).font      = ROW_FONT
        ws2.cell(row=r, column=c).border   = BORDER
        ws2.cell(row=r, column=c).alignment = LEFT
    ws2.cell(row=r, column=5, value="Pcs")
    ws2.cell(row=r, column=5).alignment = CENTER
    ws2.cell(row=r, column=3).alignment = CENTER
    ws2.row_dimensions[r].height = 16

_sample_in = [
    ("2026-11-03","Upper Body",  50, "ABC Traders",   "Weekly delivery",""),
    ("2026-11-04","Cat",        100, "XYZ Wholesale", "Restock",        ""),
    ("2026-11-05","Bolt Jora",   80, "ABC Traders",   "Weekly delivery",""),
    ("2026-11-03","Gas Pipe",    60, "Gas Co. Ltd",   "New stock",      ""),
    ("2026-11-06","Silver Ring",120, "Jewellery Sup.","Restock",        ""),
]
for i, (dt, item, qty, sup, ref, note) in enumerate(_sample_in):
    r = 2 + i
    ws2.cell(row=r, column=1, value=dt)
    ws2.cell(row=r, column=2, value=item)
    ws2.cell(row=r, column=4, value=qty)
    ws2.cell(row=r, column=6, value=sup)
    ws2.cell(row=r, column=7, value=ref)
    ws2.cell(row=r, column=8, value=note)
    for c in range(1, 9):
        ws2.cell(row=r, column=c).font      = ROW_FONT
        ws2.cell(row=r, column=c).border   = BORDER
        ws2.cell(row=r, column=c).alignment = LEFT
    ws2.cell(row=r, column=1).alignment = CENTER
    ws2.cell(row=r, column=3).alignment = CENTER
    ws2.cell(row=r, column=4).alignment = RIGHT
    ws2.row_dimensions[r].height = 16

ws2.freeze_panes   = "A2"
ws2.auto_filter.ref = f"A1:H{TXN+1}"

# ===================== Sheet 3: Stock Out =====================
ws3 = wb.create_sheet("Stock Out")
out_h = ["Date", "Item Name", "Item Code", "Qty Out", "Unit",
         "Customer", "Reference", "Notes"]
for i, h in enumerate(out_h, 1):
    ws3.cell(row=1, column=i, value=h)
hdr(ws3, 1, len(out_h))
ws3.row_dimensions[1].height = 22
for i, w in enumerate([13, 24, 10, 9, 7, 20, 20, 20], 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

dv3 = DataValidation(type="list",
    formula1=f"='Stock'!$B${DS}:$B${DE}", allow_blank=True)
ws3.add_data_validation(dv3)
dv3.add(f"B2:B{TXN+1}")

for r in range(2, TXN + 2):
    ws3.cell(row=r, column=3,
             value=(f'=IF($B{r}="","",'
                    f'INDEX(\'Stock\'!$A${DS}:$A${DE},'
                    f'MATCH($B{r},\'Stock\'!$B${DS}:$B${DE},0)))'))
    for c in range(1, 9):
        ws3.cell(row=r, column=c).font      = ROW_FONT
        ws3.cell(row=r, column=c).border   = BORDER
        ws3.cell(row=r, column=c).alignment = LEFT
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
    r = 2 + i
    ws3.cell(row=r, column=1, value=dt)
    ws3.cell(row=r, column=2, value=item)
    ws3.cell(row=r, column=4, value=qty)
    ws3.cell(row=r, column=6, value=cust)
    ws3.cell(row=r, column=7, value=ref)
    ws3.cell(row=r, column=8, value=note)
    for c in range(1, 9):
        ws3.cell(row=r, column=c).font      = ROW_FONT
        ws3.cell(row=r, column=c).border   = BORDER
        ws3.cell(row=r, column=c).alignment = LEFT
    ws3.cell(row=r, column=1).alignment = CENTER
    ws3.cell(row=r, column=3).alignment = CENTER
    ws3.cell(row=r, column=4).alignment = RIGHT
    ws3.row_dimensions[r].height = 16

ws3.freeze_panes   = "A2"
ws3.auto_filter.ref = f"A1:H{TXN+1}"

# ===================== Sheet 4: Stock Adjustments =====================
ws4 = wb.create_sheet("Stock Adjustments")
adj_h = ["Date", "Item Name", "Item Code", "Qty", "Type",
         "Reason", "Reference", "Entered By"]
adj_types = ["Damage", "Return to Supplier", "Return from Customer",
             "Found Missing", "Other"]
for i, h in enumerate(adj_h, 1):
    ws4.cell(row=1, column=i, value=h)
hdr(ws4, 1, len(adj_h))
ws4.row_dimensions[1].height = 22
for i, w in enumerate([13, 24, 10, 9, 18, 24, 18, 18], 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

dv4a = DataValidation(type="list",
    formula1=f"='Stock'!$B${DS}:$B${DE}", allow_blank=True)
ws4.add_data_validation(dv4a)
dv4a.add(f"B2:B{TXN+1}")

dv4b = DataValidation(type="list",
    formula1=f'="{",".join(adj_types)}"', allow_blank=True)
ws4.add_data_validation(dv4b)
dv4b.add(f"E2:E{TXN+1}")

for r in range(2, TXN + 2):
    ws4.cell(row=r, column=3,
             value=(f'=IF($B{r}="","",'
                    f'INDEX(\'Stock\'!$A${DS}:$A${DE},'
                    f'MATCH($B{r},\'Stock\'!$B${DS}:$B${DE},0)))'))
    for c in range(1, 9):
        ws4.cell(row=r, column=c).font      = ROW_FONT
        ws4.cell(row=r, column=c).border   = BORDER
        ws4.cell(row=r, column=c).alignment = LEFT
    ws4.cell(row=r, column=4, value="Pcs")
    ws4.cell(row=r, column=4).alignment = CENTER
    ws4.cell(row=r, column=3).alignment = CENTER
    ws4.row_dimensions[r].height = 16

ws4.freeze_panes   = "A2"
ws4.auto_filter.ref = f"A1:H{TXN+1}"

# ===================== Sheet 5: Suppliers =====================
ws5 = wb.create_sheet("Suppliers")
sup_h = ["Supplier Name", "Phone / Contact", "Email",
         "Items Supplied", "Lead Time (days)",
         "Payment Terms", "Notes"]
for i, h in enumerate(sup_h, 1):
    ws5.cell(row=1, column=i, value=h)
hdr(ws5, 1, len(sup_h))
ws5.row_dimensions[1].height = 22
for i, w in enumerate([22, 18, 22, 28, 14, 16, 24], 1):
    ws5.column_dimensions[get_column_letter(i)].width = w

seen    = set()
sup_list = []
for _, _, s in ITEMS:
    if s not in seen:
        seen.add(s)
        sup_list.append(s)

# Supplier details database
SUP_DB = {
    "ABC Traders":   ("0300-1234567",   "abctraders@example.com",  2, "Cash on delivery"),
    "XYZ Wholesale": ("0333-9876543",   "xyzwholesale@example.com", 3, "7 days credit"),
    "Local Market":  ("0345-2223334",   "",                         1, "Cash"),
    "Gas Co. Ltd":   ("0300-9998887",    "gasco@example.com",        5, "Advance 50%"),
    "Jewellery Sup.":("0348-5556667",   "jewellery@example.com",    4, "Cash on delivery"),
    "Pipe Co.":      ("0300-6665544",   "pipeco@example.com",       3, "7 days credit"),
}

for i, sup in enumerate(sup_list):
    r   = 2 + i
    info = SUP_DB.get(sup, ("", "", "", ""))

    ws5.cell(row=r, column=1, value=sup)
    ws5.cell(row=r, column=2, value=info[0])  # Phone
    ws5.cell(row=r, column=3, value=info[1])  # Email
    ws5.cell(row=r, column=4,
             value=(f'=TEXTJOIN(", ",TRUE,'
                    f'FILTER(\'Stock\'!$B${DS}:$B${DE},'
                    f'\'Stock\'!$L${DS}:$L${DE}=$A{r},""))'))
    ws5.cell(row=r, column=5, value=info[2])  # Lead time
    ws5.cell(row=r, column=6, value=info[3])  # Payment terms

    for c in range(1, 8):
        ws5.cell(row=r, column=c).font      = ROW_FONT
        ws5.cell(row=r, column=c).border   = BORDER
        ws5.cell(row=r, column=c).alignment = LEFT
    ws5.cell(row=r, column=4).alignment = LEFT
    ws5.row_dimensions[r].height = 16

ws5.freeze_panes = "A2"

# ===================== Sheet 6: Dashboard =====================
ws6 = wb.create_sheet("Dashboard")
ws6.sheet_view.showGridLines = False

for col, w in [("A",3),("B",26),("C",16),("D",16),
               ("E",16),("F",16),("G",3)]:
    ws6.column_dimensions[col].width = w

# Title
ws6.merge_cells("B2:F2")
t = ws6.cell(row=2, column=2, value="Stock Dashboard")
t.font = Font(name="Calibri", bold=True, size=18, color="FF1F4E79")
ws6.row_dimensions[2].height = 28

ws6.merge_cells("B3:F3")
s = ws6.cell(row=3, column=2,
             value="Wholesale Stock Inventory — Status at a Glance")
s.font = Font(name="Calibri", size=11, color="FF555555")
ws6.row_dimensions[3].height = 16

ws6.row_dimensions[4].height = 8  # spacer

# Big KPI cards — row 5-6 (card + label) and row 7-8 (card + label)
# Card layout:
#   Row 5: Total Items   [B5]  |  Items In Stock  [D5]
#   Row 6: labels for row 5
#   Row 7: Low Stock     [B7]  |  Out of Stock    [D7]
#   Row 8: labels for row 7
cards = [
    ("Total Items",
     f"=COUNTA('Stock'!$A${DS}:$A${DE})",
     "B5", "C5", "B6", "C6"),
    ("Items In Stock",
     f'=COUNTIF(\'Stock\'!$H${DS}:$H${DE},">0")',
     "D5", "E5", "D6", "E6"),
    ("Low Stock",
     f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Low")',
     "B7", "C7", "B8", "C8"),
    ("Out of Stock",
     f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Out of stock")',
     "D7", "E7", "D8", "E8"),
]
for label, formula, card_merge_lr, card_merge_rr, lbl_merge_lr, lbl_merge_rr in cards:
    # Card (big number)
    ws6.merge_cells(f"{card_merge_lr}:{card_merge_rr}")
    tl = card_merge_lr
    cell = ws6[tl]
    cell.value     = formula
    cell.fill      = BIGKPI_FILL
    cell.font      = BIGKPI_FONT
    cell.alignment = CENTER
    cell.border    = BORDER

    # Label
    ws6.merge_cells(f"{lbl_merge_lr}:{lbl_merge_rr}")
    ll = lbl_merge_lr
    lcell = ws6[ll]
    lcell.value     = label
    lcell.fill      = LIGHT
    lcell.font      = Font(name="Calibri", bold=True, size=11)
    lcell.alignment = CENTER
    lcell.border    = BORDER

ws6.row_dimensions[5].height = 36  # first card row
ws6.row_dimensions[6].height = 18  # first label row
ws6.row_dimensions[7].height = 36  # second card row
ws6.row_dimensions[8].height = 18  # second label row

ws6.row_dimensions[9].height = 8   # spacer

# Item status table header (row 10)
status_headers = [
    "Item Code", "Item Name", "Urdu Name", "Current Stock",
    "Min Stock",  "Status", "Cost Price", "Sell Price",
    "Stock Value (Rs)", "Reorder?"
]
for i, h in enumerate(status_headers, 1):
    ws6.cell(row=10, column=i, value=h)
hdr(ws6, 10, len(status_headers))
ws6.row_dimensions[10].height = 22

for i, w in enumerate([10, 24, 22, 13, 10, 14, 12, 12, 14, 12], 1):
    ws6.column_dimensions[get_column_letter(i)].width = w

for idx in range(N):
    r    = 11 + idx
    srow = DS + idx

    ws6.cell(row=r, column=1,  value=f"='Stock'!$A${srow}")
    ws6.cell(row=r, column=2,  value=f"='Stock'!$B${srow}")
    ws6.cell(row=r, column=3,  value=f"='Stock'!$C${srow}")
    ws6.cell(row=r, column=4,  value=f"='Stock'!$H${srow}")
    ws6.cell(row=r, column=5,  value=f"='Stock'!$I${srow}")
    ws6.cell(row=r, column=6,  value=f"='Stock'!$M${srow}")
    ws6.cell(row=r, column=7,  value=f"='Stock'!$J${srow}")
    ws6.cell(row=r, column=8,  value=f"='Stock'!$K${srow}")
    ws6.cell(row=r, column=9,  value=f"=IFERROR($D{r}*$G{r},0)")
    ws6.cell(row=r, column=10,
             value=f'=IF($F{r}="OK","","Reorder now")')

    base_row(ws6, r, 10)
    ws6.cell(row=r, column=3).alignment  = RIGHT
    ws6.cell(row=r, column=4).alignment  = CENTER
    ws6.cell(row=r, column=5).alignment  = CENTER
    ws6.cell(row=r, column=6).alignment  = CENTER
    ws6.cell(row=r, column=7).alignment  = RIGHT
    ws6.cell(row=r, column=8).alignment  = RIGHT
    ws6.cell(row=r, column=9).alignment  = RIGHT
    ws6.cell(row=r, column=10).alignment = CENTER
    ws6.row_dimensions[r].height = 16

DE_dash = 11 + N - 1  # row 34

ws6.conditional_formatting.add(f"F11:F{DE_dash}",
    CellIsRule(operator="equal",    formula=['"Low"'],           fill=LOW))
ws6.conditional_formatting.add(f"F11:F{DE_dash}",
    CellIsRule(operator="equal",    formula=['"Out of stock"'], fill=ZERO))
ws6.conditional_formatting.add(f"J11:J{DE_dash}",
    CellIsRule(operator="equal",    formula=['"Reorder now"'],   fill=NEG))
ws6.conditional_formatting.add(f"D11:D{DE_dash}",
    CellIsRule(operator="lessThan", formula=["0"],               fill=NEG))
ws6.conditional_formatting.add(f"D11:D{DE_dash}",
    CellIsRule(operator="equal",    formula=["0"],               fill=ZERO))

ws6.freeze_panes = "A11"

# ===================== Sheet 7: How to Use =====================
ws7 = wb.create_sheet("How to Use")
ws7.column_dimensions["A"].width = 3
ws7.column_dimensions["B"].width = 100
ws7.sheet_view.showGridLines = False

title_f = Font(name="Calibri", bold=True, size=14, color="FF1F4E79")
sec_f   = Font(name="Calibri", bold=True, size=12, color="FF1F4E79")
body_f  = Font(name="Calibri", size=11)

def h1(row, text):
    c = ws7.cell(row=row, column=2, value=text)
    c.font = title_f
    ws7.row_dimensions[row].height = 22

def h2(row, text):
    c = ws7.cell(row=row, column=2, value=text)
    c.font = sec_f
    ws7.row_dimensions[row].height = 18

def body(row, text):
    c = ws7.cell(row=row, column=2, value=text)
    c.font = body_f
    c.alignment = WRAP
    lines = max(1, len(text)//95 + text.count("\n") + 1)
    ws7.row_dimensions[row].height = 15 * lines

rr = 1
h1(rr, "Stock Inventory — Professional Edition — How to Use"); rr += 2

h2(rr, "1. STOCK SHEET — main inventory view"); rr += 1
body(rr, "Item Code (STK001–STK024): a unique code for each item. "
         "Use this code in Stock In, Stock Out, and Adjustments sheets. "
         "The code auto-fills when you pick the item name from the dropdown."); rr += 1
body(rr, "Item Name: English name of the part. Select from the dropdown in "
         "Stock In / Out / Adjustments sheets — the Item Code fills automatically."); rr += 1
body(rr, "Urdu Name: Urdu label for quick recognition on mobile."); rr += 1
body(rr, "Unit: default 'Pcs'. Change if you sell in kg, metres, boxes, etc."); rr += 1
body(rr, "Opening Stock: starting count. Currently set to 200 for all items — "
         "update after your first real shelf count."); rr += 1
body(rr, "Stock In (col F): auto-sum of all deliveries for this item from the "
         "Stock In sheet. Do NOT type here — updates automatically."); rr += 1
body(rr, "Stock Out (col G): auto-sum of all sales/issues for this item from "
         "the Stock Out sheet. Do NOT type here."); rr += 1
body(rr, "Current Stock (col H): auto = Opening Stock + Stock In − Stock Out. "
         "Updates instantly when you add entries to Stock In or Stock Out."); rr += 1
body(rr, "Minimum Stock (col I): the level at which you want to be alerted "
         "to reorder. Edit per item as needed."); rr += 1
body(rr, "Cost Price (Rs) (col J): what you pay per piece. Fill this in for "
         "your items. Used for total stock value on the Dashboard."); rr += 1
body(rr, "Selling Price (Rs) (col K): what you charge per piece. Fill this in. "
         "Used for stock value and profit tracking on the Dashboard."); rr += 1
body(rr, "Supplier (col L): who supplies this item. Update as needed."); rr += 1
body(rr, "Status (col M): auto — OK / Low / Out of stock / Check entries."); rr += 1
body(rr, "Reorder Qty (col N): auto-suggested quantity to order when stock is "
         "Low or Out. Formula: 2×Min − Current Stock. Blank when OK."); rr += 1
body(rr, "Notes (col O): free text for anything about the item."); rr += 2

h2(rr, "2. STOCK IN SHEET — record every delivery"); rr += 1
body(rr, "Date: date of delivery. Item Name: pick from dropdown (auto-fills "
         "Item Code in col C). Qty In: how many pieces received. Supplier: who "
         "delivered. Reference: invoice number or note. Notes: anything else."); rr += 1
body(rr, "Tip: always add a Stock In entry on the day the goods arrive — this "
         "keeps Current Stock accurate on the Stock sheet."); rr += 2

h2(rr, "3. STOCK OUT SHEET — record every sale or issue"); rr += 1
body(rr, "Date: date of sale. Item Name: pick from dropdown (auto-fills Item "
         "Code in col C). Qty Out: how many pieces went out. Customer: who bought "
         "it. Reference: order number. Notes: anything else."); rr += 1
body(rr, "Tip: add a Stock Out entry the same day as the sale — the Current "
         "Stock on the Stock sheet updates automatically."); rr += 2

h2(rr, "4. STOCK ADJUSTMENTS SHEET — for non-sale changes"); rr += 1
body(rr, "Use this for: Damage (items broken/lost), Return to Supplier (sent "
         "back), Return from Customer (came back), Found Missing (counted more "
         "than sheet showed). Type: pick from the dropdown. Reason: explain why. "
         "Reference: note or invoice. Entered By: your name."); rr += 1
body(rr, "These adjustments do NOT go into Stock In or Stock Out — they are "
         "separate so you can see what is normal sales vs. losses."); rr += 2

h2(rr, "5. SUPPLIERS SHEET — who supplies what"); rr += 1
body(rr, "Lists each supplier and which items they supply (auto-filtered from "
         "the Stock sheet). Phone, email, lead time (days), and payment terms are "
         "pre-filled for known suppliers — update as needed. Lead time helps you "
         "plan reorders before items go out of stock."); rr += 2

h2(rr, "6. DASHBOARD — status at a glance"); rr += 1
body(rr, "Shows big numbers: Total Items, Items In Stock, Low Stock, Out of "
         "Stock. Below them, a table of all 24 items with Current Stock, Status, "
         "Cost Price, Sell Price, Stock Value (Rs), and a 'Reorder now' flag. "
         "Use this as your daily check — items flagged 'Reorder now' need attention."); rr += 1
body(rr, "Stock Value (Rs) = Current Stock × Cost Price. Fill in Cost Price for "
         "each item to see your total inventory value."); rr += 2

h2(rr, "7. DAILY / WEEKLY / MONTHLY ROUTINE"); rr += 1
body(rr, "Daily: add every delivery to Stock In, every sale to Stock Out. "
         "Check Dashboard for 'Reorder now' items."); rr += 1
body(rr, "Weekly: count shelves. Fix differences. Review Low / Out of stock "
         "items. Call suppliers for reorders using the Reorder Qty suggested on "
         "the Stock sheet."); rr += 1
body(rr, "Monthly: fill in Cost Price and Selling Price if not done. Download "
         "a backup (File → Download → Microsoft Excel). Review total stock value "
         "on the Dashboard."); rr += 2

h2(rr, "8. HOW TO SHARE (Google Sheets)"); rr += 1
body(rr, "Open the Google Sheet → click Share (top right) → add your brother's "
         "email and partner's email → set as Viewer. You are the only Editor. "
         "They can see stock on phone or PC but cannot change it."); rr += 2

h2(rr, "9. NOTES"); rr += 1
body(rr, "• Item codes STK001–STK024 are auto-assigned. The code auto-fills "
         "in Stock In, Stock Out, and Adjustments when you pick the item from "
         "the dropdown."); rr += 1
body(rr, "• Cost Price and Selling Price are blank — fill them in when you know "
         "them. The Dashboard Stock Value and total value will update automatically."); rr += 1
body(rr, "• Reorder Qty on the Stock sheet shows how many to order when stock is "
         "Low or Out. Formula = 2×Minimum Stock − Current Stock."); rr += 1
body(rr, "• Always use the item name dropdown — this keeps the auto-calculations "
         "and Item Code lookups correct."); rr += 1
body(rr, "• Excel file = offline backup. Google Sheet = live online copy to share."); rr += 1

ws7.freeze_panes = "A1"

wb.save(OUT)
print(f"Saved {OUT}: {N} items, {TXN} txn rows, SKU codes STK001-STK{N:03d}")

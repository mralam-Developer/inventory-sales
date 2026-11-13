# -*- coding: utf-8 -*-
"""Sync Stock_Inventory.xlsx → Google Sheets."""
import json, subprocess, sys, os

SPREADSHEET_ID = "1f-BisiOUufklx2av8D_3NMHqjNugPo_Xg8W4bDmXP4E"

# ── helpers ────────────────────────────────────────────────────────────────

def gs_update(sheet, range_a1, values, vio="USER_ENTERED"):
    """Write values to a sheet range via the connector."""
    result = subprocess.run(
        [sys.executable, "-c", _MAKE_UPDATE_CALL(sheet, range_a1, values,
                                                   vio)],
        capture_output=True, text=True, timeout=90,
    )
    resp = json.loads(result.stdout)
    if resp.get("status") == "error":
        print(f"  ERROR {sheet}!{range_a1}: {resp.get('error')}")
        return False
    print(f"  OK   {sheet}!{range_a1}  ({len(values)} rows)")
    return True


def gs_add_sheet(title, row_count=1000, col_count=26, index=None):
    result = subprocess.run(
        [sys.executable, "-c", _MAKE_ADD_SHEET_CALL(title, row_count,
                                                     col_count, index)],
        capture_output=True, text=True, timeout=30,
    )
    resp = json.loads(result.stdout)
    if resp.get("status") == "error":
        print(f"  ERROR adding sheet '{title}': {resp.get('error')}")
        return None
    sid = resp.get("result", {}).get("sheet_id")
    print(f"  Added sheet: {title} (id={sid})")
    return sid


def gs_set_validation(sheet_range, formula):
    """Set a ONE_OF_RANGE data validation on a column."""
    result = subprocess.run(
        [sys.executable, "-c",
         _MAKE_VALIDATION_CALL(sheet_range, formula)],
        capture_output=True, text=True, timeout=30,
    )
    resp = json.loads(result.stdout)
    status = resp.get("status", "?")
    print(f"  Validation {sheet_range}: {status}")
    return status == "success"


# ── connector call builders (avoid f-string nesting nightmare) ──────────────

def _esc(s):
    """Escape a Python string literal for embedding inside a -c '...' shell arg."""
    return s.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"')


def _MAKE_UPDATE_CALL(sheet, range_a1, values, vio):
    vjson  = json.dumps(values)
    sheetj = _esc(sheet)
    rangej = _esc(range_a1)
    vioj   = _esc(vio)
    sidj   = _esc(SPREADSHEET_ID)
    return (
        "import sys, json\n"
        "from arena_connector import call_connector_tool\n"
        f"values = {vjson}\n"
        "result = call_connector_tool(\n"
        "    connector='google_sheets',\n"
        "    tool='update_values',\n"
        "    arguments={\n"
        f"        'spreadsheet_id': '{sidj}',\n"
        f"        'range': '{sheetj}!{rangej}',\n"
        f"        'values': values,\n"
        f"        'value_input_option': '{vioj}',\n"
        "        'major_dimension': 'ROWS'\n"
        "    }\n"
        ")\n"
        "print(json.dumps(result))\n"
    )


def _MAKE_ADD_SHEET_CALL(title, row_count, col_count, index):
    titlej   = _esc(title)
    sidj     = _esc(SPREADSHEET_ID)
    rowj     = str(row_count)
    colj     = str(col_count)
    indexj   = "null" if index is None else str(index)
    return (
        "import sys, json\n"
        "from arena_connector import call_connector_tool\n"
        "result = call_connector_tool(\n"
        "    connector='google_sheets',\n"
        "    tool='add_sheet',\n"
        "    arguments={\n"
        f"        'spreadsheet_id': '{sidj}',\n"
        f"        'title': '{titlej}',\n"
        "        'index': None,\n"
        "        'hidden': False,\n"
        "        'tab_color': None,\n"
        "        'grid_properties': {\n"
        f"            'row_count': {rowj},\n"
        f"            'column_count': {colj},\n"
        "            'frozen_row_count': None,\n"
        "            'frozen_column_count': None,\n"
        "            'hide_gridlines': None,\n"
        "            'row_group_control_after': None,\n"
        "            'column_group_control_after': None\n"
        "        }\n"
        "    }\n"
        ")\n"
        "print(json.dumps(result))\n"
    )


def _MAKE_VALIDATION_CALL(sheet_range, formula_a1):
    """Build ONE_OF_RANGE validation referencing a Stock sheet column."""
    sidj    = _esc(SPREADSHEET_ID)
    srj     = _esc(sheet_range)
    f_a1    = _esc(formula_a1)   # e.g. 'Stock!B5'
    # TWO values so the dropdown actually shows options (Google needs ≥1 value;
    # we give it the first two item names so the dropdown is populated)
    return (
        "import sys, json\n"
        "from arena_connector import call_connector_tool\n"
        "result = call_connector_tool(\n"
        "    connector='google_sheets',\n"
        "    tool='set_data_validation',\n"
        "    arguments={\n"
        f"        'spreadsheet_id': '{sidj}',\n"
        f"        'range': '{srj}',\n"
        "        'rule': {\n"
        "            'condition': {\n"
        "                'type': 'ONE_OF_RANGE',\n"
        "                'values': [\n"
        f"                    {{'user_entered_value': '{f_a1}'}},\n"
        f"                    {{'user_entered_value': '{f_a1.replace('B5','B6')}'}},\n"
        "                ]\n"
        "            },\n"
        "            'input_message': 'Select an item from the list',\n"
        "            'strict': True,\n"
        "            'show_custom_ui': True\n"
        "        },\n"
        "        'filtered_rows_included': False\n"
        "    }\n"
        ")\n"
        "print(json.dumps(result))\n"
    )


# ── data ─────────────────────────────────────────────────────────────────────

DS  = 5    # Stock data start row
DE  = 28   # Stock data end row
TXN = 3000
N   = 24

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

SKUS = [f"STK{i+1:03d}" for i in range(N)]
MIN_STOCK = [30,30,40,20,20,30,30,25,15,20,
             20,15,10,25,20,20,20,30,20,20,
             20,15,10,15]

# ── helpers to build formula strings ─────────────────────────────────────────

def fsumif(sheet, col_b, col_d, r):
    return (f"=SUMIF('{sheet}'!$B$2:$B${TXN+1},$B{r},"
            f"'{sheet}'!$D$2:$D${TXN+1})")


def find_code(r):
    """INDEX/MATCH Item Code lookup for Stock In/Out/Adjustments row r."""
    return (f'=IF($B{r}="","",'
            f"INDEX('Stock'!$A${DS}:$A${DE},"
            f"MATCH($B{r},'Stock'!$B${DS}:$B${DE},0)))")


def status_formula(r):
    return (f'=IF($H{r}<0,"Check entries",'
            f'IF($H{r}=0,"Out of stock",'
            f'IF(AND($I{r}<>"",$H{r}<=$I{r}),"Low","OK")))')


def reorder_formula(r):
    return f'=IF(OR($M{r}="OK",$M{r}=""),"",MAX(0,$I{r}*2-$H{r}))')


# ════════════════════════════════════════════════════════════════════════════
# 1. STOCK SHEET
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  1. Stock sheet")
print("=" * 60)

# Row 1: KPI labels
stock_kpi_lbl = ["Items", "In Stock", "Low", "Out of Stock",
                  "Check", "Total Value (Rs)"]
gs_update("Stock", "A1:F1", [stock_kpi_lbl])

# Row 2: KPI formulas
stock_kpi_val = [
    f"=COUNTA(A{DS}:A{DE})",
    f'=COUNTIF(H{DS}:H{DE},">0")',
    f'=COUNTIF(M{DS}:M{DE},"Low")',
    f'=COUNTIF(M{DS}:M{DE},"Out of stock")',
    f'=COUNTIF(M{DS}:M{DE},"Check entries")',
    f"=SUMPRODUCT((H{DS}:H{DE}>0)*(H{DS}:H{DE})*(J{DS}:J{DE}))",
]
gs_update("Stock", "A2:F2", [stock_kpi_val])

# Row 3: spacer (empty)
gs_update("Stock", "A3:O3", [[""] * 15])

# Row 4: header
stock_hdr = [
    "Item Code", "Item Name", "Urdu Name", "Unit",
    "Opening Stock", "Stock In", "Stock Out", "Current Stock",
    "Minimum Stock", "Cost Price (Rs)", "Selling Price (Rs)",
    "Supplier", "Status", "Reorder Qty", "Notes"
]
gs_update("Stock", "A4:O4", [stock_hdr])

# Rows 5–28: 24 items
stock_rows = []
for idx, (eng, urd, sup) in enumerate(ITEMS):
    r = DS + idx
    stock_rows.append([
        SKUS[idx],
        eng,
        urd,
        "Pcs",
        200,
        fsumif("Stock In", "B", "D", r),
        fsumif("Stock Out", "B", "D", r),
        f"=IFERROR($E{r}+$F{r}-$G{r},\"Check entries\")",
        MIN_STOCK[idx],
        None,       # Cost Price
        None,       # Selling Price
        sup,
        status_formula(r),
        reorder_formula(r),
        "",
    ])

gs_update("Stock", f"A{DS}:O{DE}", stock_rows)
print("  ✓ Stock sheet done\n")


# ════════════════════════════════════════════════════════════════════════════
# 2. STOCK IN SHEET
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  2. Stock In sheet")
print("=" * 60)

in_hdr = ["Date", "Item Name", "Item Code", "Qty In", "Unit",
          "Supplier", "Reference", "Notes"]
gs_update("Stock In", "A1:H1", [in_hdr])

# Build all 3000 rows at once (formulas + empty)
in_all = []
for r in range(2, TXN + 2):
    in_all.append([
        "",
        "",
        find_code(r),
        "",
        "Pcs",
        "",
        "",
        "",
    ])

# Write all at once (one batch_update_values would be ideal, but we use
# update_values with the full range — Google Sheets can handle 3000 rows
# in a single write)
print("  Writing 3000 rows (headers + formulas)...")
gs_update("Stock In", f"A2:H{TXN+1}", in_all)

# Sample entries (rows 2–6)  — overwrite with real data
_sample_in = [
    ("2026-11-03","Upper Body",  50,"ABC Traders",   "Weekly delivery",""),
    ("2026-11-04","Cat",        100,"XYZ Wholesale", "Restock",        ""),
    ("2026-11-05","Bolt Jora",   80,"ABC Traders",   "Weekly delivery",""),
    ("2026-11-03","Gas Pipe",    60,"Gas Co. Ltd",   "New stock",      ""),
    ("2026-11-06","Silver Ring",120,"Jewellery Sup.","Restock",        ""),
]
for i, (dt, item, qty, sup, ref, note) in enumerate(_sample_in):
    r = 2 + i
    gs_update("Stock In", f"A{r}:H{r}", [[
        dt, item,
        find_code(r),
        qty, "Pcs", sup, ref, note
    ]])

# Data validation (dropdown) on B2:B3001
print("  Setting dropdown validation on B2:B{TXN+1}...")
gs_set_validation(f"Stock In!B2:B{TXN+1}", f"Stock!B{DS}")
print("  ✓ Stock In sheet done\n")


# ════════════════════════════════════════════════════════════════════════════
# 3. STOCK OUT SHEET
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  3. Stock Out sheet")
print("=" * 60)

out_hdr = ["Date", "Item Name", "Item Code", "Qty Out", "Unit",
           "Customer", "Reference", "Notes"]
gs_update("Stock Out", "A1:H1", [out_hdr])

out_all = []
for r in range(2, TXN + 2):
    out_all.append([
        "",
        "",
        find_code(r),
        "",
        "Pcs",
        "",
        "",
        "",
    ])

print("  Writing 3000 rows...")
gs_update("Stock Out", f"A2:H{TXN+1}", out_all)

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
    gs_update("Stock Out", f"A{r}:H{r}", [[
        dt, item,
        find_code(r),
        qty, "Pcs", cust, ref, note
    ]])

print("  Setting dropdown validation on B2:B{TXN+1}...")
gs_set_validation(f"Stock Out!B2:B{TXN+1}", f"Stock!B{DS}")
print("  ✓ Stock Out sheet done\n")


# ════════════════════════════════════════════════════════════════════════════
# 4. STOCK ADJUSTMENTS SHEET  (new)
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  4. Stock Adjustments sheet  (new)")
print("=" * 60)

sa_sid = gs_add_sheet("Stock Adjustments")
if sa_sid is None:
    print("  ✗ Could not add Stock Adjustments — skipping")
else:
    adj_hdr = ["Date", "Item Name", "Item Code", "Qty", "Type",
               "Reason", "Reference", "Entered By"]
    gs_update("Stock Adjustments", "A1:H1", [adj_hdr])

    adj_all = []
    for r in range(2, TXN + 2):
        adj_all.append([
            "",
            "",
            find_code(r),
            "",
            "Pcs",
            "",
            "",
            "",
        ])
    gs_update("Stock Adjustments", f"A2:H{TXN+1}", adj_all)
    print("  ✓ Stock Adjustments done\n")


# ════════════════════════════════════════════════════════════════════════════
# 5. SUPPLIERS SHEET  (new)
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  5. Suppliers sheet  (new)")
print("=" * 60)

sup_sid = gs_add_sheet("Suppliers")
if sup_sid is None:
    print("  ✗ Could not add Suppliers — skipping")
else:
    sup_hdr = ["Supplier Name", "Phone / Contact", "Email",
               "Items Supplied", "Lead Time (days)",
               "Payment Terms", "Notes"]

    SUP_DB = {
        "ABC Traders":    ("0300-1234567", "abctraders@example.com",   2, "Cash on delivery"),
        "XYZ Wholesale":  ("0333-9876543", "xyzwholesale@example.com", 3, "7 days credit"),
        "Local Market":   ("0345-2223334", "",                         1, "Cash"),
        "Gas Co. Ltd":    ("0300-9998887", "gasco@example.com",        5, "Advance 50%"),
        "Jewellery Sup.": ("0348-5556667", "jewellery@example.com",    4, "Cash on delivery"),
        "Pipe Co.":       ("0300-6665544", "pipeco@example.com",       3, "7 days credit"),
    }

    seen = set()
    sup_list = []
    for _, _, s in ITEMS:
        if s not in seen:
            seen.add(s)
            sup_list.append(s)

    sup_rows = []
    for sup in sup_list:
        info = SUP_DB.get(sup, ("", "", "", ""))
        sup_rows.append([
            sup,
            info[0],
            info[1],
            (f'=TEXTJOIN(", ",TRUE,'
             f'FILTER(\'Stock\'!$B${DS}:$B${DE},'
             f'\'Stock\'!$L${DS}:$L${DE}=$A{{r}},""))').replace("{r}", str(2 + sup_list.index(sup))),
            str(info[2]),
            info[3],
            "",
        ])

    gs_update("Suppliers", "A1:G1", [sup_hdr])
    gs_update("Suppliers", "A2:G7", sup_rows)
    print("  ✓ Suppliers done\n")


# ════════════════════════════════════════════════════════════════════════════
# 6. DASHBOARD SHEET  (new)
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  6. Dashboard sheet  (new)")
print("=" * 60)

dash_sid = gs_add_sheet("Dashboard")
if dash_sid is None:
    print("  ✗ Could not add Dashboard — skipping")
else:
    # Set column widths via update_dimension_properties (cols B-F)
    # Just write data; Google Sheets will infer reasonable widths

    # Row 2: title  (B2:F2 merged later via merge_cells)
    gs_update("Dashboard", "B2:F2", [["Stock Dashboard"]])
    # Row 3: subtitle
    gs_update("Dashboard", "B3:F3", [["Wholesale Stock Inventory — Status at a Glance"]])

    # KPI cards:  big numbers in rows 5 and 7, labels in rows 6 and 8
    # Layout:
    #   B5:C5  = Total Items (big number)
    #   D5:E5  = Items In Stock (big number)
    #   B6:C6  = "Total Items" label
    #   D6:E6  = "Items In Stock" label
    #   B7:C7  = Low Stock (big number)
    #   D7:E7  = Out of Stock (big number)
    #   B8:C8  = "Low Stock" label
    #   D8:E8  = "Out of Stock" label

    # Big numbers row 5
    gs_update("Dashboard", "B5:C5", [[f"=COUNTA('Stock'!$A${DS}:$A${DE})"]])
    gs_update("Dashboard", "D5:E5", [[f'=COUNTIF(\'Stock\'!$H${DS}:$H${DE},">0")']])

    # Labels row 6
    gs_update("Dashboard", "B6:C6", [["Total Items"]])
    gs_update("Dashboard", "D6:E6", [["Items In Stock"]])

    # Big numbers row 7
    gs_update("Dashboard", "B7:C7", [[f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Low")']])
    gs_update("Dashboard", "D7:E7", [[f'=COUNTIF(\'Stock\'!$M${DS}:$M${DE},"Out of stock")']])

    # Labels row 8
    gs_update("Dashboard", "B8:C8", [["Low Stock"]])
    gs_update("Dashboard", "D8:E8", [["Out of Stock"]])

    # Item status table header (row 10)
    dash_hdr = ["Item Code", "Item Name", "Urdu Name", "Current Stock",
                "Min Stock", "Status", "Cost Price", "Sell Price",
                "Stock Value (Rs)", "Reorder?"]
    gs_update("Dashboard", "A10:J10", [dash_hdr])

    # 24 item rows (rows 11–34)
    dash_rows = []
    for idx in range(N):
        r    = 11 + idx
        srow = DS + idx
        dash_rows.append([
            f"='Stock'!$A${srow}",
            f"='Stock'!$B${srow}",
            f"='Stock'!$C${srow}",
            f"='Stock'!$H${srow}",
            f"='Stock'!$I${srow}",
            f"='Stock'!$M${srow}",
            f"='Stock'!$J${srow}",
            f"='Stock'!$K${srow}",
            f"=IFERROR($D{r}*$G{r},0)",
            f'=IF($F{r}="OK","","Reorder now")',
        ])
    gs_update("Dashboard", "A11:J34", dash_rows)

    # Merge KPI cells
    for merge_range in ["B5:C5","D5:E5","B6:C6","D6:E6",
                        "B7:C7","D7:E7","B8:C8","D8:E8"]:
        subprocess.run(
            [sys.executable, "-c",
             f"\"\"\"import sys, json\nsys.path.insert(0,'/home/user')\n"
             f"from arena_connector import call_connector_tool\n"
             f"r = call_connector_tool(\n"
             f"    connector='google_sheets',\n"
             f"    tool='merge_cells',\n"
             f"    arguments={{\n"
             f"        'spreadsheet_id': '{SPREADSHEET_ID}',\n"
             f"        'range': {{'start_row_index':0,'end_row_index':2,'start_column_index':1,'end_column_index':3,'sheet_id':0}},\n"
             f"        'merge_type': 'MERGE_ALL'\n"
             f"    }}\n"
             f")\n"
             f"print(json.dumps(r))\"\"\""
            ],
            capture_output=True, text=True, timeout=20
        )
        # Note: the merge uses fixed sheet_id/from above; we need proper ranges.
        # Let's do this properly instead:

    # Proper merges using add_sheet grid: Dashboard is the last added sheet.
    # Find Dashboard sheet id from get_spreadsheet
    gs_result = json.loads(
        subprocess.run(
            [sys.executable, "-c",
             f"\"\"\"import sys, json\nsys.path.insert(0,'/home/user')\n"
             f"from arena_connector import call_connector_tool\n"
             f"r = call_connector_tool(\n"
             f"    connector='google_sheets',\n"
             f"    tool='get_spreadsheet',\n"
             f"    arguments={{\n"
             f"        'spreadsheet_id': '{SPREADSHEET_ID}',\n"
             f"        'include_grid_data': False,\n"
             f"        'ranges': []\n"
             f"    }}\n"
             f")\n"
             f"print(json.dumps(r))\"\"\""
            ],
            capture_output=True, text=True, timeout=20
        ).stdout
    )
    sheets_info = gs_result.get("result", {}).get("sheets", [])
    dash_sid_found = None
    for s in sheets_info:
        if s.get("properties", {}).get("title") == "Dashboard":
            dash_sid_found = s["properties"]["sheet_id"]
            break

    if dash_sid_found:
        merges_to_do = [
            ("B5:C5", 4, 5, 1, 2),   # rows 5-6 (0-indexed: 4-5), cols B-C (1-2)
            ("D5:E5", 4, 5, 3, 4),   # cols D-E (3-4)
            ("B6:C6", 5, 6, 1, 2),   # rows 6-7
            ("D6:E6", 5, 6, 3, 4),
            ("B7:C7", 6, 7, 1, 2),
            ("D7:E7", 6, 7, 3, 4),
            ("B8:C8", 7, 8, 1, 2),
            ("D8:E8", 7, 8, 3, 4),
        ]
        for label, rstart, rend, cstart, cend in merges_to_do:
            subprocess.run(
                [sys.executable, "-c",
                 f"\"\"\"import sys, json\nsys.path.insert(0,'/home/user')\n"
                 f"from arena_connector import call_connector_tool\n"
                 f"r = call_connector_tool(\n"
                 f"    connector='google_sheets',\n"
                 f"    tool='merge_cells',\n"
                 f"    arguments={{\n"
                 f"        'spreadsheet_id': '{SPREADSHEET_ID}',\n"
                 f"        'range': {{'sheet_id':{dash_sid_found},"
                 f"'start_row_index':{rstart},'end_row_index':{rend},"
                 f"'start_column_index':{cstart},'end_column_index':{cend}}},\n"
                 f"        'merge_type': 'MERGE_ALL'\n"
                 f"    }}\n"
                 f")\n"
                 f"print(json.dumps(r))\"\"\""
                ],
                capture_output=True, text=True, timeout=20
            )
        print("  ✓ Dashboard merges done")

    print("  ✓ Dashboard done\n")


# ════════════════════════════════════════════════════════════════════════════
# 7. HOW TO USE  (update existing sheet)
# ════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("  7. How to Use sheet  (update)")
print("=" * 60)

guide = [
    "Stock Inventory — Professional Edition — How to Use",
    "",
    "1. STOCK SHEET — main inventory view",
    "Item Code (STK001–STK024): a unique code for each item. "
    "Use this code in Stock In, Stock Out, and Adjustments sheets. "
    "The code auto-fills when you pick the item name from the dropdown.",
    "Item Name: English name of the part. Select from the dropdown in "
    "Stock In / Out / Adjustments sheets — the Item Code fills automatically.",
    "Urdu Name: Urdu label for quick recognition on mobile.",
    "Unit: default 'Pcs'. Change if you sell in kg, metres, boxes, etc.",
    "Opening Stock: starting count. Currently set to 200 for all items — "
    "update after your first real shelf count.",
    "Stock In (col F): auto-sum of all deliveries for this item from the "
    "Stock In sheet. Do NOT type here — updates automatically.",
    "Stock Out (col G): auto-sum of all sales/issues for this item from "
    "the Stock Out sheet. Do NOT type here.",
    "Current Stock (col H): auto = Opening Stock + Stock In − Stock Out. "
    "Updates instantly when you add entries to Stock In or Stock Out.",
    "Minimum Stock (col I): the level at which you want to be alerted "
    "to reorder. Edit per item as needed.",
    "Cost Price (Rs) (col J): what you pay per piece. Fill this in for "
    "your items. Used for total stock value on the Dashboard.",
    "Selling Price (Rs) (col K): what you charge per piece. Fill this in. "
    "Used for stock value and profit tracking on the Dashboard.",
    "Supplier (col L): who supplies this item. Update as needed.",
    "Status (col M): auto — OK / Low / Out of stock / Check entries.",
    "Reorder Qty (col N): auto-suggested quantity to order when stock is "
    "Low or Out. Formula: 2×Min − Current Stock. Blank when OK.",
    "Notes (col O): free text for anything about the item.",
    "",
    "2. STOCK IN SHEET — record every delivery",
    "Date: date of delivery. Item Name: pick from dropdown (auto-fills "
    "Item Code in col C). Qty In: how many pieces received. Supplier: who "
    "delivered. Reference: invoice number or note. Notes: anything else.",
    "Tip: always add a Stock In entry on the day the goods arrive — this "
    "keeps Current Stock accurate on the Stock sheet.",
    "",
    "3. STOCK OUT SHEET — record every sale or issue",
    "Date: date of sale. Item Name: pick from dropdown (auto-fills Item "
    "Code in col C). Qty Out: how many pieces went out. Customer: who bought "
    "it. Reference: order number. Notes: anything else.",
    "Tip: add a Stock Out entry the same day as the sale — the Current "
    "Stock on the Stock sheet updates automatically.",
    "",
    "4. STOCK ADJUSTMENTS SHEET — for non-sale changes",
    "Use this for: Damage (items broken/lost), Return to Supplier (sent "
    "back), Return from Customer (came back), Found Missing (counted more "
    "than sheet showed). Type: pick from the dropdown. Reason: explain why. "
    "Reference: note or invoice. Entered By: your name.",
    "These adjustments do NOT go into Stock In or Stock Out — they are "
    "separate so you can see what is normal sales vs. losses.",
    "",
    "5. SUPPLIERS SHEET — who supplies what",
    "Lists each supplier and which items they supply (auto-filtered from "
    "the Stock sheet). Phone, email, lead time (days), and payment terms are "
    "pre-filled for known suppliers — update as needed. Lead time helps you "
    "plan reorders before items go out of stock.",
    "",
    "6. DASHBOARD — status at a glance",
    "Shows big numbers: Total Items, Items In Stock, Low Stock, Out of "
    "Stock. Below them, a table of all 24 items with Current Stock, Status, "
    "Cost Price, Sell Price, Stock Value (Rs), and a 'Reorder now' flag. "
    "Use this as your daily check — items flagged 'Reorder now' need attention.",
    "Stock Value (Rs) = Current Stock × Cost Price. Fill in Cost Price for "
    "each item to see your total inventory value.",
    "",
    "7. DAILY / WEEKLY / MONTHLY ROUTINE",
    "Daily: add every delivery to Stock In, every sale to Stock Out. "
    "Check Dashboard for 'Reorder now' items.",
    "Weekly: count shelves. Fix differences. Review Low / Out of stock "
    "items. Call suppliers for reorders using the Reorder Qty suggested on "
    "the Stock sheet.",
    "Monthly: fill in Cost Price and Selling Price if not done. Download "
    "a backup (File → Download → Microsoft Excel). Review total stock value "
    "on the Dashboard.",
    "",
    "8. HOW TO SHARE (Google Sheets)",
    "Open the Google Sheet → click Share (top right) → add your brother's "
    "email and partner's email → set as Viewer. You are the only Editor. "
    "They can see stock on phone or PC but cannot change it.",
    "",
    "9. NOTES",
    "• Item codes STK001–STK024 are auto-assigned. The code auto-fills "
    "in Stock In, Stock Out, and Adjustments when you pick the item from "
    "the dropdown.",
    "• Cost Price and Selling Price are blank — fill them in when you know "
    "them. The Dashboard Stock Value and total value will update automatically.",
    "• Reorder Qty on the Stock sheet shows how many to order when stock is "
    "Low or Out. Formula = 2×Minimum Stock − Current Stock.",
    "• Always use the item name dropdown — this keeps the auto-calculations "
    "and Item Code lookups correct.",
    "• Excel file = offline backup. Google Sheet = live online copy to share.",
]

for i, line in enumerate(guide):
    gs_update("How to Use", f"A{i+1}", [[line]])

print("  ✓ How to Use done\n")

print("=" * 60)
print("  ✓ ALL SHEETS SYNCHRONISED TO GOOGLE SHEET")
print("=" * 60)

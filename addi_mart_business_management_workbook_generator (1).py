import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, PieChart, Reference
from datetime import date, timedelta

wb = openpyxl.Workbook()

# Color Palette - Muted Earthy Forest & Warm Terracotta/Cream (Addi Mart theme)
HEADER_FILL = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="2E5A36", end_color="2E5A36", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F5F7F5", end_color="F5F7F5", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="E6ECE7", end_color="E6ECE7", fill_type="solid")
KPI_BG_FILL = PatternFill(start_color="FAF8F5", end_color="FAF8F5", fill_type="solid")

# Conditional Formatting Fills
OK_FILL = PatternFill(start_color="DDF3E4", end_color="DDF3E4", fill_type="solid")
OK_FONT = Font(name="Calibri", size=10, bold=True, color="1E5631")
WARN_FILL = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
WARN_FONT = Font(name="Calibri", size=10, bold=True, color="856404")
CRIT_FILL = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
CRIT_FONT = Font(name="Calibri", size=10, bold=True, color="842029")

TITLE_FONT = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="E0E8E0")
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=10, color="1F2937")
BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="1F2937")
KPI_LABEL_FONT = Font(name="Calibri", size=9, bold=True, color="4B5563")
KPI_VAL_FONT = Font(name="Calibri", size=14, bold=True, color="1E4620")

THIN_GRAY = Side(style="thin", color="D1D5DB")
MED_GREEN = Side(style="medium", color="1E4620")
DOUBLE_BOTTOM = Side(style="double", color="1E4620")

CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
TOTAL_BORDER = Border(top=MED_GREEN, bottom=DOUBLE_BOTTOM, left=THIN_GRAY, right=THIN_GRAY)
KPI_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=MED_GREEN, bottom=THIN_GRAY)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

FMT_ETB = '#,##0.00 "ETB"'
FMT_QTY = '#,##0.0'
FMT_INT = '#,##0'
FMT_PCT = '0.0%'
FMT_DATE = 'yyyy-mm-dd'

def style_banner(ws, title, subtitle, max_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    c1 = ws.cell(row=1, column=1, value=title)
    c2 = ws.cell(row=2, column=1, value=subtitle)
    c1.font = TITLE_FONT
    c2.font = SUBTITLE_FONT
    c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for r in (1, 2):
        for col in range(1, max_col + 1):
            ws.cell(row=r, column=col).fill = HEADER_FILL
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 8

# Tab 1: 01_Dashboard
ws_dash = wb.active
ws_dash.title = "01_Dashboard"
ws_dash.views.sheetView[0].showGridLines = True
style_banner(ws_dash, "ADDI MART (ዓዲ ማርት) — EXECUTIVE OPERATIONS DASHBOARD",
             "Mekelle Grocery Retail | Revenue, Margin, Payment Mix, Cashflow & Reorder Alerts", 9)

# Tab 2: 02_Inventory_Master
ws_inv = wb.create_sheet("02_Inventory_Master")
ws_inv.views.sheetView[0].showGridLines = True
inv_headers = [
    "SKU_ID", "Category", "Item_Name (EN / ትግርኛ)", "Unit",
    "Opening_Stock", "Stock_In", "Stock_Out", "Current_Stock",
    "Reorder_Point", "Stock_Status", "Unit_Cost_ETB",
    "Selling_Price_ETB", "Gross_Margin_%", "Stock_Value_ETB", "Featured"
]
style_banner(ws_inv, "ADDI MART (ዓዲ ማርት) — MASTER INVENTORY & PRICING",
             "Live Stock Balances, Reorder Thresholds, Landed Unit Costs & Gross Margins", len(inv_headers))

ws_inv.row_dimensions[4].height = 26
for idx, h in enumerate(inv_headers, 1):
    cell = ws_inv.cell(row=4, column=idx, value=h)
    cell.font = HEADER_FONT
    cell.fill = SUBHEADER_FILL
    cell.alignment = ALIGN_CENTER
    cell.border = CELL_BORDER

sample_items = [
    ("AM-GR-001", "Grains & Flour", "White Teff (ጻዕዳ ጣፍ)", "kg", 150, 30, 104.00, 118.00, "Yes"),
    ("AM-GR-002", "Grains & Flour", "Red Teff (ቀይሕ ጣፍ)", "kg", 100, 25, 92.00, 105.00, "No"),
    ("AM-GR-003", "Grains & Flour", "Wheat Flour (ሓሪጭ ስርናይ)", "kg", 120, 25, 68.00, 78.00, "Yes"),
    ("AM-GR-004", "Grains & Flour", "Sorghum / Mashila (መሸላ)", "kg", 80, 20, 54.00, 64.00, "No"),
    ("AM-OL-001", "Oils & Sweeteners", "Sunflower Cooking Oil 5L (ዘይቲ)", "Bottle (5L)", 40, 10, 940.00, 1040.00, "Yes"),
    ("AM-OL-002", "Oils & Sweeteners", "Refined Sugar (ሽኮር)", "kg", 90, 20, 88.00, 98.00, "Yes"),
    ("AM-SP-001", "Spices & Legumes", "Mekelle Berbere (በርበረ)", "kg", 45, 10, 380.00, 450.00, "Yes"),
    ("AM-SP-002", "Spices & Legumes", "Roasted Shiro Powder (ሽሮ)", "kg", 50, 12, 165.00, 198.00, "Yes"),
    ("AM-SP-003", "Spices & Legumes", "Split Red Lentils (ምስር ክክ)", "kg", 60, 15, 132.00, 155.00, "No"),
    ("AM-PR-001", "Fresh Produce", "Fresh Tomatoes (ኮሚደረ)", "kg", 35, 12, 34.00, 45.00, "Yes"),
    ("AM-PR-002", "Fresh Produce", "Red Onions (ቀይሕ ሽጉርቲ)", "kg", 55, 15, 48.00, 62.00, "Yes"),
    ("AM-PR-003", "Fresh Produce", "Potatoes (ድንሽ)", "kg", 70, 20, 28.00, 38.00, "No"),
    ("AM-DY-001", "Dairy & Eggs", "Farm Eggs - Tray of 30 (እንቋቑሖ)", "Tray (30)", 25, 8, 310.00, 365.00, "Yes"),
    ("AM-DY-002", "Dairy & Eggs", "Pasteurized Milk 1L (ጸባ)", "Liter", 30, 10, 65.00, 80.00, "No"),
    ("AM-HH-001", "Household & Beverage", "Ethiopian Coffee Beans (ቡን)", "kg", 30, 8, 440.00, 520.00, "Yes"),
    ("AM-HH-002", "Household & Beverage", "Laundry Bar Soap (ሳሙና)", "Piece", 60, 15, 42.00, 52.00, "No"),
]

start_row = 5
for idx, item in enumerate(sample_items):
    r = start_row + idx
    ws_inv.row_dimensions[r].height = 21
    sku, cat, name, unit, open_stk, reord, cost, price, feat = item
    ws_inv.cell(row=r, column=1, value=sku).alignment = ALIGN_CENTER
    ws_inv.cell(row=r, column=2, value=cat).alignment = ALIGN_LEFT
    ws_inv.cell(row=r, column=3, value=name).alignment = ALIGN_LEFT
    ws_inv.cell(row=r, column=4, value=unit).alignment = ALIGN_CENTER
    ws_inv.cell(row=r, column=5, value=open_stk).number_format = FMT_QTY
    ws_inv.cell(row=r, column=6, value=f"=SUMIFS('03_Purchases_Log'!$E$5:$E$200, '03_Purchases_Log'!$C$5:$C$200, A{r})").number_format = FMT_QTY
    ws_inv.cell(row=r, column=7, value=f'=SUMIFS(\'04_Orders_&_Sales\'!$G$5:$G$500, \'04_Orders_&_Sales\'!$E$5:$E$500, A{r}, \'04_Orders_&_Sales\'!$K$5:$K$500, "<>Cancelled")').number_format = FMT_QTY
    ws_inv.cell(row=r, column=8, value=f"=E{r}+F{r}-G{r}").number_format = FMT_QTY
    ws_inv.cell(row=r, column=9, value=reord).number_format = FMT_QTY
    ws_inv.cell(row=r, column=10, value=f'=IF(H{r}<=0,"OUT OF STOCK",IF(H{r}<=I{r},"REORDER NOW","OK"))').alignment = ALIGN_CENTER
    ws_inv.cell(row=r, column=11, value=cost).number_format = FMT_ETB
    ws_inv.cell(row=r, column=12, value=price).number_format = FMT_ETB
    ws_inv.cell(row=r, column=13, value=f"=IF(L{r}>0,(L{r}-K{r})/L{r},0)").number_format = FMT_PCT
    ws_inv.cell(row=r, column=14, value=f"=H{r}*K{r}").number_format = FMT_ETB
    ws_inv.cell(row=r, column=15, value=feat).alignment = ALIGN_CENTER
    for col_i in range(1, len(inv_headers) + 1):
        c = ws_inv.cell(row=r, column=col_i)
        c.font = BOLD_FONT if col_i == 8 else BODY_FONT
        c.border = CELL_BORDER

# Tab 3: 03_Purchases_Log
ws_pur = wb.create_sheet("03_Purchases_Log")
ws_pur.views.sheetView[0].showGridLines = True
pur_headers = [
    "Purchase_ID", "Date", "SKU_ID", "Item_Name (Auto)", "Qty_Purchased",
    "Wholesale_Cost_ETB", "Bajaj_Transport_ETB", "Total_Cost_ETB",
    "Effective_Unit_Cost", "Supplier_Market", "Payment_Method"
]
style_banner(ws_pur, "ADDI MART (ዓዲ ማርት) — SUPPLIER & WHOLESALE PURCHASES LOG",
             "Mekelle Wholesale Restocks & Landed Transport Cost Allocation", len(pur_headers))
ws_pur.row_dimensions[4].height = 26
for idx, h in enumerate(pur_headers, 1):
    c = ws_pur.cell(row=4, column=idx, value=h)
    c.font = HEADER_FONT
    c.fill = SUBHEADER_FILL
    c.alignment = ALIGN_CENTER
    c.border = CELL_BORDER

# Tab 4: 04_Orders_&_Sales
ws_ord = wb.create_sheet("04_Orders_&_Sales")
ws_ord.views.sheetView[0].showGridLines = True
ord_headers = [
    "Order_ID", "Date", "Channel", "Sub_City / Area", "SKU_ID",
    "Item_Name (Auto)", "Qty_Sold", "Unit_Price_ETB", "Delivery_Fee_ETB",
    "Total_Revenue_ETB", "Order_Status", "Payment_Method", "Est_Gross_Profit_ETB",
    "Customer_Phone", "Customer_Chat_ID", "Notes"
]
style_banner(ws_ord, "ADDI MART (ዓዲ ማርት) — ORDERS, SALES & TELEGRAM BOT LEDGER",
             "Synced with @AddiMartOrderBot (Google Apps Script) & Walk-In Sales", len(ord_headers))
ws_ord.row_dimensions[4].height = 26
for idx, h in enumerate(ord_headers, 1):
    c = ws_ord.cell(row=4, column=idx, value=h)
    c.font = HEADER_FONT
    c.fill = SUBHEADER_FILL
    c.alignment = ALIGN_CENTER
    c.border = CELL_BORDER

# Tab 5: 05_Expenses_&_Cashflow
ws_exp = wb.create_sheet("05_Expenses_&_Cashflow")
ws_exp.views.sheetView[0].showGridLines = True
exp_headers = ["Expense_ID", "Date", "Category", "Description", "Amount_ETB", "Paid_Via", "Approved_By"]
style_banner(ws_exp, "ADDI MART (ዓዲ ማርት) — OPERATING EXPENSES & CASHFLOW LOG",
             "Rent, Staff Payroll, Bajaj Rider Stipends, Utilities & Marketing", len(exp_headers))
ws_exp.row_dimensions[4].height = 26
for idx, h in enumerate(exp_headers, 1):
    c = ws_exp.cell(row=4, column=idx, value=h)
    c.font = HEADER_FONT
    c.fill = SUBHEADER_FILL
    c.alignment = ALIGN_CENTER
    c.border = CELL_BORDER

# Tab 6: 06_Suppliers_&_Riders
ws_sup = wb.create_sheet("06_Suppliers_&_Riders")
ws_sup.views.sheetView[0].showGridLines = True
sup_headers = ["Partner_ID", "Type", "Name", "Hub / Sub-City", "Phone", "Payment_Details", "Notes"]
style_banner(ws_sup, "ADDI MART (ዓዲ ማርት) — SUPPLIERS & BAJAJ DELIVERY RIDERS",
             "Mekelle Wholesale Directory & Sub-City Dispatch Coverage", len(sup_headers))
ws_sup.row_dimensions[4].height = 26
for idx, h in enumerate(sup_headers, 1):
    c = ws_sup.cell(row=4, column=idx, value=h)
    c.font = HEADER_FONT
    c.fill = SUBHEADER_FILL
    c.alignment = ALIGN_CENTER
    c.border = CELL_BORDER

wb.save("Addi_Mart_Business_Management_Workbook.xlsx")
print("Generated Addi_Mart_Business_Management_Workbook.xlsx")

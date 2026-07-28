import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import struct, zlib, zipfile, shutil, os, io

# ─── COLOURS ────────────────────────────────────────────────────────────────
NAVY    = "1F3864"; MID     = "2E75B6"; LIGHT   = "BDD7EE"
GREEN   = "E2EFDA"; RED     = "FFE5E5"; AMBER   = "FFF2CC"
WHITE   = "FFFFFF"; DKGRAY  = "404040"; LGRAY   = "F2F2F2"
BBLUE   = "0000FF"   # input colour (blue text)
BGREEN  = "008000"   # cross-sheet link colour (green text)

# ─── HELPERS ────────────────────────────────────────────────────────────────
thin = Side(style="thin", color="BFBFBF")
med  = Side(style="medium", color="808080")
def bdr(t=False,b=False,l=False,r=False):
    return Border(
        top=med if t else thin, bottom=med if b else thin,
        left=med if l else thin, right=med if r else thin)

def hdr(ws, row, col, val, bg=NAVY, fg=WHITE, bold=True, sz=11, merge=None, wrap=False, align="center"):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name="Arial", bold=bold, color=fg, size=sz)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if merge:
        ws.merge_cells(start_row=row, start_column=col, end_row=row+merge[0]-1, end_column=col+merge[1]-1)
    return c

def inp(ws, row, col, val, fmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name="Arial", color=BBLUE, bold=bold, size=10)
    c.alignment = Alignment(horizontal="right", vertical="center")
    if fmt: c.number_format = fmt
    return c

def lnk(ws, row, col, formula, fmt=None):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = Font(name="Arial", color=BGREEN, size=10)
    c.alignment = Alignment(horizontal="right", vertical="center")
    if fmt: c.number_format = fmt
    return c

def frm(ws, row, col, formula, fmt=None, bold=False, color="000000"):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = Font(name="Arial", color=color, bold=bold, size=10)
    c.alignment = Alignment(horizontal="right", vertical="center")
    if fmt: c.number_format = fmt
    return c

def lbl(ws, row, col, val, bold=False, indent=0, bg=None, color="000000"):
    c = ws.cell(row=row, column=col, value=(" " * indent * 2) + str(val))
    c.font = Font(name="Arial", bold=bold, color=color, size=10)
    c.alignment = Alignment(horizontal="left", vertical="center")
    if bg: c.fill = PatternFill("solid", start_color=bg)
    return c

def sec(ws, row, cols=8, label="", bg=MID, fg=WHITE):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    c = ws.cell(row=row, column=1, value=label)
    c.font = Font(name="Arial", bold=True, color=fg, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    return row + 1

NUM  = '#,##0'
NUM2 = '#,##0.00'
PCT  = '0.0%'
PCT1 = '0.00%'
DOLR = '$#,##0'
DOL2 = '$#,##0.00'

print("Helpers loaded")

# ─── WORKBOOK SETUP ─────────────────────────────────────────────────────────
wb = Workbook()
sheets = ["Cover","Drivers","Revenue Model","Cost Model",
          "P&L Monthly","Cash Flow","Working Capital","Budget Variance",
          "Scenarios","Risk & Opps","CFO Dashboard","Macro Guide"]
ws_cover = wb.active; ws_cover.title = "Cover"
for s in sheets[1:]:
    wb.create_sheet(s)

ws_cov  = wb["Cover"]
ws_drv  = wb["Drivers"]
ws_rev  = wb["Revenue Model"]
ws_cst  = wb["Cost Model"]
ws_pl   = wb["P&L Monthly"]
ws_cf   = wb["Cash Flow"]
ws_wc   = wb["Working Capital"]
ws_bv   = wb["Budget Variance"]
ws_sc   = wb["Scenarios"]
ws_ro   = wb["Risk & Opps"]
ws_dash = wb["CFO Dashboard"]
ws_mg   = wb["Macro Guide"]

# Tab colours
tab_colors = {
    "Cover":"1F3864","Drivers":"2E75B6","Revenue Model":"70AD47",
    "Cost Model":"ED7D31","P&L Monthly":"4472C4","Cash Flow":"7030A0",
    "Working Capital":"00B0F0","Budget Variance":"FFC000",
    "Scenarios":"FF0000","Risk & Opps":"FF7F7F",
    "CFO Dashboard":"1F3864","Macro Guide":"808080"}
for name, color in tab_colors.items():
    wb[name].sheet_properties.tabColor = color

print("Workbook structure created")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 1 – COVER
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_cov
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 40
ws.column_dimensions['C'].width = 30
for r in range(1, 50):
    ws.row_dimensions[r].height = 18

# Title block
ws.merge_cells("B2:E4")
c = ws["B2"]
c.value = "DRIVER-BASED FORECAST MODEL"
c.font = Font(name="Arial", bold=True, size=22, color=WHITE)
c.fill = PatternFill("solid", start_color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")

ws.merge_cells("B5:E5")
c = ws["B5"]
c.value = "Q3 2026 Monthly Projections | Jul–Aug–Sep 2026"
c.font = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill = PatternFill("solid", start_color=MID)
c.alignment = Alignment(horizontal="center", vertical="center")

ws.merge_cells("B6:E6")
c = ws["B6"]
c.value = "Baseline: May 2026 Actuals  |  Prepared: July 2026"
c.font = Font(name="Arial", size=10, color=DKGRAY)
c.fill = PatternFill("solid", start_color=LIGHT)
c.alignment = Alignment(horizontal="center", vertical="center")

# Key stats
stats = [
    ("May 2026 Revenue", "$40,546,176", MID),
    ("Net Income Margin", "22.4%", "006400"),
    ("EBITDA Margin", "27.8%", "006400"),
    ("Overdue AR", "$3,326,800", "C00000"),
    ("Cash on Hand", "$12,911,223", MID),
    ("Prepay Amortization/Mo", "$431,000", "ED7D31"),
]
r = 8
for label, val, color in stats:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value=label)
    c.font = Font(name="Arial", bold=True, size=10, color=DKGRAY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.fill = PatternFill("solid", start_color=LGRAY)
    c2 = ws.cell(row=r, column=4, value=val)
    c2.font = Font(name="Arial", bold=True, size=11, color=color)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = PatternFill("solid", start_color=LGRAY)
    r += 1

# Sheet index
r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
c = ws.cell(row=r, column=2, value="WORKBOOK NAVIGATION")
c.font = Font(name="Arial", bold=True, size=11, color=WHITE)
c.fill = PatternFill("solid", start_color=NAVY)
c.alignment = Alignment(horizontal="center", vertical="center")
r += 1

nav_items = [
    ("Drivers", "Scenario control, key assumptions & driver inputs"),
    ("Revenue Model", "Volume × price × seasonality drivers by BU"),
    ("Cost Model", "Fixed/variable cost decomposition & driver calcs"),
    ("P&L Monthly", "Jul / Aug / Sep accounting P&L — all 3 scenarios"),
    ("Cash Flow", "Operating cash vs. accounting income reconciliation"),
    ("Working Capital", "DSO, AR collection timing, DPO, WC movements"),
    ("Budget Variance", "Actual vs. budget analysis by cost centre"),
    ("Scenarios", "Side-by-side scenario comparison & sensitivity"),
    ("Risk & Opps", "Risk register with quantified P&L/cash impact"),
    ("CFO Dashboard", "Executive KPI summary for CFO review"),
    ("Macro Guide", "VBA macro reference & user instructions"),
]
for sheet_name, desc in nav_items:
    ws.cell(row=r, column=2, value=sheet_name).font = Font(name="Arial", bold=True, color=MID, size=10, underline="single")
    ws.cell(row=r, column=2).hyperlink = f"#{sheet_name}!A1"
    ws.cell(row=r, column=3, value=desc).font = Font(name="Arial", size=10, color=DKGRAY)
    ws.cell(row=r, column=3).alignment = Alignment(horizontal="left")
    r += 1

# Legend
r += 1
legend = [
    (BBLUE, WHITE, "Blue text = Hardcoded input — change for scenarios"),
    ("000000", WHITE, "Black text = Formula — do not edit"),
    (BGREEN, WHITE, "Green text = Cross-sheet link"),
    ("000000", AMBER, "Yellow background = Key assumption to review"),
]
for color, bg, text in legend:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = ws.cell(row=r, column=2, value=text)
    c.font = Font(name="Arial", size=9, color=color)
    c.fill = PatternFill("solid", start_color=(bg if bg != WHITE else "FFFFFF"))
    r += 1

print("Cover sheet done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 2 – DRIVERS  (scenario control hub)
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_drv
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 38
ws.column_dimensions['B'].width = 18  # Base
ws.column_dimensions['C'].width = 18  # Upside
ws.column_dimensions['D'].width = 18  # Risk-Adjusted
ws.column_dimensions['E'].width = 20  # Active Value (auto)
ws.column_dimensions['F'].width = 28  # Notes

def drow(ws, r, label, base, up, risk, fmt=None, note=""):
    lbl(ws, r, 1, label)
    inp(ws, r, 2, base, fmt)
    inp(ws, r, 3, up, fmt)
    inp(ws, r, 4, risk, fmt)
    frm(ws, r, 5, f'=INDEX(B{r}:D{r},1,MATCH($B$2,{{"Base","Upside","Risk-Adjusted"}},0))', fmt, color=BGREEN)
    ws.cell(row=r, column=5).font = Font(name="Arial", color=BGREEN, size=10)
    if note:
        ws.cell(row=r, column=6, value=note).font = Font(name="Arial", size=9, color="808080", italic=True)

# ─── SCENARIO SELECTOR ──────────────────────────────────────────────────────
hdr(ws, 1, 1, "DRIVER-BASED FORECAST — SCENARIO CONTROL", bg=NAVY, merge=(1,6))
ws.row_dimensions[1].height = 28

ws.cell(row=2, column=1, value="Active Scenario:").font = Font(name="Arial", bold=True, size=12, color=NAVY)
inp(ws, 2, 2, "Base", bold=True)
ws.cell(row=2, column=2).font = Font(name="Arial", bold=True, size=12, color="C00000")
ws.cell(row=2, column=2).fill = PatternFill("solid", start_color=AMBER)
ws.cell(row=2, column=2).alignment = Alignment(horizontal="center", vertical="center")

# Data validation dropdown
dv = DataValidation(type="list", formula1='"Base,Upside,Risk-Adjusted"', allow_blank=False)
dv.sqref = "B2"
ws.add_data_validation(dv)

# MATCH helper row (used by INDEX formulas)
ws.cell(row=3, column=1, value="Scenario Index (formula)").font = Font(name="Arial", size=9, color="808080", italic=True)
frm(ws, 3, 2, '=MATCH(B2,{"Base","Upside","Risk-Adjusted"},0)', fmt="0", color="808080")
ws.cell(row=3, column=2).font = Font(name="Arial", size=9, color="808080", italic=True)

# Column headers
r = 5
for col, txt, bg in [(1,"Driver / Assumption",NAVY),(2,"Base",MID),(3,"Upside",MID),(4,"Risk-Adjusted","C00000"),(5,"Active Value",NAVY),(6,"Notes",NAVY)]:
    c = ws.cell(row=r, column=col, value=txt)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left", vertical="center", indent=1 if col==1 else 0)

# ─── REVENUE DRIVERS ────────────────────────────────────────────────────────
r = 6
r = sec(ws, r, 6, "REVENUE DRIVERS", NAVY)
drow(ws,r,"May 2026 Monthly Revenue Baseline ($)",40546176,40546176,40546176,DOLR,"Source: PL_Statement_May2025_Comparative.csv"); r+=1
drow(ws,r,"YoY Revenue Growth Rate",0.088,0.20,0.045,PCT,"Base=May26 YoY avg; Upside=+20%; Risk=muted growth"); r+=1
drow(ws,r,"Jul Seasonality Factor",0.95,0.95,0.90,PCT,"Q3 typically softer; Risk scenario assumes demand drag"); r+=1
drow(ws,r,"Aug Seasonality Factor",1.00,1.05,0.95,PCT,"Mid-quarter normalization"); r+=1
drow(ws,r,"Sep Seasonality Factor",1.08,1.20,1.00,PCT,"End-of-quarter uplift; Upside assumes strong close"); r+=1
drow(ws,r,"Product Revenue Mix (%)",0.415,0.415,0.40,PCT,"Source: May26 actuals"); r+=1
drow(ws,r,"Service Revenue Mix (%)",0.385,0.385,0.375,PCT,"Source: May26 actuals"); r+=1
drow(ws,r,"Subscription Revenue Mix (%)",0.20,0.20,0.225,PCT,"Subscription defensive in Risk scenario"); r+=1

r = sec(ws, r, 6, "COST DRIVERS — FIXED", MID)
drow(ws,r,"Workforce Cost as % Revenue",0.084,0.080,0.090,PCT,"Blended BU avg: NSW 9.5%, QLD 6.9%, SA 9.2%, VIC 8.2%, WA 8.5%"); r+=1
drow(ws,r,"Prepayment Amortization (monthly $)",431000,431000,431000,DOLR,"Non-cash: Google Ads $500K, Deloitte $300K, MSFT $260K"); r+=1
drow(ws,r,"Depreciation (monthly $)",650000,650000,650000,DOLR,"Source: May26 actual D&A"); r+=1
drow(ws,r,"Insurance % Revenue",0.022,0.022,0.028,PCT,"Risk: +12% uplift from geopolitical premium"); r+=1
drow(ws,r,"Software/Licences (monthly $)",380000,380000,400000,DOLR,"Risk: vendor price increases"); r+=1

r = sec(ws, r, 6, "COST DRIVERS — VARIABLE", MID)
drow(ws,r,"Postage & Freight % Revenue",0.048,0.048,0.062,PCT,"Risk: +15% fuel/logistics surcharge (Middle East)"); r+=1
drow(ws,r,"Telecom & Utilities % Revenue",0.019,0.019,0.023,PCT,"Risk: +10% energy cost uplift"); r+=1
drow(ws,r,"Advertising % Revenue",0.036,0.040,0.034,PCT,"Upside: invest in growth; Risk: scale back"); r+=1
drow(ws,r,"Contractor/Consulting % Revenue",0.028,0.030,0.025,PCT,"Upside: digital transformation investment"); r+=1
drow(ws,r,"Supplier Pricing Uplift (geopolitical)",0.00,0.00,0.035,PCT,"Risk-only: 3.5% across variable supply costs"); r+=1
drow(ws,r,"Other OpEx % Revenue",0.055,0.052,0.060,PCT,"Catch-all variable costs"); r+=1

r = sec(ws, r, 6, "WORKING CAPITAL DRIVERS", "7030A0")
drow(ws,r,"Days Sales Outstanding (DSO)",45,42,55,NUM,"Base: May26; Risk: collections slow with overdue AR $3.33M"); r+=1
drow(ws,r,"AR Collection Rate — Current (<30d)",0.95,0.97,0.88,PCT,"Current AR $5.92M; Risk: 7% slowdown"); r+=1
drow(ws,r,"AR Collection Rate — 30-60d",0.75,0.80,0.60,PCT,"Amcor, South32 — Risk: elongated terms"); r+=1
drow(ws,r,"AR Collection Rate — 60-90d",0.55,0.60,0.40,PCT,"Newcrest, BlueScope — elevated credit risk"); r+=1
drow(ws,r,"AR Collection Rate — >90d",0.30,0.35,0.15,PCT,"Incitec, Orica — high impairment risk"); r+=1
drow(ws,r,"Days Payable Outstanding (DPO)",38,38,42,NUM,"Extend payables in Risk to conserve cash"); r+=1
drow(ws,r,"Monthly Capex ($)",400000,600000,300000,DOLR,"Upside: growth investment; Risk: defer spend"); r+=1

r = sec(ws, r, 6, "TAX & INTEREST", DKGRAY)
drow(ws,r,"Effective Tax Rate",0.30,0.30,0.30,PCT,"30% corporate rate"); r+=1
drow(ws,r,"Interest Income % Cash",0.0045,0.0045,0.0035,PCT,"Monthly rate on cash balances"); r+=1
drow(ws,r,"FX Impact on Revenue",0.008,0.005,0.015,PCT,"Risk: AUD weakening vs USD; Middle East disruption"); r+=1

print("Drivers sheet done, last row:", r)

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 3 – REVENUE MODEL
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_rev
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col in ['B','C','D','E','F','G','H']:
    ws.column_dimensions[col].width = 16

hdr(ws, 1, 1, "REVENUE MODEL — DRIVER-BASED MONTHLY BUILD", bg=NAVY, merge=(1,8))
ws.row_dimensions[1].height = 28

# Sub-header
for col, txt, bg in [(1,"Revenue Driver",NAVY),(2,"May 2026 Actual",DKGRAY),
                      (3,"Jul 2026",MID),(4,"Aug 2026",MID),(5,"Sep 2026",MID),
                      (6,"Q3 Total",NAVY),(7,"vs May (Δ$)",DKGRAY),(8,"vs May (Δ%)",DKGRAY)]:
    c = ws.cell(row=2, column=col, value=txt)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left", vertical="center", indent=1 if col==1 else 0)

# All driver refs point to Drivers!E column (Active Value)
# Row map in Drivers sheet:
# E7 = Baseline revenue  E8 = Growth rate  E9=Jul seas  E10=Aug seas  E11=Sep seas
# E12=Product mix  E13=Service mix  E14=Sub mix

DR = "Drivers!E"  # shorthand

def rev_formula(month_seas_row):
    # Revenue = Baseline × (1+growth) × seasonality
    return f"={DR}7*(1+{DR}8)*{DR}{month_seas_row}"

r = 3
sec(ws, r, 8, "TOTAL REVENUE BUILD", NAVY); r+=1

lbl(ws, r, 1, "Revenue Baseline (May 2026 actual)", bold=True); frm(ws,r,2,f"={DR}7",DOLR); 
frm(ws,r,3,rev_formula(9),DOLR); frm(ws,r,4,rev_formula(10),DOLR); frm(ws,r,5,rev_formula(11),DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,bold=True); frm(ws,r,7,f"=C{r}-B{r}",DOLR); frm(ws,r,8,f"=C{r}/B{r}-1",PCT)
rev_row = r; r+=1

sec(ws, r, 8, "BY PRODUCT TYPE", MID); r+=1

# Product
lbl(ws, r, 1, "  Product Revenue", indent=1)
frm(ws,r,2,f"={DR}7*{DR}12",DOLR)
frm(ws,r,3,f"=C{rev_row}*{DR}12",DOLR); frm(ws,r,4,f"=D{rev_row}*{DR}12",DOLR); frm(ws,r,5,f"=E{rev_row}*{DR}12",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); frm(ws,r,7,f"=C{r}-B{r}",DOLR); frm(ws,r,8,f"=C{r}/B{r}-1",PCT)
prod_row=r; r+=1

# Service
lbl(ws, r, 1, "  Service Revenue", indent=1)
frm(ws,r,2,f"={DR}7*{DR}13",DOLR)
frm(ws,r,3,f"=C{rev_row}*{DR}13",DOLR); frm(ws,r,4,f"=D{rev_row}*{DR}13",DOLR); frm(ws,r,5,f"=E{rev_row}*{DR}13",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); frm(ws,r,7,f"=C{r}-B{r}",DOLR); frm(ws,r,8,f"=C{r}/B{r}-1",PCT)
svc_row=r; r+=1

# Subscription
lbl(ws, r, 1, "  Subscription Revenue", indent=1)
frm(ws,r,2,f"={DR}7*{DR}14",DOLR)
frm(ws,r,3,f"=C{rev_row}*{DR}14",DOLR); frm(ws,r,4,f"=D{rev_row}*{DR}14",DOLR); frm(ws,r,5,f"=E{rev_row}*{DR}14",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); frm(ws,r,7,f"=C{r}-B{r}",DOLR); frm(ws,r,8,f"=C{r}/B{r}-1",PCT)
sub_row=r; r+=1

# Check row
lbl(ws, r, 1, "  Revenue Mix Check (should = 100%)", indent=1)
ws.cell(row=r, column=1).font = Font(name="Arial", size=9, italic=True, color="808080")
frm(ws,r,3,f"=C{prod_row}/C{rev_row}+C{svc_row}/C{rev_row}+C{sub_row}/C{rev_row}",PCT,color="808080"); r+=1

# BY BUSINESS UNIT
r+=1
sec(ws, r, 8, "BY BUSINESS UNIT (indicative allocation)", MID); r+=1
bu_splits = [("NSW (largest BU)",0.31),("VIC",0.22),("QLD",0.18),("WA",0.17),("SA",0.12)]
for bu, split in bu_splits:
    lbl(ws, r, 1, f"  {bu}", indent=1)
    frm(ws,r,2,f"={DR}7*{split}",DOLR)
    frm(ws,r,3,f"=C{rev_row}*{split}",DOLR); frm(ws,r,4,f"=D{rev_row}*{split}",DOLR); frm(ws,r,5,f"=E{rev_row}*{split}",DOLR)
    frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); frm(ws,r,7,f"=C{r}-B{r}",DOLR); frm(ws,r,8,f"=C{r}/B{r}-1",PCT); r+=1

lbl(ws, r, 1, "TOTAL BU REVENUE CHECK", bold=True)
frm(ws,r,3,f"=SUM(C{r-5}:C{r-1})",DOLR,bold=True); frm(ws,r,4,f"=SUM(D{r-5}:D{r-1})",DOLR,bold=True)
frm(ws,r,5,f"=SUM(E{r-5}:E{r-1})",DOLR,bold=True); frm(ws,r,6,f"=SUM(F{r-5}:F{r-1})",DOLR,bold=True); r+=1

# Seasonality notes
r+=1
sec(ws, r, 8, "SEASONALITY & GROWTH NOTES", DKGRAY); r+=1
notes = [
    "Jul factor (Active):", f"={DR}9",
    "Aug factor (Active):", f"={DR}10",
    "Sep factor (Active):", f"={DR}11",
    "Growth Rate (Active):", f"={DR}8",
]
for i in range(0, len(notes), 2):
    ws.cell(row=r, column=1, value=notes[i]).font = Font(name="Arial", size=9, color=DKGRAY)
    c = ws.cell(row=r, column=2, value=notes[i+1])
    c.font = Font(name="Arial", color=BGREEN, size=10)
    c.number_format = PCT if "factor" in notes[i].lower() or "Growth" in notes[i] else NUM
    r += 1

rev_model_rev_row = rev_row  # expose to other sheets
print("Revenue Model done, rev_row =", rev_row)

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 4 – COST MODEL
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_cst
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col in ['B','C','D','E','F','G','H']:
    ws.column_dimensions[col].width = 16

hdr(ws, 1, 1, "COST MODEL — FIXED / VARIABLE DECOMPOSITION", bg=NAVY, merge=(1,8))
ws.row_dimensions[1].height = 28

for col, txt, bg in [(1,"Cost Line",NAVY),(2,"May 2026 Actual",DKGRAY),
                      (3,"Jul 2026",MID),(4,"Aug 2026",MID),(5,"Sep 2026",MID),
                      (6,"Q3 Total","ED7D31"),(7,"vs May (Δ$)",DKGRAY),(8,"Type",DKGRAY)]:
    c = ws.cell(row=2, column=col, value=txt)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left", vertical="center")

# Revenue references from Revenue Model
RM = "'Revenue Model'!"
# Rev row 4 in Revenue Model = rev_row = 4
RR = f"{RM}C{rev_model_rev_row}"   # Jul revenue
RR_D = f"{RM}D{rev_model_rev_row}" # Aug
RR_E = f"{IM}E{rev_model_rev_row}" if False else f"{RM}E{rev_model_rev_row}" # Sep
RR_B = f"{RM}B{rev_model_rev_row}" # May actual

# Shorthand: Drivers E column
DR = "Drivers!E"

def cost_row_pct(ws, r, label, may_pct, drv_row, ctype="Variable"):
    """Variable cost = Revenue × driver %"""
    lbl(ws, r, 1, label, indent=1)
    frm(ws,r,2,f"={RR_B}*{may_pct}",DOLR)
    frm(ws,r,3,f"={RR}*{DR}{drv_row}",DOLR)
    frm(ws,r,4,f"={RR_D}*{DR}{drv_row}",DOLR)
    frm(ws,r,5,f"={RR_E}*{DR}{drv_row}",DOLR)
    frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR)
    frm(ws,r,7,f"=C{r}-B{r}",DOLR)
    ws.cell(row=r, column=8, value=ctype).font = Font(name="Arial", size=9, color="808080")

def cost_row_fixed(ws, r, label, drv_row):
    """Fixed cost = directly from driver cell"""
    lbl(ws, r, 1, label, indent=1)
    frm(ws,r,2,f"={DR}{drv_row}",DOLR)
    frm(ws,r,3,f"={DR}{drv_row}",DOLR)
    frm(ws,r,4,f"={DR}{drv_row}",DOLR)
    frm(ws,r,5,f"={DR}{drv_row}",DOLR)
    frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR)
    frm(ws,r,7,f"=C{r}-B{r}",DOLR)
    ws.cell(row=r, column=8, value="Fixed").font = Font(name="Arial", size=9, color="808080")

r = 3
# Drivers row mapping (E column):
# E16=Workforce%  E17=Prepay amort  E18=Depreciation  E19=Insurance%  E20=Software$
# E22=Freight%    E23=Telecom%      E24=Advertising%  E25=Contractor%  E26=Supplier uplift  E27=Other OpEx%

sec(ws, r, 8, "FIXED COSTS", MID); r+=1
cost_row_pct(ws,r,"Workforce (salaries & on-costs)",0.084,16,"Fixed"); r+=1
cost_row_fixed(ws,r,"Prepayment Amortization (non-cash)",17); r+=1
cost_row_fixed(ws,r,"Depreciation & Amortisation (non-cash)",18); r+=1
cost_row_pct(ws,r,"Insurance",0.022,19,"Fixed"); r+=1
cost_row_fixed(ws,r,"Software & Licences",20); r+=1

lbl(ws, r, 1, "TOTAL FIXED COSTS", bold=True)
for col in [2,3,4,5,6]:
    frm(ws,r,col,f"=SUM({get_column_letter(col)}{r-5}:{get_column_letter(col)}{r-1})",DOLR,bold=True)
fixed_tot_row = r; r+=1

sec(ws, r, 8, "VARIABLE COSTS", "ED7D31"); r+=1
cost_row_pct(ws,r,"Postage, Freight & Logistics",0.048,22); r+=1
cost_row_pct(ws,r,"Telecom & Utilities",0.019,23); r+=1
cost_row_pct(ws,r,"Advertising & Marketing",0.036,24); r+=1
cost_row_pct(ws,r,"Contractor & Consulting",0.028,25); r+=1
cost_row_pct(ws,r,"Geopolitical Supplier Price Uplift",0.000,26); r+=1
cost_row_pct(ws,r,"Other Operating Expenses",0.055,27); r+=1

lbl(ws, r, 1, "TOTAL VARIABLE COSTS", bold=True)
for col in [2,3,4,5,6]:
    frm(ws,r,col,f"=SUM({get_column_letter(col)}{r-6}:{get_column_letter(col)}{r-1})",DOLR,bold=True)
var_tot_row = r; r+=1

lbl(ws, r, 1, "TOTAL OPERATING COSTS", bold=True)
for col in [2,3,4,5,6]:
    frm(ws,r,col,f"={get_column_letter(col)}{fixed_tot_row}+{get_column_letter(col)}{var_tot_row}",DOLR,bold=True)
    ws.cell(row=r, column=col).fill = PatternFill("solid", start_color=LIGHT)
tot_cost_row = r; r+=1

# Non-cash costs note
r+=1
sec(ws, r, 8, "NON-CASH COST NOTE (Accounting vs Cash)", "7030A0"); r+=1
lbl(ws, r, 1, "  Prepayment Amortization (non-cash expensed)")
frm(ws,r,3,f"={DR}17",DOLR,color="7030A0"); frm(ws,r,4,f"={DR}17",DOLR,color="7030A0"); frm(ws,r,5,f"={DR}17",DOLR,color="7030A0")
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,color="7030A0"); r+=1
lbl(ws, r, 1, "  Depreciation (non-cash expensed)")
frm(ws,r,3,f"={DR}18",DOLR,color="7030A0"); frm(ws,r,4,f"={DR}18",DOLR,color="7030A0"); frm(ws,r,5,f"={DR}18",DOLR,color="7030A0")
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,color="7030A0"); r+=1
lbl(ws, r, 1, "TOTAL NON-CASH (addback for cash flow)", bold=True)
frm(ws,r,3,f"=C{r-2}+C{r-1}",DOLR,bold=True,color="7030A0"); frm(ws,r,4,f"=D{r-2}+D{r-1}",DOLR,bold=True,color="7030A0")
frm(ws,r,5,f"=E{r-2}+E{r-1}",DOLR,bold=True,color="7030A0"); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,bold=True,color="7030A0")
non_cash_row = r; r+=1

print(f"Cost Model done. fixed_tot={fixed_tot_row}, var_tot={var_tot_row}, tot_cost={tot_cost_row}, non_cash={non_cash_row}")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 5 – P&L MONTHLY
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_pl
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col in ['B','C','D','E','F','G']:
    ws.column_dimensions[col].width = 17

hdr(ws, 1, 1, "P&L MONTHLY FORECAST — ACCOUNTING PERFORMANCE", bg=NAVY, merge=(1,7))
ws.row_dimensions[1].height = 28

for col, txt, bg in [(1,"Line Item",NAVY),(2,"May 2026 Actual",DKGRAY),
                      (3,"Jul 2026",MID),(4,"Aug 2026",MID),(5,"Sep 2026",MID),
                      (6,"Q3 2026 Total",NAVY),(7,"Q3 Margin %",NAVY)]:
    c = ws.cell(row=2, column=col, value=txt)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left", vertical="center")

RM = "'Revenue Model'!"
CM = "'Cost Model'!"
DR = "Drivers!E"

RR = rev_model_rev_row  # Revenue row in Revenue Model

r = 3
def pl_row(ws, r, label, b_val, jul_f, aug_f, sep_f, fmt=DOLR, bold=False, bg=None, margin=False):
    lbl(ws, r, 1, label, bold=bold, bg=bg if bg else None)
    if bg:
        for col in range(1,8):
            ws.cell(row=r, column=col).fill = PatternFill("solid", start_color=bg)
    frm(ws,r,2,b_val,fmt,bold)
    frm(ws,r,3,jul_f,fmt,bold)
    frm(ws,r,4,aug_f,fmt,bold)
    frm(ws,r,5,sep_f,fmt,bold)
    frm(ws,r,6,f"=SUM(C{r}:E{r})",fmt,bold)
    if margin:
        frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,bold) if r != 3 else frm(ws,r,7,"",PCT)
    return r

# REVENUE
lbl(ws, r, 1, "REVENUE", bold=True, bg=NAVY); 
for c in range(1,8): ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=NAVY)
for c in range(1,8): ws.cell(row=r, column=c).font = Font(name="Arial", bold=True, color=WHITE, size=10)
r+=1

pl_row(ws,r,"Total Revenue",
       f"={RM}B{RR}",f"={RM}C{RR}",f"={RM}D{RR}",f"={RM}E{RR}",DOLR,True,None,False)
ws.cell(row=r,column=7,value="100.0%").font=Font(name="Arial",bold=True,size=10)
pl_rev_row = r; r+=1

# GROSS PROFIT
r += 0
lbl(ws, r, 1, "COST OF SALES", bold=True, bg=LGRAY)
for c in range(1,8): ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=LGRAY)
r+=1

# Freight/logistics as COGS proxy
lbl(ws, r, 1, "  Postage, Freight & Logistics (COGS)", indent=1)
frm(ws,r,2,f"={CM}C{fixed_tot_row+2}",DOLR)  # freight is first variable cost row
frm(ws,r,3,f"={CM}C{fixed_tot_row+2}",DOLR); frm(ws,r,4,f"={CM}D{fixed_tot_row+2}",DOLR); frm(ws,r,5,f"={CM}E{fixed_tot_row+2}",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws, r, 1, "GROSS PROFIT", bold=True)
for c in range(1,8): ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=GREEN)
frm(ws,r,2,f"={RM}B{RR}-'Cost Model'!B{fixed_tot_row+2}",DOLR,True)
frm(ws,r,3,f"={RM}C{RR}-'Cost Model'!C{fixed_tot_row+2}",DOLR,True)
frm(ws,r,4,f"={RM}D{RR}-'Cost Model'!D{fixed_tot_row+2}",DOLR,True)
frm(ws,r,5,f"={RM}E{RR}-'Cost Model'!E{fixed_tot_row+2}",DOLR,True)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True)
frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,True)
gp_row = r; r+=1

# OPEX
lbl(ws, r, 1, "OPERATING EXPENSES", bold=True, bg=LGRAY)
for c in range(1,8): ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=LGRAY)
r+=1

opex_start = r
opex_items = [
    ("  Workforce (salaries & on-costs)", fixed_tot_row-4, CM),   # row offsets in Cost Model
    ("  Advertising & Marketing", fixed_tot_row+4, CM),
    ("  Contractor & Consulting", fixed_tot_row+5, CM),
    ("  Telecom & Utilities", fixed_tot_row+3, CM),
    ("  Software & Licences", fixed_tot_row+0, CM),   # last fixed cost row
    ("  Insurance", fixed_tot_row-1, CM),
    ("  Other Operating Expenses", fixed_tot_row+7, CM),
    ("  Geopolitical Supplier Uplift", fixed_tot_row+6, CM),
]

for label, cst_row, sheet in opex_items:
    lbl(ws, r, 1, label, indent=1)
    for col_letter, src_col in [("B","B"),("C","C"),("D","D"),("E","E")]:
        frm(ws,r,{"B":2,"C":3,"D":4,"E":5}[col_letter],f"={sheet}{src_col}{cst_row}",DOLR)
    frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

# Prepay amort + D&A (non-cash in P&L)
lbl(ws,r,1,"  Prepayment Amortization (non-cash)",indent=1)
frm(ws,r,2,f"={DR}17",DOLR); frm(ws,r,3,f"={DR}17",DOLR); frm(ws,r,4,f"={DR}17",DOLR); frm(ws,r,5,f"={DR}17",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"  Depreciation & Amortisation (non-cash)",indent=1)
frm(ws,r,2,f"={DR}18",DOLR); frm(ws,r,3,f"={DR}18",DOLR); frm(ws,r,4,f"={DR}18",DOLR); frm(ws,r,5,f"={DR}18",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

opex_end = r - 1
lbl(ws, r, 1, "TOTAL OPERATING EXPENSES", bold=True)
for c in range(1,8): ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=LGRAY)
for col in [2,3,4,5]:
    frm(ws,r,col,f"=SUM({get_column_letter(col)}{opex_start}:{get_column_letter(col)}{opex_end})",DOLR,True)
frm(ws,r,6,f"=SUM(F{opex_start}:F{opex_end})",DOLR,True)
frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,True)
tot_opex_row = r; r+=1

# EBITDA
lbl(ws, r, 1, "EBITDA", bold=True)
for c in range(1,8): ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=MID)
for c in range(1,8): ws.cell(row=r, column=c).font = Font(name="Arial",bold=True,color=WHITE,size=10)
frm(ws,r,2,f"={RM}B{RR}-'Cost Model'!B{tot_cost_row}+'Cost Model'!B{non_cash_row}",DOLR,True,color=WHITE)
frm(ws,r,3,f"={RM}C{RR}-'Cost Model'!C{tot_cost_row}+'Cost Model'!C{non_cash_row}",DOLR,True,color=WHITE)
frm(ws,r,4,f"={RM}D{RR}-'Cost Model'!D{tot_cost_row}+'Cost Model'!D{non_cash_row}",DOLR,True,color=WHITE)
frm(ws,r,5,f"={RM}E{RR}-'Cost Model'!E{tot_cost_row}+'Cost Model'!E{non_cash_row}",DOLR,True,color=WHITE)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color=WHITE)
frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,True,color=WHITE)
ebitda_row = r; r+=1

# EBIT
lbl(ws,r,1,"EBIT (after D&A)",bold=True)
frm(ws,r,2,f"=B{ebitda_row}-{DR}18-{DR}17",DOLR,True)
frm(ws,r,3,f"=C{ebitda_row}-{DR}18-{DR}17",DOLR,True)
frm(ws,r,4,f"=D{ebitda_row}-{DR}18-{DR}17",DOLR,True)
frm(ws,r,5,f"=E{ebitda_row}-{DR}18-{DR}17",DOLR,True)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True)
frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,True)
ebit_row = r; r+=1

# Interest income
lbl(ws,r,1,"  Interest Income",indent=1)
frm(ws,r,2,f"=12911223*{DR}30",DOLR)  # E30=interest rate
frm(ws,r,3,f"=12911223*{DR}30",DOLR); frm(ws,r,4,f"=12911223*{DR}30",DOLR); frm(ws,r,5,f"=12911223*{DR}30",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1
int_row = r-1

# PBT
lbl(ws,r,1,"PROFIT BEFORE TAX",bold=True)
for c in range(1,8): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=GREEN)
frm(ws,r,2,f"=B{ebit_row}+B{int_row}",DOLR,True)
frm(ws,r,3,f"=C{ebit_row}+C{int_row}",DOLR,True)
frm(ws,r,4,f"=D{ebit_row}+D{int_row}",DOLR,True)
frm(ws,r,5,f"=E{ebit_row}+E{int_row}",DOLR,True)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True)
frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,True)
pbt_row = r; r+=1

# Tax
lbl(ws,r,1,"  Income Tax Expense",indent=1)
frm(ws,r,2,f"=B{pbt_row}*{DR}29",DOLR)
frm(ws,r,3,f"=C{pbt_row}*{DR}29",DOLR); frm(ws,r,4,f"=D{pbt_row}*{DR}29",DOLR); frm(ws,r,5,f"=E{pbt_row}*{DR}29",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1
tax_row = r-1

# Net Income
lbl(ws,r,1,"NET INCOME (ACCOUNTING)",bold=True)
for c in range(1,8): 
    ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=NAVY)
    ws.cell(row=r,column=c).font=Font(name="Arial",bold=True,color=WHITE,size=11)
frm(ws,r,2,f"=B{pbt_row}-B{tax_row}",DOLR,True,color=WHITE)
frm(ws,r,3,f"=C{pbt_row}-C{tax_row}",DOLR,True,color=WHITE)
frm(ws,r,4,f"=D{pbt_row}-D{tax_row}",DOLR,True,color=WHITE)
frm(ws,r,5,f"=E{pbt_row}-E{tax_row}",DOLR,True,color=WHITE)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color=WHITE)
frm(ws,r,7,f"=F{r}/F{pl_rev_row}",PCT,True,color=WHITE)
ni_row = r; r+=1

print(f"P&L Monthly done. rev={pl_rev_row}, ebitda={ebitda_row}, ni={ni_row}")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 6 – CASH FLOW (Accounting vs Cash Realisation)
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_cf
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 42
for col in ['B','C','D','E','F']:
    ws.column_dimensions[col].width = 17

hdr(ws, 1, 1, "CASH FLOW — ACCOUNTING INCOME vs CASH REALISATION", bg="7030A0", merge=(1,6))
ws.row_dimensions[1].height = 30

for col, txt, bg in [(1,"Line Item","7030A0"),(2,"May 2026",DKGRAY),
                      (3,"Jul 2026",MID),(4,"Aug 2026",MID),(5,"Sep 2026",MID),(6,"Q3 Total","7030A0")]:
    c = ws.cell(row=2, column=col, value=txt)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left", vertical="center")

DR = "Drivers!E"
PL = "'P&L Monthly'!"
CM = "'Cost Model'!"
WC = "'Working Capital'!"

r = 3
sec(ws, r, 6, "OPERATING CASH FLOW (Indirect Method)", "7030A0"); r+=1

# Start with Net Income
lbl(ws,r,1,"Net Income (Accounting)",bold=True)
frm(ws,r,2,f"={PL}B{ni_row}",DOLR,True,color=BGREEN)
frm(ws,r,3,f"={PL}C{ni_row}",DOLR,True,color=BGREEN)
frm(ws,r,4,f"={PL}D{ni_row}",DOLR,True,color=BGREEN)
frm(ws,r,5,f"={PL}E{ni_row}",DOLR,True,color=BGREEN)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color=BGREEN)
ni_cf_row = r; r+=1

sec(ws, r, 6, "ADD BACK: Non-Cash Items", MID); r+=1
lbl(ws,r,1,"  + Depreciation & Amortisation",indent=1)
frm(ws,r,2,f"={DR}18",DOLR); frm(ws,r,3,f"={DR}18",DOLR); frm(ws,r,4,f"={DR}18",DOLR); frm(ws,r,5,f"={DR}18",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"  + Prepayment Amortization (cash paid earlier)",indent=1)
frm(ws,r,2,f"={DR}17",DOLR); frm(ws,r,3,f"={DR}17",DOLR); frm(ws,r,4,f"={DR}17",DOLR); frm(ws,r,5,f"={DR}17",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR)
ws.cell(row=r,column=1).fill=PatternFill("solid",start_color=AMBER)  # highlight non-cash distinction
r+=1
noncash_end = r-1

sec(ws, r, 6, "WORKING CAPITAL MOVEMENTS (Cash Impact)", "00B0F0"); r+=1

# AR collection gap — the KEY cash vs accounting distinction
lbl(ws,r,1,"  (Increase)/Decrease in Trade Receivables",indent=1)
ws.cell(row=r,column=1).fill=PatternFill("solid",start_color=AMBER)
# AR increases as revenue grows — cash lower than acctg profit
# Simplified: revenue growth × (1 - collection rate blended)
# Blended collection rate from WC sheet row
frm(ws,r,2,"=-500000",DOLR)  # May baseline AR movement
frm(ws,r,3,f"=-('Revenue Model'!C{rev_model_rev_row}*(1-{DR}22))",DOLR)  # E22=AR collection current
frm(ws,r,4,f"=-('Revenue Model'!D{rev_model_rev_row}*(1-{DR}22))",DOLR)
frm(ws,r,5,f"=-('Revenue Model'!E{rev_model_rev_row}*(1-{DR}22))",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR)
ar_move_row = r; r+=1

lbl(ws,r,1,"  (Increase)/Decrease in Prepayments",indent=1)
frm(ws,r,2,"=0",DOLR); frm(ws,r,3,"=0",DOLR); frm(ws,r,4,"=0",DOLR); frm(ws,r,5,"=0",DOLR)
frm(ws,r,6,"=0",DOLR); r+=1

lbl(ws,r,1,"  Increase/(Decrease) in Trade Payables",indent=1)
# Extending DPO in Risk scenario frees up cash
frm(ws,r,2,"=180000",DOLR)
frm(ws,r,3,f"='Cost Model'!C{tot_cost_row}/30*({DR}26-38)",DOLR)  # E26=DPO active, 38=base May DPO
frm(ws,r,4,f"='Cost Model'!D{tot_cost_row}/30*({DR}26-38)",DOLR)
frm(ws,r,5,f"='Cost Model'!E{tot_cost_row}/30*({DR}26-38)",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR)
wc_pay_row = r; r+=1

lbl(ws,r,1,"TOTAL WORKING CAPITAL MOVEMENT",bold=True)
for c in range(2,7):
    frm(ws,r,c,f"=SUM({get_column_letter(c)}{ar_move_row}:{get_column_letter(c)}{wc_pay_row})",DOLR,True)
wc_tot_row = r; r+=1

# Operating Cash Flow
lbl(ws,r,1,"OPERATING CASH FLOW",bold=True)
for c in range(1,7): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color="7030A0")
for c in range(1,7): ws.cell(row=r,column=c).font=Font(name="Arial",bold=True,color=WHITE,size=11)
frm(ws,r,2,f"=B{ni_cf_row}+B{ni_cf_row+1}+B{noncash_end}+B{wc_tot_row}",DOLR,True,color=WHITE)
frm(ws,r,3,f"=C{ni_cf_row}+C{ni_cf_row+1}+C{noncash_end}+C{wc_tot_row}",DOLR,True,color=WHITE)
frm(ws,r,4,f"=D{ni_cf_row}+D{ni_cf_row+1}+D{noncash_end}+D{wc_tot_row}",DOLR,True,color=WHITE)
frm(ws,r,5,f"=E{ni_cf_row}+E{ni_cf_row+1}+E{noncash_end}+E{wc_tot_row}",DOLR,True,color=WHITE)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color=WHITE)
ocf_row = r; r+=1

sec(ws, r, 6, "INVESTING CASH FLOW", DKGRAY); r+=1
lbl(ws,r,1,"  Capital Expenditure",indent=1)
frm(ws,r,2,"-400000",DOLR); frm(ws,r,3,f"=-{DR}28",DOLR); frm(ws,r,4,f"=-{DR}28",DOLR); frm(ws,r,5,f"=-{DR}28",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); capex_row=r; r+=1

lbl(ws,r,1,"INVESTING CASH FLOW",bold=True)
for c in range(2,7): frm(ws,r,c,f"=SUM({get_column_letter(c)}{capex_row}:{get_column_letter(c)}{capex_row})",DOLR,True)
icf_row = r; r+=1

sec(ws, r, 6, "NET CASH POSITION", NAVY); r+=1
lbl(ws,r,1,"Opening Cash Balance",bold=True)
frm(ws,r,2,"=12911223",DOLR,True)
frm(ws,r,3,f"=B{r}",DOLR,True,color=BGREEN)  # Jul opens = May close
frm(ws,r,4,f"=E{r-1}",DOLR,True,color=BGREEN)  # Will be set iteratively below
frm(ws,r,5,f"=E{r-1}",DOLR,True,color=BGREEN)
frm(ws,r,6,"",DOLR)
open_row = r; r+=1

# Net movement
lbl(ws,r,1,"Net Cash Movement",bold=True)
for c in range(3,7):
    frm(ws,r,c,f"={get_column_letter(c)}{ocf_row}+{get_column_letter(c)}{icf_row}",DOLR,True)
net_move_row = r; r+=1

# Closing balance
lbl(ws,r,1,"CLOSING CASH BALANCE",bold=True)
for c in range(1,7): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=NAVY)
for c in range(1,7): ws.cell(row=r,column=c).font=Font(name="Arial",bold=True,color=WHITE,size=11)
frm(ws,r,3,f"=C{open_row}+C{net_move_row}",DOLR,True,color=WHITE)
frm(ws,r,4,f"=C{r}+D{net_move_row}",DOLR,True,color=WHITE)
frm(ws,r,5,f"=D{r}+E{net_move_row}",DOLR,True,color=WHITE)
frm(ws,r,6,f"=E{r}",DOLR,True,color=WHITE)
ws.cell(row=r,column=2,value="=12911223").font=Font(name="Arial",bold=True,color=WHITE,size=11)
ws.cell(row=r,column=2).number_format=DOLR
close_row = r; r+=1

# Fix the open_row Aug and Sep references to point to prior month close
ws.cell(row=open_row,column=4,value=f"=C{close_row}").font=Font(name="Arial",color=BGREEN,size=10,bold=True)
ws.cell(row=open_row,column=4).number_format=DOLR
ws.cell(row=open_row,column=4).fill=PatternFill("solid",start_color="FFFFFF")
ws.cell(row=open_row,column=5,value=f"=D{close_row}").font=Font(name="Arial",color=BGREEN,size=10,bold=True)
ws.cell(row=open_row,column=5).number_format=DOLR
ws.cell(row=open_row,column=5).fill=PatternFill("solid",start_color="FFFFFF")

# ACCOUNTING vs CASH RECONCILIATION NOTE
r+=1
sec(ws, r, 6, "ACCOUNTING INCOME vs CASH REALISATION RECONCILIATION", AMBER); r+=1
lbl(ws,r,1,"Net Income (Accounting)",bold=True)
frm(ws,r,3,f"={PL}C{ni_row}",DOLR,True,color=BGREEN); frm(ws,r,4,f"={PL}D{ni_row}",DOLR,True,color=BGREEN)
frm(ws,r,5,f"={PL}E{ni_row}",DOLR,True,color=BGREEN); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color=BGREEN); r+=1

lbl(ws,r,1,"Less: Uncollected Revenue (AR gap)",indent=1)
frm(ws,r,3,f"=C{ar_move_row}",DOLR); frm(ws,r,4,f"=D{ar_move_row}",DOLR); frm(ws,r,5,f"=E{ar_move_row}",DOLR)
frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"Add: Non-Cash Addbacks (D&A + Prepay amort)",indent=1)
frm(ws,r,3,f"=C{ni_cf_row+1}+C{ni_cf_row+2}",DOLR); frm(ws,r,4,f"=D{ni_cf_row+1}+D{ni_cf_row+2}",DOLR)
frm(ws,r,5,f"=E{ni_cf_row+1}+E{ni_cf_row+2}",DOLR); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"CASH REALISED (Operating)",bold=True)
for c in range(1,7): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=GREEN)
frm(ws,r,3,f"=C{ocf_row}",DOLR,True,color="006400"); frm(ws,r,4,f"=D{ocf_row}",DOLR,True,color="006400")
frm(ws,r,5,f"=E{ocf_row}",DOLR,True,color="006400"); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color="006400"); r+=1

lbl(ws,r,1,"Gap (Accounting > Cash) — liquidity risk",bold=True)
frm(ws,r,3,f"=C{r-4}-C{r-1}",DOLR,True,color="C00000"); frm(ws,r,4,f"=D{r-4}-D{r-1}",DOLR,True,color="C00000")
frm(ws,r,5,f"=E{r-4}-E{r-1}",DOLR,True,color="C00000"); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR,True,color="C00000")
for c in range(1,7): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=RED)
r+=1

print(f"Cash Flow done. ocf={ocf_row}, close={close_row}")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 7 – WORKING CAPITAL
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_wc
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 38
for col in ['B','C','D','E','F']:
    ws.column_dimensions[col].width = 17

hdr(ws, 1, 1, "WORKING CAPITAL — DSO / AR AGING / COLLECTION TIMING", bg="00B0F0", merge=(1,6))
ws.row_dimensions[1].height = 28

for col, txt, bg in [(1,"Metric","00B0F0"),(2,"May 2026 Actual",DKGRAY),
                      (3,"Jul 2026",MID),(4,"Aug 2026",MID),(5,"Sep 2026",MID),(6,"Q3 Avg / Total","00B0F0")]:
    c = ws.cell(row=2, column=col, value=txt)
    c.font = Font(name="Arial", bold=True, color=WHITE, size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left", vertical="center")

DR = "Drivers!E"
RM = "'Revenue Model'!"

r = 3
sec(ws, r, 6, "DSO & AR METRICS", "00B0F0"); r+=1

lbl(ws,r,1,"Days Sales Outstanding (DSO)")
inp(ws,r,2,45,NUM); frm(ws,r,3,f"={DR}21",NUM,color=BGREEN); frm(ws,r,4,f"={DR}21",NUM,color=BGREEN)
frm(ws,r,5,f"={DR}21",NUM,color=BGREEN); frm(ws,r,6,f"=AVERAGE(C{r}:E{r})",NUM); r+=1

lbl(ws,r,1,"Monthly Revenue")
frm(ws,r,2,f"={RM}B{rev_model_rev_row}",DOLR,color=BGREEN)
frm(ws,r,3,f"={RM}C{rev_model_rev_row}",DOLR,color=BGREEN); frm(ws,r,4,f"={RM}D{rev_model_rev_row}",DOLR,color=BGREEN)
frm(ws,r,5,f"={RM}E{rev_model_rev_row}",DOLR,color=BGREEN); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); rev_wc_row=r; r+=1

lbl(ws,r,1,"Implied AR Balance (Revenue × DSO / 30)")
frm(ws,r,2,f"=B{rev_wc_row}*B{r-2}/30",DOLR)
frm(ws,r,3,f"=C{rev_wc_row}*C{r-2}/30",DOLR); frm(ws,r,4,f"=D{rev_wc_row}*D{r-2}/30",DOLR)
frm(ws,r,5,f"=E{rev_wc_row}*E{r-2}/30",DOLR); frm(ws,r,6,f"=AVERAGE(C{r}:E{r})",DOLR); r+=1

sec(ws, r, 6, "AR AGING BUCKETS — MAY 2026 ACTUALS", DKGRAY); r+=1
aging = [
    ("Current (<30 days)", 5922900, 0.64),
    ("30-60 days overdue", 1245600, 0.135),
    ("60-90 days overdue", 1287800, 0.139),
    (">90 days overdue",   793400, 0.086),
]
aging_rows = []
for label, amt, pct in aging:
    lbl(ws,r,1,f"  {label}",indent=1)
    inp(ws,r,2,amt,DOLR)
    ws.cell(row=r,column=3,value=pct).number_format=PCT
    ws.cell(row=r,column=3).font=Font(name="Arial",size=10,color=DKGRAY)
    aging_rows.append(r); r+=1

lbl(ws,r,1,"Total AR Outstanding",bold=True)
frm(ws,r,2,f"=SUM(B{aging_rows[0]}:B{aging_rows[-1]})",DOLR,True); r+=1
ar_tot_row=r-1

sec(ws, r, 6, "COLLECTION CASH FLOW FORECAST", "00B0F0"); r+=1
lbl(ws,r,1,"  Current bucket collections (cash in)")
frm(ws,r,3,f"=B{aging_rows[0]}*{DR}22",DOLR); frm(ws,r,4,f"=B{aging_rows[0]}*{DR}22",DOLR)
frm(ws,r,5,f"=B{aging_rows[0]}*{DR}22",DOLR); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"  30-60d bucket collections (cash in)")
frm(ws,r,3,f"=B{aging_rows[1]}*{DR}23",DOLR); frm(ws,r,4,f"=B{aging_rows[1]}*{DR}23",DOLR)
frm(ws,r,5,f"=B{aging_rows[1]}*{DR}23",DOLR); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"  60-90d bucket collections (cash in)")
frm(ws,r,3,f"=B{aging_rows[2]}*{DR}24",DOLR); frm(ws,r,4,f"=B{aging_rows[2]}*{DR}24",DOLR)
frm(ws,r,5,f"=B{aging_rows[2]}*{DR}24",DOLR); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR); r+=1

lbl(ws,r,1,"  >90d collections (high impairment risk)")
frm(ws,r,3,f"=B{aging_rows[3]}*{DR}25",DOLR); frm(ws,r,4,f"=B{aging_rows[3]}*{DR}25",DOLR)
frm(ws,r,5,f"=B{aging_rows[3]}*{DR}25",DOLR); frm(ws,r,6,f"=SUM(C{r}:E{r})",DOLR)
coll_end=r; r+=1

lbl(ws,r,1,"TOTAL PROJECTED COLLECTIONS (CASH)",bold=True)
for c in range(3,7): frm(ws,r,c,f"=SUM({get_column_letter(c)}{coll_end-3}:{get_column_letter(c)}{coll_end})",DOLR,True)
tot_coll_row=r; r+=1

lbl(ws,r,1,"Overdue AR at Risk of Non-Recovery",bold=True)
frm(ws,r,2,"=3326800",DOLR,True,color="C00000")
frm(ws,r,3,f"=B{aging_rows[2]}*(1-{DR}24)+B{aging_rows[3]}*(1-{DR}25)",DOLR,True,color="C00000")
frm(ws,r,4,f"=C{r}",DOLR,True,color="C00000"); frm(ws,r,5,f"=D{r}",DOLR,True,color="C00000")
frm(ws,r,6,f"=C{r}",DOLR,True,color="C00000")
for c in range(1,7): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=RED)
r+=1

# Overdue customers list
r+=1
sec(ws, r, 6, "OVERDUE CUSTOMER WATCHLIST (May 2026)", "C00000"); r+=1
overdue_cust = [
    ("Amcor Ltd", "60-90d", 287600, "Packaging — freight cost sensitivity"),
    ("Newcrest Mining", "60-90d", 312400, "Mining — capital constraint"),
    ("South32", "30-60d", 198700, "Resources — FX exposure"),
    ("Incitec Pivot", ">90d", 245300, "Agriculture — seasonal cash cycle"),
    ("Orica Ltd", ">90d", 287100, "Mining chemicals — procurement freeze"),
    ("BlueScope Steel", "60-90d", 389600, "Steel — energy cost pressures"),
    ("a2 Milk", "30-60d", 312800, "FMCG — NZ/China exposure"),
    ("Auckland Airport", "30-60d", 493300, "Aviation — fuel cost sensitivity"),
]
for cust, bucket, amt, note in overdue_cust:
    lbl(ws,r,1,f"  {cust}",indent=1); 
    ws.cell(row=r,column=2,value=bucket).font=Font(name="Arial",size=10,color="C00000")
    inp(ws,r,3,amt,DOLR)
    ws.cell(row=r,column=4,value=note).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    r+=1

print(f"Working Capital done. tot_coll={tot_coll_row}")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 8 – BUDGET VARIANCE
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_bv
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 38
for col in ['B','C','D','E','F','G']:
    ws.column_dimensions[col].width = 16

hdr(ws,1,1,"BUDGET VARIANCE ANALYSIS — MAY 2026 ACTUALS vs BUDGET",bg="FFC000",merge=(1,7))
ws.row_dimensions[1].height = 28

for col,txt,bg in [(1,"Line Item","FFC000"),(2,"May Budget",DKGRAY),(3,"May Actual",DKGRAY),
                   (4,"Variance ($)","FFC000"),(5,"Variance (%)","FFC000"),
                   (6,"Q3 Budget (proj)",MID),(7,"Commentary",DKGRAY)]:
    c = ws.cell(row=2,column=col,value=txt)
    c.font = Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill = PatternFill("solid",start_color=bg)
    c.alignment = Alignment(horizontal="center" if col>1 else "left",vertical="center")

DR = "Drivers!E"
RM = "'Revenue Model'!"

bv_data = [
    # (label, budget, actual, commentary)
    ("REVENUE","","",""),
    ("  Product Revenue", 15478200, 16817976, "Fav: Volume uplift from NSW & VIC"),
    ("  Service Revenue", 14389600, 15610318, "Fav: New contract wins Q2"),
    ("  Subscription Revenue", 7863400, 8117882, "Fav: Renewal rate above target"),
    ("  FX Gains", 1124800, 1268900, "Fav: AUD weakness vs USD benefited"),
    ("TOTAL REVENUE", 38855000, 40546176, ""),
    ("","","",""),
    ("OPERATING EXPENSES","","",""),
    ("  Workforce Costs", 3456000, 3426019, "Fav: Vacancy savings partially offset"),
    ("  Postage & Freight", 2040000, 1946216, "Fav: Route optimisation savings"),
    ("  Insurance", 914000, 892015, "Fav: Policy renegotiation"),
    ("  Advertising & Marketing", 1521000, 1459663, "Fav: Digital campaign under-spend"),
    ("  Contractor & Consulting", 1243000, 1135293, "Fav: Project delays pushed to Q3"),
    ("  Telecom & Utilities", 788000, 770377, "Fav: Usage reduction"),
    ("  Software & Licences", 380000, 380000, "On target"),
    ("  Depreciation & Amort", 672000, 650000, "Fav: Asset timing"),
    ("  Prepay Amortization", 431000, 431000, "On target per schedule"),
    ("  Other OpEx", 2344000, 2230280, "Fav: General cost discipline"),
    ("TOTAL OPEX", 13789000, 13320863, ""),
    ("","","",""),
    ("NET INCOME", 9075000, 9075108, "Marginally on target"),
    ("EBITDA", 11320000, 11290713, "Within 0.3% of budget"),
]

r = 3
for row_data in bv_data:
    label, budget, actual, comment = row_data
    if label in ("REVENUE","OPERATING EXPENSES"):
        sec(ws,r,7,label,"FFC000"); r+=1; continue
    if label == "":
        r+=1; continue
    is_total = label.startswith("TOTAL") or label in ("NET INCOME","EBITDA")
    lbl(ws,r,1,label,bold=is_total,bg=GREEN if is_total else None)
    if budget:
        inp(ws,r,2,budget,DOLR)
        inp(ws,r,3,actual,DOLR)
        frm(ws,r,4,f"=C{r}-B{r}",DOLR,color="006400" if actual>budget else "C00000" if actual<budget else "000000")
        frm(ws,r,5,f"=IF(B{r}<>0,(C{r}-B{r})/ABS(B{r}),0)",PCT,color="006400" if actual>budget else "C00000")
        # Q3 budget projection = budget × 3 × quarterly seasonality adjustment
        frm(ws,r,6,f"=B{r}*3*1.05",DOLR)
    ws.cell(row=r,column=7,value=comment).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    r+=1

# Colour the variance column conditionally (manual — no Excel conditional formatting needed since we set colour in formula)
r+=2
sec(ws,r,7,"KEY VARIANCE DRIVERS — CFO NOTES",NAVY); r+=1
notes = [
    "Revenue beat driven by Product (+$1.34M) and Service (+$1.22M) — NSW and VIC BUs outperformed",
    "FX Gains $144K favourable — AUD/USD weakening benefited international revenue streams",
    "Advertising under-budget by $61K — digital campaign timing; budget to redeploy in Q3",
    "Contractor costs $108K under — project delays pushed to Q3; risk of Q3 cost spike",
    "Freight $94K favourable — route optimisation; sustainability at risk with fuel cost increases in Risk scenario",
]
for note in notes:
    lbl(ws,r,1,f"• {note}"); r+=1

print("Budget Variance done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 9 – SCENARIOS (side-by-side comparison)
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_sc
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col,w in [('B',18),('C',18),('D',18),('E',18),('F',16)]:
    ws.column_dimensions[col].width = w

hdr(ws,1,1,"SCENARIO COMPARISON — Q3 2026 (JUL+AUG+SEP COMBINED)",bg="C00000",merge=(1,6))
ws.row_dimensions[1].height = 28

for col,txt,bg in [(1,"KPI / Driver","C00000"),(2,"May 2026 Actual",DKGRAY),
                   (3,"Base Scenario",MID),(4,"Upside (+20%)",BGREEN),(5,"Risk-Adjusted","C00000"),
                   (6,"Upside vs Risk Δ","FFC000")]:
    c = ws.cell(row=2,column=col,value=txt)
    c.font = Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill = PatternFill("solid",start_color=bg if bg!="008000" else "006400")
    c.alignment = Alignment(horizontal="center" if col>1 else "left",vertical="center")

# We build this as a static reference table pulling from P&L Monthly
# switching via INDIRECT so each column hardcodes Base/Upside/Risk assumptions
# For simplicity: use driver-calculated values directly

DR = "Drivers!"  # We'll reference base/upside/risk directly

def sc_kpi(ws, r, label, b_val, base_f, up_f, risk_f, fmt=DOLR, bold=False):
    lbl(ws,r,1,label,bold=bold)
    if isinstance(b_val,str): frm(ws,r,2,b_val,fmt,bold,color=BGREEN)
    else: inp(ws,r,2,b_val,fmt,bold)
    frm(ws,r,3,base_f,fmt,bold)
    frm(ws,r,4,up_f,fmt,bold,color="006400")
    frm(ws,r,5,risk_f,fmt,bold,color="C00000")
    frm(ws,r,6,f"=D{r}-E{r}",fmt,bold)

# Use Revenue Model and P&L Monthly via cell references
# Q3 totals = col F in the respective sheets

PL="'P&L Monthly'!"
RM="'Revenue Model'!"

r=3
sec(ws,r,6,"REVENUE METRICS","C00000"); r+=1

# To show per-scenario values, we need to calculate them for each scenario
# We do this by noting the active value depends on Drivers!B2 (the dropdown)
# For a true side-by-side, we reference the driver rows B/C/D directly
# Revenue = Drivers!B7*(1+Drivers!Bx8) × sum of seasonality

def q3_rev(scen_col):
    """scen_col: B=Base, C=Upside, D=Risk"""
    return (f"=Drivers!{scen_col}7*(1+Drivers!{scen_col}8)*"
            f"(Drivers!{scen_col}9+Drivers!{scen_col}10+Drivers!{scen_col}11)")

sc_kpi(ws,r,"Q3 Total Revenue","=40546176*3",q3_rev("B"),q3_rev("C"),q3_rev("D"),DOLR,True); r+=1

def q3_ni(scen_col):
    sc={"B":"Base","C":"Upside","D":"Risk-Adjusted"}[scen_col]
    # Net income ≈ Revenue × (1 - cost ratio) × (1-tax)
    # Simplified: Reference P&L Monthly col F (Q3 total) when that scenario is active
    # Better: calculate inline using driver ratios
    return (f"=Drivers!{scen_col}7*(1+Drivers!{scen_col}8)*"
            f"(Drivers!{scen_col}9+Drivers!{scen_col}10+Drivers!{scen_col}11)*"
            f"(1-Drivers!{scen_col}16-Drivers!{scen_col}22-Drivers!{scen_col}23-"
            f"Drivers!{scen_col}24-Drivers!{scen_col}25-Drivers!{scen_col}27)*"
            f"(1-Drivers!{scen_col}29)")

def q3_ebitda(scen_col):
    return (f"=Drivers!{scen_col}7*(1+Drivers!{scen_col}8)*"
            f"(Drivers!{scen_col}9+Drivers!{scen_col}10+Drivers!{scen_col}11)*"
            f"(1-Drivers!{scen_col}16-Drivers!{scen_col}22-Drivers!{scen_col}23-"
            f"Drivers!{scen_col}24-Drivers!{scen_col}25-Drivers!{scen_col}27)")

sc_kpi(ws,r,"Q3 Revenue — Jul","=40546176*0.95",
       f"=Drivers!B7*(1+Drivers!B8)*Drivers!B9",
       f"=Drivers!C7*(1+Drivers!C8)*Drivers!C9",
       f"=Drivers!D7*(1+Drivers!D8)*Drivers!D9",DOLR); r+=1
sc_kpi(ws,r,"Q3 Revenue — Aug","=40546176*1.00",
       f"=Drivers!B7*(1+Drivers!B8)*Drivers!B10",
       f"=Drivers!C7*(1+Drivers!C8)*Drivers!C10",
       f"=Drivers!D7*(1+Drivers!D8)*Drivers!D10",DOLR); r+=1
sc_kpi(ws,r,"Q3 Revenue — Sep","=40546176*1.08",
       f"=Drivers!B7*(1+Drivers!B8)*Drivers!B11",
       f"=Drivers!C7*(1+Drivers!C8)*Drivers!C11",
       f"=Drivers!D7*(1+Drivers!D8)*Drivers!D11",DOLR); r+=1

sec(ws,r,6,"PROFITABILITY METRICS",MID); r+=1
sc_kpi(ws,r,"Q3 EBITDA","=11290713*3",q3_ebitda("B"),q3_ebitda("C"),q3_ebitda("D"),DOLR,True); r+=1
sc_kpi(ws,r,"Q3 Net Income","=9075108*3",q3_ni("B"),q3_ni("C"),q3_ni("D"),DOLR,True); r+=1

# Margins
lbl(ws,r,1,"Net Income Margin %")
frm(ws,r,2,"=22.4%",PCT); 
frm(ws,r,3,f"=D{r-1}/D{r-4}",PCT); frm(ws,r,4,f"=E{r-1}/E{r-4}",PCT,color="006400")
frm(ws,r,5,f"=F{r-1}/F{r-4}",PCT,color="C00000"); frm(ws,r,6,f"=D{r}-E{r}",PCT); r+=1

lbl(ws,r,1,"EBITDA Margin %")
frm(ws,r,2,"=27.8%",PCT)
frm(ws,r,3,f"=D{r-2}/D{r-5}",PCT); frm(ws,r,4,f"=E{r-2}/E{r-5}",PCT,color="006400")
frm(ws,r,5,f"=F{r-2}/F{r-5}",PCT,color="C00000"); frm(ws,r,6,f"=D{r}-E{r}",PCT); r+=1

sec(ws,r,6,"KEY DRIVER DIFFERENTIALS","FFC000"); r+=1
driver_compare = [
    ("Revenue Growth Rate",8,"=Drivers!B8","=Drivers!C8","=Drivers!D8",PCT),
    ("Freight Cost % Revenue",22,"=Drivers!B22","=Drivers!C22","=Drivers!D22",PCT),
    ("Days Sales Outstanding (DSO)",21,"=Drivers!B21","=Drivers!C21","=Drivers!D21",NUM),
    ("Current AR Collection Rate",22,"=Drivers!B22","=Drivers!C22","=Drivers!D22",PCT),
    (">90d AR Collection Rate",25,"=Drivers!B25","=Drivers!C25","=Drivers!D25",PCT),
    ("Insurance % Revenue",19,"=Drivers!B19","=Drivers!C19","=Drivers!D19",PCT),
    ("Supplier Price Uplift",26,"=Drivers!B26","=Drivers!C26","=Drivers!D26",PCT),
    ("Capital Expenditure (monthly)",28,"=Drivers!B28","=Drivers!C28","=Drivers!D28",DOLR),
]
for dname, drow_num, bf, uf, rf, fmt in driver_compare:
    lbl(ws,r,1,f"  {dname}",indent=1)
    ws.cell(row=r,column=2,value="—").font=Font(name="Arial",size=10,color=DKGRAY)
    frm(ws,r,3,bf,fmt); frm(ws,r,4,uf,fmt,color="006400"); frm(ws,r,5,rf,fmt,color="C00000")
    frm(ws,r,6,f"=D{r}-E{r}",fmt); r+=1

print("Scenarios done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 10 – RISK & OPPS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_ro
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 32
ws.column_dimensions['B'].width = 22
ws.column_dimensions['C'].width = 12
ws.column_dimensions['D'].width = 12
ws.column_dimensions['E'].width = 14
ws.column_dimensions['F'].width = 14
ws.column_dimensions['G'].width = 28
ws.column_dimensions['H'].width = 28

hdr(ws,1,1,"RISK REGISTER & OPPORTUNITIES — Q3 2026",bg="C00000",merge=(1,8))
ws.row_dimensions[1].height = 28

for col,txt,bg in [(1,"Risk / Opportunity","C00000"),(2,"Category",DKGRAY),(3,"Likelihood",DKGRAY),
                   (4,"Impact",DKGRAY),(5,"P&L Impact ($)",DKGRAY),(6,"Cash Impact ($)",DKGRAY),
                   (7,"Mitigation / Action","006400"),(8,"Owner",DKGRAY)]:
    c = ws.cell(row=2,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center",wrap_text=True)

risks = [
    # Risks
    ("RISKS","","","","","","",""),
    ("Middle East conflict — fuel surcharge spike","External / Macro","High","High",
     "-$620K/mo","−$620K/mo","Lock in freight contracts; seek fuel adjustment clauses","CFO / Procurement"),
    ("Overdue AR — Incitec & Orica (>90d)","Credit / Collections","High","High",
     "-$532K write-off risk","−$532K cash gap","Escalate to CEO level; engage legal/debt recovery","AR Manager"),
    ("FX depreciation (AUD vs USD/NZD)","Market / FX","Medium","High",
     "-$300K/qtr","−$300K/qtr","Consider FX hedging for USD exposures >$500K","Treasurer"),
    ("Contractor cost spike — deferred Q2 projects hit Q3","Operational","High","Medium",
     "-$200K incremental","−$200K","Scope & stage projects; convert to FTE where >6mo","CFO / HR"),
    ("Supply chain disruption — Tier 2 suppliers","Supply Chain","Medium","Medium",
     "-$180K","−$180K","Qualify alternative suppliers; increase safety stock","COO"),
    ("Subscription churn if service delivery impacted","Revenue","Low","High",
     "-$1.5M ARR risk","−$125K/mo","SLA monitoring; customer success programme","CRO"),
    ("Energy & utility cost escalation","Operational","Medium","Low",
     "-$95K/qtr","−$95K","Fixed-rate energy contracts for 12 months","COO"),
    ("Auckland Airport overdue $493K — cash gap","Collections","High","Medium",
     "0 (already booked)","−$493K cash","Weekly follow-up; offer settlement discount of 2%","AR Manager"),
    ("","","","","","","",""),
    ("OPPORTUNITIES","","","","","","",""),
    ("Revenue upside — Sep quarter-end acceleration","Revenue","Medium","High",
     "+$1.2M","+$1.0M","Pre-close pipeline review; accelerate invoicing cycle","Sales / CFO"),
    ("Subscription mix shift — lower cost to serve","Margin","Medium","Medium",
     "+$350K margin","0","Promote subscription renewals and upsell","CRO"),
    ("Extend DPO with key suppliers (Risk scenario)","Working Capital","High","Medium",
     "0","+$800K cash timing","Negotiate 45d terms with top 5 suppliers","Procurement"),
    ("Digital transformation Deloitte project — productivity gain","Strategic","Low","High",
     "+$500K run-rate","0","Ensure Q3 milestones on track to capture efficiency savings","CEO / CTO"),
    ("Google Ads prepayment — optimise ROI","Marketing","Medium","Medium",
     "+$200K revenue","0","Re-allocate remaining $500K spend to highest-converting channels","CMO"),
]

r = 3
for row_data in risks:
    label = row_data[0]
    if label in ("RISKS","OPPORTUNITIES"):
        bg_color = "C00000" if label=="RISKS" else "006400"
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
        c = ws.cell(row=r,column=1,value=label)
        c.font=Font(name="Arial",bold=True,color=WHITE,size=11)
        c.fill=PatternFill("solid",start_color=bg_color)
        c.alignment=Alignment(horizontal="left",vertical="center",indent=1)
        r+=1; continue
    if label=="":
        r+=1; continue
    
    label,cat,lik,imp,pl_imp,cash_imp,mitig,owner = row_data
    is_opp = pl_imp.startswith("+")
    row_bg = None
    if lik=="High" and imp=="High": row_bg=RED
    elif lik=="High" or imp=="High": row_bg=AMBER
    
    for col in range(1,9):
        if row_bg: ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=row_bg)
    
    ws.cell(row=r,column=1,value=label).font=Font(name="Arial",size=10,bold=True)
    ws.cell(row=r,column=1).alignment=Alignment(wrap_text=True,vertical="center")
    ws.cell(row=r,column=2,value=cat).font=Font(name="Arial",size=9,color=DKGRAY)
    ws.cell(row=r,column=3,value=lik).font=Font(name="Arial",size=10,
        color="C00000" if lik=="High" else "ED7D31" if lik=="Medium" else "006400")
    ws.cell(row=r,column=4,value=imp).font=Font(name="Arial",size=10,
        color="C00000" if imp=="High" else "ED7D31" if imp=="Medium" else "006400")
    ws.cell(row=r,column=5,value=pl_imp).font=Font(name="Arial",size=10,
        color="006400" if is_opp else "C00000")
    ws.cell(row=r,column=6,value=cash_imp).font=Font(name="Arial",size=10,
        color="006400" if is_opp else "C00000")
    ws.cell(row=r,column=7,value=mitig).font=Font(name="Arial",size=9,italic=True)
    ws.cell(row=r,column=7).alignment=Alignment(wrap_text=True,vertical="center")
    ws.cell(row=r,column=8,value=owner).font=Font(name="Arial",size=9,color=MID)
    ws.row_dimensions[r].height=30
    r+=1

print("Risk & Opps done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 11 – CFO DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dash
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 30
ws.column_dimensions['C'].width = 20
ws.column_dimensions['D'].width = 20
ws.column_dimensions['E'].width = 20
ws.column_dimensions['F'].width = 20
ws.column_dimensions['G'].width = 20

hdr(ws,1,1,"CFO EXECUTIVE DASHBOARD — Q3 2026 FORECAST",bg=NAVY,merge=(1,7))
hdr(ws,2,1,f"Active Scenario: ",bg=LGRAY,fg=NAVY,bold=True,merge=(1,1),align="left")
ws.merge_cells("C2:D2")
c=ws.cell(row=2,column=3,value="=Drivers!B2")
c.font=Font(name="Arial",bold=True,size=12,color="C00000")
c.fill=PatternFill("solid",start_color=AMBER)
c.alignment=Alignment(horizontal="center",vertical="center")
ws.row_dimensions[1].height=30; ws.row_dimensions[2].height=24

DR="Drivers!E"; PL="'P&L Monthly'!"; RM="'Revenue Model'!"; CF="'Cash Flow'!"

# KPI cards row
r=4
kpi_cards = [
    ("Q3 Revenue",f"={RM}F{rev_model_rev_row}",DOLR,MID),
    ("Q3 EBITDA",f"={PL}F{ebitda_row}",DOLR,"006400"),
    ("EBITDA Margin",f"={PL}G{ebitda_row}",PCT,"006400"),
    ("Q3 Net Income",f"={PL}F{ni_row}",DOLR,MID),
    ("Net Margin",f"={PL}G{ni_row}",PCT,MID),
    ("Closing Cash (Sep)",f"={CF}E{close_row}",DOLR,"7030A0"),
]
hdr(ws,r,2,"KEY PERFORMANCE INDICATORS — Q3 2026",bg=NAVY,merge=(1,6)); r+=1
for i,(label,formula,fmt,color) in enumerate(kpi_cards):
    col = 2+i
    if col > 7: break
    hdr(ws,r,col,label,bg=LGRAY,fg=NAVY,bold=True)
    c=ws.cell(row=r+1,column=col,value=formula)
    c.font=Font(name="Arial",bold=True,size=14,color=color)
    c.number_format=fmt
    c.alignment=Alignment(horizontal="center",vertical="center")
    c.fill=PatternFill("solid",start_color="F8F8F8")
ws.row_dimensions[r].height=22; ws.row_dimensions[r+1].height=36
r+=3

# Monthly P&L summary
hdr(ws,r,2,"MONTHLY P&L SUMMARY",bg=MID,merge=(1,6)); r+=1
for col,txt,bg in [(2,"Line Item",MID),(3,"Jul 2026",MID),(4,"Aug 2026",MID),(5,"Sep 2026",MID),(6,"Q3 Total",NAVY),(7,"NI Margin",NAVY)]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>2 else "left",vertical="center")
r+=1

dash_pl=[
    ("Total Revenue",pl_rev_row,DOLR,True),
    ("EBITDA",ebitda_row,DOLR,True),
    ("EBIT",ebit_row,DOLR,False),
    ("Net Income",ni_row,DOLR,True),
    ("EBITDA Margin",ebitda_row,PCT,False),
    ("Net Income Margin",ni_row,PCT,False),
]
for label,src_row,fmt,bold in dash_pl:
    lbl(ws,r,2,label,bold=bold)
    if fmt==PCT:
        # Margin = row / rev row
        ni_src = ebitda_row if "EBITDA" in label else ni_row
        frm(ws,r,3,f"={PL}C{ni_src}/{PL}C{pl_rev_row}",PCT,bold,color=BGREEN)
        frm(ws,r,4,f"={PL}D{ni_src}/{PL}D{pl_rev_row}",PCT,bold,color=BGREEN)
        frm(ws,r,5,f"={PL}E{ni_src}/{PL}E{pl_rev_row}",PCT,bold,color=BGREEN)
        frm(ws,r,6,f"={PL}F{ni_src}/{PL}F{pl_rev_row}",PCT,bold,color=BGREEN)
    else:
        frm(ws,r,3,f"={PL}C{src_row}",fmt,bold,color=BGREEN)
        frm(ws,r,4,f"={PL}D{src_row}",fmt,bold,color=BGREEN)
        frm(ws,r,5,f"={PL}E{src_row}",fmt,bold,color=BGREEN)
        frm(ws,r,6,f"={PL}F{src_row}",fmt,bold,color=BGREEN)
    r+=1

r+=1
# Cash flow summary
hdr(ws,r,2,"CASH FLOW SUMMARY",bg="7030A0",merge=(1,6)); r+=1
for col,txt in [(2,"Line Item"),(3,"Jul"),(4,"Aug"),(5,"Sep"),(6,"Q3 Total")]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color="7030A0")
    c.alignment=Alignment(horizontal="center" if col>2 else "left",vertical="center")
r+=1
for label,cf_row in [("Operating Cash Flow",ocf_row),("Net Cash Movement",net_move_row),("Closing Cash Balance",close_row)]:
    lbl(ws,r,2,label,bold=(label=="Closing Cash Balance"))
    for col,cf_col in [(3,"C"),(4,"D"),(5,"E")]:
        frm(ws,r,col,f"={CF}{cf_col}{cf_row}",DOLR,label=="Closing Cash Balance",color="7030A0" if col<6 else BGREEN)
    frm(ws,r,6,f"={CF}F{cf_row}" if label!="Closing Cash Balance" else f"={CF}E{cf_row}",DOLR,True,color="7030A0")
    r+=1

r+=1
# CFO Action Items
hdr(ws,r,2,"TOP CFO ACTIONS — Q3 2026",bg=NAVY,merge=(1,6)); r+=1
cfo_actions=[
    ("🔴 URGENT","Escalate Incitec ($245K) & Orica ($287K) >90d AR — debt recovery by Jul 15","AR Manager"),
    ("🔴 URGENT","Lock in freight contracts before fuel surcharge escalation — target Jul 1","Procurement"),
    ("🟡 HIGH","Negotiate extended DPO (42d) with top 5 suppliers to release $800K working capital","CFO"),
    ("🟡 HIGH","Implement FX hedging programme for USD exposures >$500K — Middle East risk","Treasurer"),
    ("🟢 MEDIUM","Redirect Google Ads $500K prepayment to highest-ROI channels before Sep renewal","CMO"),
    ("🟢 MEDIUM","Pre-close pipeline review in Sep — accelerate invoicing to capture quarter-end uplift","Sales/CFO"),
    ("🔵 INFO","Monitor Auckland Airport ($493K) — offer 2% early settlement discount","AR Manager"),
    ("🔵 INFO","Review contractor project scope — $108K deferred from May may hit Q3 cost base","CFO / HR"),
]
for priority,action,owner in cfo_actions:
    ws.cell(row=r,column=2,value=priority).font=Font(name="Arial",bold=True,size=10)
    ws.cell(row=r,column=3,value=action).font=Font(name="Arial",size=9)
    ws.cell(row=r,column=3).alignment=Alignment(wrap_text=True)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
    ws.cell(row=r,column=7,value=owner).font=Font(name="Arial",size=9,color=MID)
    ws.row_dimensions[r].height=28
    r+=1

print("CFO Dashboard done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 12 – MACRO GUIDE
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_mg
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 28
ws.column_dimensions['B'].width = 50
ws.column_dimensions['C'].width = 22

hdr(ws,1,1,"VBA MACRO GUIDE & USER INSTRUCTIONS",bg=DKGRAY,merge=(1,3))
ws.row_dimensions[1].height=28

r=3
macros=[
    ("SCENARIO CONTROL","",""),
    ("Switch to Base Scenario","In Excel: Press Alt+F8 → Run 'SwitchToBase'","Sets Drivers!B2 = 'Base' and refreshes all sheets"),
    ("Switch to Upside","Alt+F8 → Run 'SwitchToUpside'","Sets Drivers!B2 = 'Upside'"),
    ("Switch to Risk-Adjusted","Alt+F8 → Run 'SwitchToRisk'","Sets Drivers!B2 = 'Risk-Adjusted'"),
    ("","",""),
    ("NAVIGATION","",""),
    ("Go to Dashboard","Alt+F8 → 'GotoDashboard'","Jumps to CFO Dashboard tab"),
    ("Go to Drivers","Alt+F8 → 'GotoDrivers'","Jumps to scenario control hub"),
    ("Go to P&L","Alt+F8 → 'GotoPL'","Jumps to P&L Monthly sheet"),
    ("Go to Cash Flow","Alt+F8 → 'GotoCashFlow'","Jumps to Cash Flow sheet"),
    ("","",""),
    ("ANALYSIS TOOLS","",""),
    ("Highlight Variances","Alt+F8 → 'HighlightVariances'","Highlights cells >10% variance from budget in red"),
    ("Clear Highlights","Alt+F8 → 'ClearHighlights'","Removes all conditional highlighting"),
    ("Refresh All","Alt+F8 → 'RefreshAll'","Forces full workbook recalculation"),
    ("","",""),
    ("EXPORT","",""),
    ("Export to PDF","Alt+F8 → 'ExportToPDF'","Saves all key sheets as PDF to same folder"),
    ("","",""),
    ("KEY ASSUMPTIONS TO CHANGE","",""),
    ("Switch scenario","Drivers!B2 — use dropdown","Base / Upside / Risk-Adjusted"),
    ("Revenue baseline","Drivers!B7 / C7 / D7","May 2026 actual $40.5M"),
    ("Growth rate","Drivers!B8 / C8 / D8","Base 8.8%, Upside 20%, Risk 4.5%"),
    ("DSO","Drivers!B21 / C21 / D21","Base 45d, Risk 55d"),
    ("Freight % Rev","Drivers!B22 / C22 / D22","Risk 6.2% vs Base 4.8%"),
    ("Supplier uplift","Drivers!B26 / C26 / D26","Risk 3.5% additional cost"),
    ("","",""),
    ("COLOUR LEGEND","",""),
    ("Blue text","Hardcoded input — safe to change for scenarios",""),
    ("Black text","Formula — do not manually overwrite",""),
    ("Green text","Cross-sheet link — automatically updates",""),
    ("Yellow background","Key assumption requiring CFO review",""),
    ("Red background","High-risk metric or at-risk cash item",""),
]

sec(ws,3,3,"MACRO REFERENCE",DKGRAY); r=4
for row_data in macros:
    name,how,desc=row_data
    if name in ("SCENARIO CONTROL","NAVIGATION","ANALYSIS TOOLS","EXPORT","KEY ASSUMPTIONS TO CHANGE","COLOUR LEGEND"):
        sec(ws,r,3,name,MID); r+=1; continue
    if name=="":
        r+=1; continue
    ws.cell(row=r,column=1,value=name).font=Font(name="Arial",bold=True,size=10,color=NAVY)
    ws.cell(row=r,column=2,value=how).font=Font(name="Arial",size=10)
    ws.cell(row=r,column=3,value=desc).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    ws.cell(row=r,column=3).alignment=Alignment(wrap_text=True)
    ws.row_dimensions[r].height=20
    r+=1

print("Macro Guide done")

# ═══════════════════════════════════════════════════════════════════════════
# FREEZE PANES & ROW HEIGHTS
# ═══════════════════════════════════════════════════════════════════════════
for ws_name, freeze_cell in [
    ("Drivers","A6"),("Revenue Model","A3"),("Cost Model","A3"),
    ("P&L Monthly","A3"),("Cash Flow","A3"),("Working Capital","A3"),
    ("Budget Variance","A3"),("Scenarios","A3"),
]:
    wb[ws_name].freeze_panes = freeze_cell

# ═══════════════════════════════════════════════════════════════════════════
# SAVE AS XLSX FIRST (then we'll inject VBA to make .xlsm)
# ═══════════════════════════════════════════════════════════════════════════
OUT_XLSX = "/sessions/practical-blissful-maxwell/mnt/outputs/Driver_Forecast_Q3_2026.xlsx"
OUT_XLSM = "/sessions/practical-blissful-maxwell/mnt/outputs/Driver_Forecast_Q3_2026.xlsm"

wb.save(OUT_XLSX)
print(f"Saved XLSX: {OUT_XLSX}")

# ═══════════════════════════════════════════════════════════════════════════
# VBA INJECTION — Build vbaProject.bin and convert to .xlsm
# ═══════════════════════════════════════════════════════════════════════════
import struct, io, zipfile, shutil, re

VBA_CODE = r'''
Attribute VB_Name = "Module1"
' ===== SCENARIO SWITCHERS =====
Sub SwitchToBase()
    SetScenario "Base"
End Sub

Sub SwitchToUpside()
    SetScenario "Upside"
End Sub

Sub SwitchToRisk()
    SetScenario "Risk-Adjusted"
End Sub

Sub SetScenario(scenName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Drivers")
    ws.Range("B2").Value = scenName
    Application.Calculate
    MsgBox "Switched to: " & scenName & Chr(13) & Chr(13) & _
           "All sheets have been updated.", vbInformation, "Scenario Updated"
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Scenario Switch Failed"
End Sub

' ===== NAVIGATION =====
Sub GotoDashboard()
    GoToSheet "CFO Dashboard", "A1"
End Sub

Sub GotoDrivers()
    GoToSheet "Drivers", "B2"
End Sub

Sub GotoPL()
    GoToSheet "P&L Monthly", "A1"
End Sub

Sub GotoCashFlow()
    GoToSheet "Cash Flow", "A1"
End Sub

Sub GoToSheet(sheetName As String, cellAddr As String)
    On Error GoTo ErrHandler
    ThisWorkbook.Sheets(sheetName).Activate
    ThisWorkbook.Sheets(sheetName).Range(cellAddr).Select
    Exit Sub
ErrHandler:
    MsgBox "Sheet not found: " & sheetName, vbCritical
End Sub

' ===== HIGHLIGHT VARIANCES =====
Sub HighlightVariances()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Budget Variance")
    ws.Activate
    Dim c As Range
    For Each c In ws.UsedRange
        If IsNumeric(c.Value) And c.Column = 4 Then
            If Abs(c.Value) > 100000 Then
                If c.Value < 0 Then
                    c.Interior.Color = RGB(255, 199, 206)
                Else
                    c.Interior.Color = RGB(198, 239, 206)
                End If
            End If
        End If
    Next c
    MsgBox "Variances highlighted. Run ClearHighlights to reset.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Sub ClearHighlights()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Cells.Interior.ColorIndex = xlNone
        ' Note: this clears all fills, reopen to restore formatting
    Next ws
    MsgBox "Highlights cleared.", vbInformation
End Sub

' ===== REFRESH =====
Sub RefreshAll()
    Application.CalculateFull
    MsgBox "All formulas recalculated.", vbInformation, "Refresh Complete"
End Sub

' ===== SHOW SCENARIO =====
Sub ShowCurrentScenario()
    Dim scenName As String
    scenName = ThisWorkbook.Sheets("Drivers").Range("B2").Value
    MsgBox "Active Scenario: " & scenName, vbInformation, "Current Scenario"
End Sub

' ===== EXPORT TO PDF =====
Sub ExportToPDF()
    On Error GoTo ErrHandler
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Driver_Forecast_Q3_2026_Export.pdf"
    
    Dim sheetsToPrint As Variant
    sheetsToPrint = Array("CFO Dashboard", "P&L Monthly", "Cash Flow", "Scenarios", "Risk & Opps")
    
    Dim i As Integer
    For i = 0 To UBound(sheetsToPrint)
        ThisWorkbook.Sheets(sheetsToPrint(i)).Select (i = 0)
    Next i
    
    ThisWorkbook.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False
    
    MsgBox "PDF exported to:" & Chr(13) & pdfPath, vbInformation, "Export Complete"
    Exit Sub
ErrHandler:
    MsgBox "PDF export failed: " & Err.Description, vbCritical
End Sub

' ===== PROTECT INPUTS =====
Sub ProtectInputs()
    Dim ws As Worksheet
    Dim pwd As String
    pwd = "CFO2026"
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Drivers" Then
            ws.Protect Password:=pwd, UserInterfaceOnly:=True
        End If
    Next ws
    MsgBox "All sheets except Drivers are now protected. Password: CFO2026", vbInformation
End Sub

Sub UnprotectAll()
    Dim ws As Worksheet
    Dim pwd As String
    pwd = "CFO2026"
    For Each ws In ThisWorkbook.Sheets
        ws.Unprotect Password:=pwd
    Next ws
    MsgBox "All sheets unprotected.", vbInformation
End Sub
'''

# ─── OLE2/CFB minimal vbaProject.bin builder ────────────────────────────────
SECT_SIZE = 512
MINI_STREAM_CUTOFF = 4096

def pad(data, size=SECT_SIZE):
    if len(data) % size:
        data += b'\xff' * (size - len(data) % size)
    return data

def ole_header(fat_sectors, dir_sectors, num_fat, first_dir, first_mini_fat=-1, num_mini_fat=0):
    hdr = bytearray(512)
    # Magic
    hdr[0:8] = bytes([0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1])
    # Minor/major version
    struct.pack_into('<HH',hdr,24,0x003E,0x0003)
    # Byte order
    struct.pack_into('<H',hdr,28,0xFFFE)
    # Sector size (512 = 2^9)
    struct.pack_into('<H',hdr,30,9)
    # Mini sector size (64 = 2^6)
    struct.pack_into('<H',hdr,32,6)
    # Dir sectors
    struct.pack_into('<I',hdr,40,0)
    # FAT sectors
    struct.pack_into('<I',hdr,44,num_fat)
    # First dir sector
    struct.pack_into('<I',hdr,48,first_dir)
    # Mini stream cutoff
    struct.pack_into('<I',hdr,56,MINI_STREAM_CUTOFF)
    # First mini FAT sector (-2 = FREESECT if none)
    struct.pack_into('<I',hdr,60, first_mini_fat & 0xFFFFFFFF)
    # Num mini FAT sectors
    struct.pack_into('<I',hdr,64,num_mini_fat)
    # First DIFAT sector (-2=none)
    struct.pack_into('<I',hdr,68,0xFFFFFFFE)
    # Num DIFAT sectors
    struct.pack_into('<I',hdr,72,0)
    # DIFAT array (109 entries at offset 76)
    for i,s in enumerate(fat_sectors[:109]):
        struct.pack_into('<I',hdr,76+i*4,s)
    for i in range(len(fat_sectors),109):
        struct.pack_into('<I',hdr,76+i*4,0xFFFFFFFF)
    return bytes(hdr)

def dir_entry(name, entry_type, color, child, left_sib, right_sib,
              start_sect, size, clsid=None):
    entry = bytearray(128)
    name_enc = name.encode('utf-16-le')[:62]
    name_len = len(name_enc) + 2
    entry[0:len(name_enc)] = name_enc
    struct.pack_into('<H',entry,64,name_len)
    entry[66] = entry_type  # 1=storage,2=stream,5=root
    entry[67] = color       # 0=red,1=black
    struct.pack_into('<I',entry,68, left_sib & 0xFFFFFFFF)
    struct.pack_into('<I',entry,72, right_sib & 0xFFFFFFFF)
    struct.pack_into('<I',entry,76, child & 0xFFFFFFFF)
    if clsid:
        entry[80:96] = clsid
    struct.pack_into('<I',entry,116, start_sect & 0xFFFFFFFF)
    struct.pack_into('<Q',entry,120, size)
    return bytes(entry)

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD

# We'll build a minimal OLE2 with:
# Sector 0: FAT
# Sector 1-2: Directory (2 sectors = 8 entries)
# Sector 3-N: VBA_PROJECT stream (padded to >= 4096 to stay in regular stream)
# Sector N+1..M: dir stream (compressed VBA)
# Sector M+1..P: Module1 stream (compressed VBA code)

# Build streams
def compress_vba(source: bytes) -> bytes:
    """MS-OVBA raw-chunk compression (simplified: use raw chunks only)"""
    result = bytearray([0x01])  # SignatureByte
    i = 0
    while i < len(source):
        chunk = source[i:i+4096]
        # Raw chunk: flag byte 0x00..0x02 then data
        # CompressedChunkHeader: isCompressed bit = 0 means raw (4096 bytes uncompressed)
        if len(chunk) < 4096:
            chunk = chunk + b'\x00' * (4096 - len(chunk))
        # Raw chunk header: size-1 in bits 0-11, isCompressed=0 in bit 15
        raw_size = 4096
        hdr_val = (raw_size - 1) & 0x0FFF  # isCompressed=0
        result += struct.pack('<H', hdr_val)
        result += chunk
        i += 4096
    return bytes(result)

vba_source = VBA_CODE.encode('latin-1')
# Pad to multiple of 4096 so compression is straightforward
padded_len = ((len(vba_source) + 4095) // 4096) * 4096
vba_source_padded = vba_source + b'\x00' * (padded_len - len(vba_source))
module1_compressed = compress_vba(vba_source_padded)

# _VBA_PROJECT is a required stream but can be minimal placeholder
vba_project_data = b'\xCC\x61' + b'\x00' * 62  # minimal valid header
# Pad to >= 4096 so it uses regular stream (not mini stream)
vba_project_data = vba_project_data + b'\x00' * (4096 - len(vba_project_data))

# dir stream: describes modules — minimal but structurally valid
# We build a real dir stream per MS-OVBA spec
def make_dir_stream(module_name="Module1", module_offset=0, module_compressed_size=None,
                     module_source_size=None):
    if module_compressed_size is None: module_compressed_size = len(module1_compressed)
    if module_source_size is None: module_source_size = len(vba_source_padded)
    buf = bytearray()
    def rec(id_, data):
        buf.extend(struct.pack('<HI', id_, len(data)))
        buf.extend(data)
    def recstr(id_, s):
        rec(id_, s.encode('latin-1'))
    # PROJECTSYSKIND
    rec(0x0001, struct.pack('<I',0x00000001))  # Win32
    # PROJECTLCID
    rec(0x0002, struct.pack('<I',0x0409))
    # PROJECTLCIDINVOKE
    rec(0x0014, struct.pack('<I',0x0409))
    # PROJECTCODEPAGE
    rec(0x0003, struct.pack('<H',1252))
    # PROJECTNAME
    recstr(0x0004,"VBAProject")
    # PROJECTDOCSTRING
    recstr(0x0005,"Driver-Based Forecast VBA")
    recstr(0x0040,u"Driver-Based Forecast VBA".encode('utf-16-le').decode('latin-1'))
    # PROJECTHELPFILEPATH
    recstr(0x0006,""); recstr(0x003D,"")
    # PROJECTHELPCONTEXT
    rec(0x0007, struct.pack('<I',0))
    # PROJECTLIBFLAGS
    rec(0x0008, struct.pack('<I',0))
    # PROJECTVERSION
    rec(0x0009, struct.pack('<IH',0x51,0x000D))
    # PROJECTCONSTANTS
    recstr(0x000C,""); recstr(0x003C,"")
    # PROJECTREFERENCES — none
    # PROJECTMODULES: count=1
    rec(0x000F, struct.pack('<H',1))  # Count
    rec(0x0013, struct.pack('<H',0x0047))  # ProjectCookie
    # MODULE
    recstr(0x0019, module_name)           # MODULENAME
    recstr(0x0031, module_name)           # MODULENAMEUNICODE (as latin-1 encoded utf16le)
    recstr(0x001A, "")                    # MODULESTREAMNAME
    recstr(0x0032, "")
    recstr(0x001C, "")                    # MODULEDOCSTRING
    recstr(0x0048, "")
    rec(0x0031, struct.pack('<I', module_offset))   # MODULEOFFSET
    rec(0x001E, struct.pack('<I',1))                # MODULETYPE: procedural
    rec(0x002B, struct.pack('<I',0))                # MODULEREADONLY (0=no)
    rec(0x002C, struct.pack('<I',0))                # MODULEPRIVATE
    # MODULETERM
    buf.extend(struct.pack('<HI',0x002B,0))
    # PROJECTMODULES terminator
    buf.extend(struct.pack('<HI',0x0010,0))
    return compress_vba(bytes(buf))

dir_stream_data = make_dir_stream("Module1", 0)

# Module1 stream = compressed VBA source
module1_stream_data = module1_compressed

def sectors_for(data):
    n = (len(data) + SECT_SIZE - 1) // SECT_SIZE
    return n

# Lay out sectors:
# 0: FAT
# 1,2: Directory (2 sectors)
# 3..3+vba_proj_sects-1: _VBA_PROJECT
# next..+dir_sects-1: dir stream
# next..+mod_sects-1: Module1 stream

vba_proj_sects = sectors_for(vba_project_data)
dir_sects = sectors_for(dir_stream_data)
mod_sects = sectors_for(module1_stream_data)

s_fat   = 0
s_dir   = 1           # 2 sectors
s_vba   = 3           # _VBA_PROJECT
s_dstr  = s_vba + vba_proj_sects
s_mod   = s_dstr + dir_sects
total_sectors = s_mod + mod_sects

# Build FAT
fat = [FREESECT] * 128  # 1 FAT sector = 128 entries
fat[s_fat] = FATSECT
# Directory chain: sectors 1,2
fat[1] = 2; fat[2] = ENDOFCHAIN
# _VBA_PROJECT chain
for i in range(vba_proj_sects):
    fat[s_vba+i] = s_vba+i+1 if i < vba_proj_sects-1 else ENDOFCHAIN
# dir stream chain
for i in range(dir_sects):
    fat[s_dstr+i] = s_dstr+i+1 if i < dir_sects-1 else ENDOFCHAIN
# Module1 chain
for i in range(mod_sects):
    fat[s_mod+i] = s_mod+i+1 if i < mod_sects-1 else ENDOFCHAIN

fat_sector = b''.join(struct.pack('<I',x & 0xFFFFFFFF) for x in fat)

# Build Directory (8 entries over 2 sectors)
CLSID_VBAProject = bytes([
    0x06,0x09,0x02,0x00,0x00,0x00,0x00,0x00,
    0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46
])
# Entry layout: root(child=1), VBA storage(child=2), _VBA_PROJECT stream, dir stream, Module1, 3×empty
ROOT  = dir_entry("Root Entry",5,1,1,FREESECT,FREESECT,ENDOFCHAIN,0,CLSID_VBAProject)
VBA_S = dir_entry("VBA",       1,1,2,FREESECT,FREESECT,ENDOFCHAIN,0)
VBA_P = dir_entry("_VBA_PROJECT",2,1,FREESECT,3,4,s_vba,len(vba_project_data))
DIR_S = dir_entry("dir",       2,1,FREESECT,FREESECT,FREESECT,s_dstr,len(dir_stream_data))
MOD_S = dir_entry("Module1",   2,1,FREESECT,FREESECT,FREESECT,s_mod,len(module1_stream_data))
EMPTY = b'\x00'*128

dir_data = ROOT + VBA_S + VBA_P + DIR_S + MOD_S + EMPTY + EMPTY + EMPTY
# Must be exactly 2 sectors = 1024 bytes
dir_data = dir_data + b'\x00' * (1024 - len(dir_data))

# Build the complete OLE2 binary
ole_hdr = ole_header([s_fat], [], 1, s_dir)
body = bytearray()
body += fat_sector                               # sector 0: FAT
body += dir_data                                 # sectors 1-2: dir
body += pad(vba_project_data, SECT_SIZE)        # _VBA_PROJECT
body += pad(dir_stream_data, SECT_SIZE)         # dir stream
body += pad(module1_stream_data, SECT_SIZE)     # Module1

vba_bin = ole_hdr + bytes(body)
print(f"vbaProject.bin size: {len(vba_bin)} bytes")

# ─── INJECT INTO XLSM ────────────────────────────────────────────────────────
shutil.copy(OUT_XLSX, OUT_XLSM)

with zipfile.ZipFile(OUT_XLSM, 'a') as zf:
    # 1. Add vbaProject.bin
    zf.writestr("xl/vbaProject.bin", vba_bin)

# Now patch [Content_Types].xml and xl/_rels/workbook.xml.rels
import zipfile as zfmod

# Read current zip
with open(OUT_XLSM,'rb') as f:
    zip_bytes = f.read()

new_zip_buf = io.BytesIO()
with zfmod.ZipFile(io.BytesIO(zip_bytes),'r') as zin:
    with zfmod.ZipFile(new_zip_buf,'w',compression=zfmod.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                # Add vbaProject content type
                content_type_entry = b'<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.activeX+xml"/>'
                vba_ct = b'<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-excel.vbaProject"/>'
                if b'vbaProject' not in data:
                    data = data.replace(b'</Types>', vba_ct + b'</Types>')
                    # Also ensure workbook content type is xlsm
                    data = data.replace(
                        b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                        b'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
                    )
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                vba_rel = b'<Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
                if b'vbaProject' not in data:
                    data = data.replace(b'</Relationships>', vba_rel + b'</Relationships>')
            zout.writestr(item, data)

with open(OUT_XLSM,'wb') as f:
    f.write(new_zip_buf.getvalue())

print(f"Final .xlsm saved: {OUT_XLSM}")
print(f"File size: {os.path.getsize(OUT_XLSM):,} bytes")

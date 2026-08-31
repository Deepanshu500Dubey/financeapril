import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ─── COLOURS ────────────────────────────────────────────────────────────────
NAVY  = "1F3864"; MID   = "2E75B6"; LIGHT = "BDD7EE"
GREEN = "E2EFDA"; RED   = "FFE5E5"; AMBER = "FFF2CC"
WHITE = "FFFFFF"; DKGRAY= "404040"; LGRAY = "F2F2F2"
BBLUE = "0000FF"; BGREEN= "008000"; PURPLE= "7030A0"
ORNG  = "ED7D31"; TEAL  = "00B0F0"

# ─── HELPERS ────────────────────────────────────────────────────────────────
DOLR = '$#,##0'; PCT = '0.0%'; NUM = '#,##0'; NUM2='#,##0.00'

def hdr(ws, r, c, val, bg=NAVY, fg=WHITE, bold=True, sz=10, merge=None, align="center", wrap=False):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Arial", bold=bold, color=fg, size=sz)
    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if merge:
        ws.merge_cells(start_row=r, start_column=c, end_row=r+merge[0]-1, end_column=c+merge[1]-1)
    return cell

def inp(ws, r, c, val, fmt=None, bold=False):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Arial", color=BBLUE, bold=bold, size=10)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if fmt: cell.number_format = fmt
    return cell

def frm(ws, r, c, formula, fmt=None, bold=False, color="000000", bg=None):
    cell = ws.cell(row=r, column=c, value=formula)
    cell.font = Font(name="Arial", color=color, bold=bold, size=10)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if fmt: cell.number_format = fmt
    if bg: cell.fill = PatternFill("solid", start_color=bg)
    return cell

def lbl(ws, r, c, val, bold=False, color="000000", bg=None, indent=0, sz=10, wrap=False):
    cell = ws.cell(row=r, column=c, value=(" "*indent*2)+str(val))
    cell.font = Font(name="Arial", bold=bold, color=color, size=sz)
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    if bg: cell.fill = PatternFill("solid", start_color=bg)
    return cell

def sec(ws, r, cols, label, bg=MID, fg=WHITE):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    cell = ws.cell(row=r, column=1, value=label)
    cell.font = Font(name="Arial", bold=True, color=fg, size=10)
    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 18
    return r + 1

def fill_row(ws, r, cols, bg):
    for c in range(1, cols+1):
        ws.cell(row=r, column=c).fill = PatternFill("solid", start_color=bg)

# ─── WORKBOOK SETUP ─────────────────────────────────────────────────────────
wb = Workbook()
sheets = ["Cover","CC Baseline","Reallocation Engine","Regional P&L",
          "Consolidated P&L","Margin Analysis","Efficiency Metrics",
          "Risk Register","Optimal Strategy","Executive Summary"]
wb.active.title = "Cover"
for s in sheets[1:]:
    wb.create_sheet(s)

tab_colors = {
    "Cover":NAVY,"CC Baseline":"808080","Reallocation Engine":ORNG,
    "Regional P&L":"70AD47","Consolidated P&L":MID,"Margin Analysis":PURPLE,
    "Efficiency Metrics":TEAL,"Risk Register":"C00000",
    "Optimal Strategy":"006400","Executive Summary":NAVY}
for name,color in tab_colors.items():
    wb[name].sheet_properties.tabColor = color

# ─── SOURCE DATA (from May 2026 actuals) ────────────────────────────────────
TOTAL_REV = 40546176.42
TOTAL_EXP = 31471068.48
NET_INCOME = 9075107.94
EBITDA     = 11290712.92
NONCASH    = 2215605.0  # D&A + Amort

# Cost center assignments from P&L categories
CC_COSTS = {
    "IT":   {"Software Licenses":1107364.55, "Cloud Services":1354642.51, "IT Support":958460.94},
    "HR":   {"Salaries":1110732.19,"Benefits":1019737.24,"Contractor Fees":924312.86,
             "Training":1060256.02,"Recruitment":1310481.05},
    "OPS":  {"Telecommunications":1188371.93,"Postage & Courier":924546.02,"Insurance":1310206.50},
    "CORP": {"Office Rent":868729.19,"Utilities":914299.66,"Office Supplies":1030030.90,
             "Professional Services":958056.17,"Legal Fees":1371577.63,"Audit Fees":1340571.27,
             "Travel":1021794.25,"Accommodation":1122559.57,"Meals & Entertainment":1221692.84,
             "Bank Charges":1285755.28,"Interest Expense":966892.39,"FX Loss":1157387.93,
             "Depreciation":1330907.62,"Amortization":884697.36},
    "MKT":  {"Marketing":1391033.95,"Advertising":1237476.07,"Events":1098494.59},
    "NSW":  {"NSW Workforce":3571090.0},
    "QLD":  {"QLD Workforce":3582740.0},
    "WA":   {"WA Workforce":2324260.0},
    "VIC":  {"VIC Workforce":3401870.0},
    "SA":   {"SA Workforce":1447230.0},
}

CC_TOTALS = {cc: sum(v.values()) for cc, v in CC_COSTS.items()}
SUPPORT_CCS = ["IT","HR","OPS","CORP"]
REVENUE_CCS = ["NSW","QLD","WA","VIC","SA","MKT"]

SUPPORT_TOTAL = sum(CC_TOTALS[cc] for cc in SUPPORT_CCS)
print("Support total:", f"${SUPPORT_TOTAL:,.0f}")

# Revenue by region
REV_BY_CC = {"NSW":12569315,"QLD":7298312,"WA":6892850,"VIC":8920159,"SA":4865541,"MKT":0}
# MKT drives revenue but doesn't own it — kept as 0 for this model

# Reallocation parameters
REDUCTION_PCT = 0.10
SAVINGS = SUPPORT_TOTAL * REDUCTION_PCT
PROFIT_RETENTION = 0.40  # 40% falls to bottom line
REINVEST_PCT = 0.60       # 60% reinvested in revenue CCs
REINVEST = SAVINGS * REINVEST_PCT
PROFIT_BOOST = SAVINGS * PROFIT_RETENTION

# Reinvestment split to revenue regions
REINVEST_ALLOC = {"QLD":0.30,"NSW":0.30,"WA":0.25,"MKT":0.15}
# Revenue multiplier ($ revenue per $ reinvested) — based on past BU performance
REV_MULT = {"QLD":3.5,"NSW":3.0,"WA":3.2,"MKT":2.0}

print(f"Total savings: ${SAVINGS:,.0f}")
print(f"Profit boost: ${PROFIT_BOOST:,.0f}")
print(f"Reinvestment: ${REINVEST:,.0f}")
for cc in ["QLD","NSW","WA","MKT"]:
    alloc = REINVEST * REINVEST_ALLOC[cc]
    rev_uplift = alloc * REV_MULT[cc]
    print(f"  {cc}: reinvest ${alloc:,.0f} -> revenue uplift ${rev_uplift:,.0f}")


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 1: COVER
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Cover"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 36
ws.column_dimensions['C'].width = 24
ws.column_dimensions['D'].width = 24

hdr(ws,1,1,"STRATEGIC COST REALLOCATION MODEL — Q3 2026",bg=NAVY,sz=18,merge=(1,4))
hdr(ws,2,1,"Support Function Efficiency → Revenue Region Investment",bg=MID,sz=12,merge=(1,4))
hdr(ws,3,1,"Baseline: May 2026 Actuals  |  Forecast: Aug–Aug–Sep 2026",bg=LIGHT,fg=NAVY,sz=10,merge=(1,4))
ws.row_dimensions[1].height=36; ws.row_dimensions[2].height=24; ws.row_dimensions[3].height=20

r=5
kpis = [
    ("May 2026 Revenue","$40,546,176",MID),
    ("May 2026 EBITDA","$11,290,713","006400"),
    ("EBITDA Margin","27.8%","006400"),
    ("Net Income Margin","22.4%","006400"),
    ("Support Function Cost","$27,744,064","C00000"),
    ("Support as % Revenue","68.4%","C00000"),
    ("10% Efficiency Saving","$2,774,406",ORNG),
    ("Revenue Uplift (modelled)","$5,077,163","006400"),
]
for label,val,color in kpis:
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3)
    c=ws.cell(row=r,column=2,value=label)
    c.font=Font(name="Arial",bold=True,size=10,color=DKGRAY)
    c.fill=PatternFill("solid",start_color=LGRAY)
    c.alignment=Alignment(horizontal="left",indent=1)
    c2=ws.cell(row=r,column=4,value=val)
    c2.font=Font(name="Arial",bold=True,size=11,color=color)
    c2.fill=PatternFill("solid",start_color=LGRAY)
    c2.alignment=Alignment(horizontal="center")
    r+=1

r+=1
hdr(ws,r,2,"REALLOCATION LOGIC",bg=ORNG,merge=(1,3)); r+=1
logic=[
    ("IT + HR + OPS + CORP","Reduce costs 10% via automation & efficiency"),
    ("40% retained","Drops to EBITDA / Net Income immediately"),
    ("60% reinvested","Reallocated: QLD 30% | NSW 30% | WA 25% | MKT 15%"),
    ("Revenue multiplier","QLD 3.5x | NSW 3.0x | WA 3.2x | MKT 2.0x"),
    ("Quarterly execution","3-month ramp: Aug 80% | Aug 90% | Sep 100% of saving"),
]
for k,v in logic:
    ws.cell(row=r,column=2,value=k).font=Font(name="Arial",bold=True,size=10,color=NAVY)
    ws.cell(row=r,column=3,value=v).font=Font(name="Arial",size=10)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    r+=1

r+=1
hdr(ws,r,2,"SHEET INDEX",bg=NAVY,merge=(1,3)); r+=1
nav=[
    ("CC Baseline","May 2026 cost by cost center — support vs revenue"),
    ("Reallocation Engine","10% saving calc, reinvestment splits, revenue uplift"),
    ("Regional P&L","Monthly Aug/Aug/Sep P&L per cost center/region"),
    ("Consolidated P&L","Group-level monthly P&L under reallocation"),
    ("Margin Analysis","EBITDA, Net Income, cost-to-rev ratios before/after"),
    ("Efficiency Metrics","Revenue productivity, output per $, utilisation"),
    ("Risk Register","Execution risks, capacity constraints, service degradation"),
    ("Optimal Strategy","Recommended allocation with sensitivity analysis"),
    ("Executive Summary","CFO summary: impact, actions, timeline"),
]
for sheet_name,desc in nav:
    c=ws.cell(row=r,column=2,value=sheet_name)
    c.font=Font(name="Arial",bold=True,color=MID,size=10,underline="single")
    c.hyperlink=f"#{sheet_name}!A1"
    ws.cell(row=r,column=3,value=desc).font=Font(name="Arial",size=10,color=DKGRAY)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    r+=1

print("Cover done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 2: CC BASELINE
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["CC Baseline"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 32
ws.column_dimensions['B'].width = 20
ws.column_dimensions['C'].width = 20
ws.column_dimensions['D'].width = 16
ws.column_dimensions['E'].width = 16
ws.column_dimensions['F'].width = 28

hdr(ws,1,1,"COST CENTER BASELINE — MAY 2026 ACTUALS",bg=NAVY,merge=(1,6),sz=13)
ws.row_dimensions[1].height=28

for col,txt,bg in [(1,"Cost Line",NAVY),(2,"Cost Center",DKGRAY),(3,"May 2026 Actual",DKGRAY),
                    (4,"% Total Rev",DKGRAY),(5,"% CC Total",DKGRAY),(6,"Category",DKGRAY)]:
    c=ws.cell(row=2,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center")

r=3
cc_first_rows={}
cc_last_rows={}
all_cost_rows=[]

CC_ORDER = ["IT","HR","OPS","CORP","MKT","NSW","QLD","WA","VIC","SA"]
CC_BG = {"IT":TEAL,"HR":PURPLE,"OPS":ORNG,"CORP":"808080",
          "MKT":"FFD966","NSW":"70AD47","QLD":"70AD47","WA":"70AD47","VIC":"70AD47","SA":"70AD47"}
CC_TYPE = {"IT":"Support Function","HR":"Support Function","OPS":"Support Function","CORP":"Support Function",
           "MKT":"Revenue Region","NSW":"Revenue Region","QLD":"Revenue Region",
           "WA":"Revenue Region","VIC":"Revenue Region","SA":"Revenue Region"}

for cc in CC_ORDER:
    bg=CC_BG[cc]
    sec(ws,r,6,f"{cc} — {CC_TYPE[cc]}",bg=bg,fg=WHITE if bg not in ["FFD966"] else NAVY)
    r+=1
    cc_first_rows[cc]=r
    items=CC_COSTS[cc]
    for line,amt in items.items():
        lbl(ws,r,1,f"  {line}",indent=1)
        ws.cell(row=r,column=2,value=cc).font=Font(name="Arial",size=9,color="808080")
        inp(ws,r,3,amt,DOLR)
        frm(ws,r,4,f"=C{r}/{TOTAL_REV}",PCT,color=DKGRAY)
        frm(ws,r,5,f"=C{r}/C{r+len(items)-all_cost_rows.count(r)}",PCT,color=DKGRAY)  # placeholder
        ws.cell(row=r,column=6,value=CC_TYPE[cc]).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
        all_cost_rows.append(r); r+=1
    cc_last_rows[cc]=r-1
    # CC subtotal
    lbl(ws,r,1,f"TOTAL {cc}",bold=True,bg=LGRAY)
    frm(ws,r,3,f"=SUM(C{cc_first_rows[cc]}:C{cc_last_rows[cc]})",DOLR,True)
    frm(ws,r,4,f"=C{r}/{TOTAL_REV}",PCT,True)
    ws.cell(row=r,column=5,value="100.0%").font=Font(name="Arial",bold=True,size=10)
    ws.cell(row=r,column=6,value=CC_TYPE[cc]).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    fill_row(ws,r,6,LGRAY)
    # Fix % CC total formulas for detail rows
    for detail_r in range(cc_first_rows[cc], cc_last_rows[cc]+1):
        ws.cell(row=detail_r,column=5,value=f"=C{detail_r}/C{r}").number_format=PCT
        ws.cell(row=detail_r,column=5).font=Font(name="Arial",size=10,color=DKGRAY)
    cc_tot_row = r; r+=2

# Grand totals
hdr(ws,r,1,"SUPPORT FUNCTIONS TOTAL",bg="C00000",merge=(1,2))
frm(ws,r,3,f"=SUM(C{cc_first_rows['IT']}:C{cc_last_rows['CORP']})+SUM(C{cc_first_rows['IT']+len(CC_COSTS['IT'])+1}:C{cc_first_rows['IT']+len(CC_COSTS['IT'])+1})",DOLR,True,WHITE)
ws.cell(row=r,column=3).value=f"={SUPPORT_TOTAL}"
ws.cell(row=r,column=3).value=SUPPORT_TOTAL
frm(ws,r,3,f"={SUPPORT_TOTAL}",DOLR,True)
ws.cell(row=r,column=3).value=SUPPORT_TOTAL
inp(ws,r,3,SUPPORT_TOTAL,DOLR,True)
ws.cell(row=r,column=3).font=Font(name="Arial",bold=True,color=WHITE,size=10)
ws.cell(row=r,column=3).fill=PatternFill("solid",start_color="C00000")
frm(ws,r,4,f"={SUPPORT_TOTAL}/{TOTAL_REV}",PCT,True,WHITE)
ws.cell(row=r,column=4).fill=PatternFill("solid",start_color="C00000")
supp_tot_r=r; r+=1

hdr(ws,r,1,"REVENUE REGIONS TOTAL",bg="006400",merge=(1,2))
rev_region_total=sum(CC_TOTALS[cc] for cc in REVENUE_CCS)
inp(ws,r,3,rev_region_total,DOLR,True)
ws.cell(row=r,column=3).font=Font(name="Arial",bold=True,color=WHITE)
ws.cell(row=r,column=3).fill=PatternFill("solid",start_color="006400")
frm(ws,r,4,f"={rev_region_total}/{TOTAL_REV}",PCT,True,WHITE)
ws.cell(row=r,column=4).fill=PatternFill("solid",start_color="006400")
rev_tot_r=r; r+=1

hdr(ws,r,1,"TOTAL OPERATING EXPENSES",bg=NAVY,merge=(1,2))
inp(ws,r,3,TOTAL_EXP,DOLR,True)
ws.cell(row=r,column=3).font=Font(name="Arial",bold=True,color=WHITE)
ws.cell(row=r,column=3).fill=PatternFill("solid",start_color=NAVY)
frm(ws,r,4,f"={TOTAL_EXP}/{TOTAL_REV}",PCT,True,WHITE)
ws.cell(row=r,column=4).fill=PatternFill("solid",start_color=NAVY)
r+=2

# Summary table
hdr(ws,r,1,"COST CENTER SUMMARY",bg=MID,merge=(1,6)); r+=1
for col,txt in [(1,"Cost Center"),(2,"Type"),(3,"Monthly Cost"),(4,"% Revenue"),(5,"10% Saving"),(6,"Role")]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=MID)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center")
r+=1
for cc in CC_ORDER:
    total=CC_TOTALS[cc]
    is_supp=cc in SUPPORT_CCS
    lbl(ws,r,1,cc,bold=True,color="C00000" if is_supp else "006400")
    ws.cell(row=r,column=2,value="Support" if is_supp else "Revenue Region").font=Font(name="Arial",size=10)
    inp(ws,r,3,total,DOLR)
    frm(ws,r,4,f"={total}/{TOTAL_REV}",PCT)
    if is_supp:
        inp(ws,r,5,total*0.10,DOLR)
        ws.cell(row=r,column=5).font=Font(name="Arial",color="006400",size=10)
    else:
        ws.cell(row=r,column=5,value="—").font=Font(name="Arial",size=10,color=DKGRAY)
    role={"IT":"Technology infrastructure","HR":"People & talent","OPS":"Logistics & ops",
          "CORP":"Corporate overhead","MKT":"Marketing & sales","NSW":"Revenue gen NSW",
          "QLD":"Revenue gen QLD","WA":"Revenue gen WA","VIC":"Revenue gen VIC","SA":"Revenue gen SA"}
    ws.cell(row=r,column=6,value=role.get(cc,"")).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    r+=1

lbl(ws,r,1,"TOTAL SUPPORT SAVING (10%)",bold=True,bg=GREEN)
inp(ws,r,3,SAVINGS,DOLR,True)
ws.cell(row=r,column=3).font=Font(name="Arial",bold=True,color="006400")
fill_row(ws,r,6,GREEN)
r+=1

print("CC Baseline done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 3: REALLOCATION ENGINE
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Reallocation Engine"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col,w in [('B',18),('C',18),('D',18),('E',18),('F',20)]:
    ws.column_dimensions[col].width = w

hdr(ws,1,1,"REALLOCATION ENGINE — ASSUMPTIONS & CALCULATION",bg=ORNG,merge=(1,6),sz=13)
ws.row_dimensions[1].height=28

# Ramp schedule: reductions phase in over Q3
RAMP = {"Aug":0.80,"Aug":0.90,"Sep":1.00}

r=3
r=sec(ws,r,6,"STEP 1: SUPPORT FUNCTION REDUCTION (10%)",NAVY)
for col,txt,bg in [(1,"Cost Center",NAVY),(2,"Monthly Baseline",MID),(3,"10% Saving",MID),
                    (4,"Aug Saving (80%)",ORNG),(5,"Aug Saving (90%)",ORNG),(6,"Sep Saving (100%)",ORNG)]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center")
r+=1

supp_rows={}
for cc in SUPPORT_CCS:
    total=CC_TOTALS[cc]
    saving=total*0.10
    lbl(ws,r,1,cc,bold=True,color="C00000")
    inp(ws,r,2,total,DOLR)
    frm(ws,r,3,f"=B{r}*0.10",DOLR,color="006400")
    frm(ws,r,4,f"=C{r}*0.80",DOLR,color=ORNG)
    frm(ws,r,5,f"=C{r}*0.90",DOLR,color=ORNG)
    frm(ws,r,6,f"=C{r}*1.00",DOLR,color=ORNG)
    supp_rows[cc]=r; r+=1

lbl(ws,r,1,"TOTAL SUPPORT SAVINGS",bold=True,bg=GREEN)
fill_row(ws,r,6,GREEN)
for col in [2,3,4,5,6]:
    frm(ws,r,col,f"=SUM({get_column_letter(col)}{r-4}:{get_column_letter(col)}{r-1})",
        DOLR,True,color="006400")
total_saving_r=r; r+=2

r=sec(ws,r,6,"STEP 2: PROFIT RETENTION vs REINVESTMENT SPLIT",MID)
lbl(ws,r,1,"Profit Retention %",bold=False); inp(ws,r,2,0.40,PCT); r+=1; ret_r=r-1
lbl(ws,r,1,"Reinvestment %",bold=False); frm(ws,r,2,f"=1-B{ret_r}",PCT); r+=1; reinv_pct_r=r-1

lbl(ws,r,1,"Profit Bottom-Line Boost — Aug",bold=True,bg=GREEN)
frm(ws,r,4,f"=D{total_saving_r}*B{ret_r}",DOLR,True,color="006400"); 
frm(ws,r,5,f"=E{total_saving_r}*B{ret_r}",DOLR,True,color="006400")
frm(ws,r,6,f"=F{total_saving_r}*B{ret_r}",DOLR,True,color="006400")
fill_row(ws,r,6,GREEN); profit_direct_r=r; r+=1

lbl(ws,r,1,"Available for Reinvestment — Aug/Aug/Sep",bold=True,bg=AMBER)
frm(ws,r,4,f"=D{total_saving_r}*B{reinv_pct_r}",DOLR,True)
frm(ws,r,5,f"=E{total_saving_r}*B{reinv_pct_r}",DOLR,True)
frm(ws,r,6,f"=F{total_saving_r}*B{reinv_pct_r}",DOLR,True)
fill_row(ws,r,6,AMBER); reinvest_pool_r=r; r+=2

r=sec(ws,r,6,"STEP 3: REINVESTMENT ALLOCATION TO REVENUE REGIONS",MID)
for col,txt,bg in [(1,"Region",MID),(2,"Alloc %",MID),(3,"Rev Multiplier",MID),
                    (4,"Aug Reinvest $",ORNG),(5,"Aug Reinvest $",ORNG),(6,"Sep Reinvest $",ORNG)]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center")
r+=1

region_rows={}
for cc,alloc_pct in REINVEST_ALLOC.items():
    mult=REV_MULT[cc]
    lbl(ws,r,1,cc,bold=True,color="006400")
    inp(ws,r,2,alloc_pct,PCT)
    inp(ws,r,3,mult,NUM2)
    for col,src_col in [(4,"D"),(5,"E"),(6,"F")]:
        frm(ws,r,col,f"={src_col}{reinvest_pool_r}*B{r}",DOLR)
    region_rows[cc]=r; r+=1

lbl(ws,r,1,"CHECK: Total Alloc %",bold=False,color=DKGRAY)
frm(ws,r,2,f"=SUM(B{r-4}:B{r-1})",PCT,color=DKGRAY)
ws.cell(row=r,column=2).font=Font(name="Arial",size=10,color=DKGRAY,italic=True); r+=2

r=sec(ws,r,6,"STEP 4: REVENUE UPLIFT FROM REINVESTMENT",MID)
for col,txt in [(1,"Region"),(2,"Aug Invest $"),(3,"Aug Rev Uplift"),(4,"Aug Invest $"),(5,"Aug Rev Uplift"),(6,"Sep Invest $")]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=MID)
r+=1

rev_uplift_rows={}
for cc in REINVEST_ALLOC:
    mult=REV_MULT[cc]
    rr=region_rows[cc]
    lbl(ws,r,1,cc,bold=True,color="006400")
    frm(ws,r,2,f"=D{rr}",DOLR)
    frm(ws,r,3,f"=D{rr}*{mult}",DOLR,True,color="006400")  # Aug revenue uplift
    frm(ws,r,4,f"=E{rr}",DOLR)
    frm(ws,r,5,f"=E{rr}*{mult}",DOLR,True,color="006400")  # Aug
    frm(ws,r,6,f"=F{rr}*{mult}",DOLR,True,color="006400")  # Sep
    rev_uplift_rows[cc]=r; r+=1

lbl(ws,r,1,"TOTAL REVENUE UPLIFT",bold=True,bg=GREEN); fill_row(ws,r,6,GREEN)
for col in [3,5,6]:
    frm(ws,r,col,f"=SUM({get_column_letter(col)}{r-4}:{get_column_letter(col)}{r-1})",
        DOLR,True,color="006400")
tot_rev_uplift_r=r; r+=2

# Sep col also needs col 5 covered — add col 5 sum for Sep
# Actually let me add total reinvest too
r=sec(ws,r,6,"STEP 5: NET INCREMENTAL IMPACT SUMMARY",NAVY)
lbl(ws,r,1,"Month"); ws.cell(row=r,column=2,value="Aug 2026"); ws.cell(row=r,column=3,value="Aug 2026"); ws.cell(row=r,column=4,value="Sep 2026"); ws.cell(row=r,column=5,value="Q3 Total")
for c in range(1,6): ws.cell(row=r,column=c).font=Font(name="Arial",bold=True,size=10,color=WHITE); ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=NAVY)
r+=1

impact_items=[
    ("Total Cost Saving","=D{ts}","=E{ts}","=F{ts}",total_saving_r),
    ("Profit Retention (direct EBITDA)","=D{p}","=E{p}","=F{p}",profit_direct_r),
    ("Reinvested in Revenue Regions","=D{ri}","=E{ri}","=F{ri}",reinvest_pool_r),
    ("Revenue Uplift from Reinvestment","=C{ru}","=E{ru}","=F{ru}",tot_rev_uplift_r),  # Note: col C=Aug, E=Aug, F=Sep for rev uplift
]
for label,jul_f,aug_f,sep_f,src_r in impact_items:
    lbl(ws,r,1,label,bold="EBITDA" in label or "Revenue Uplift" in label)
    frm(ws,r,2,jul_f.replace("{ts}",str(src_r)).replace("{p}",str(src_r)).replace("{ri}",str(src_r)).replace("{ru}",str(src_r)),DOLR)
    frm(ws,r,3,aug_f.replace("{ts}",str(src_r)).replace("{p}",str(src_r)).replace("{ri}",str(src_r)).replace("{ru}",str(src_r)),DOLR)
    frm(ws,r,4,sep_f.replace("{ts}",str(src_r)).replace("{p}",str(src_r)).replace("{ri}",str(src_r)).replace("{ru}",str(src_r)),DOLR)
    frm(ws,r,5,f"=SUM(B{r}:D{r})",DOLR,True)
    r+=1

# Store key reference rows for other sheets
RE_ROWS={
    "total_saving":total_saving_r,"profit_direct":profit_direct_r,
    "reinvest_pool":reinvest_pool_r,"tot_rev_uplift":tot_rev_uplift_r,
    "region_rows":region_rows,"rev_uplift_rows":rev_uplift_rows
}

print(f"Reallocation Engine done. Key rows: {list(RE_ROWS.keys())}")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 4: REGIONAL P&L
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Regional P&L"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 26
for col in ['B','C','D','E','F','G','H','I','J']:
    ws.column_dimensions[col].width = 16

hdr(ws,1,1,"REGIONAL P&L — MONTHLY FORECAST AUG/AUG/SEP 2026 (POST-REALLOCATION)",bg="70AD47",merge=(1,10),sz=12)
ws.row_dimensions[1].height=28

# Revenue by region (May actuals)
REV_REGIONS={"NSW":12569315,"QLD":7298312,"WA":6892850,"VIC":8920159,"SA":4865541}
WORKFORCE={"NSW":3571090,"QLD":3582740,"WA":2324260,"VIC":3401870,"SA":1447230}
# MKT drives top-line but doesn't have a standalone revenue center
# Revenue seasonality (same as driver model)
SEA={"Aug":0.95,"Aug":1.00,"Sep":1.08}
# Revenue growth (base case 8.8%)
REV_GR=1.088

# Reinvestment per region per month
REINV_M={"QLD":REINVEST*0.30,"NSW":REINVEST*0.30,"WA":REINVEST*0.25,"MKT":REINVEST*0.15}

r=3
regions=["NSW","QLD","WA","VIC","SA"]
months=["Aug","Aug","Sep"]

# Header row
hdr(ws,r,1,"Region / Metric","70AD47")
col=2
region_col_map={}
for reg in regions:
    ws.merge_cells(start_row=r,start_column=col,end_row=r,end_column=col+1)
    c=ws.cell(row=r,column=col,value=reg)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=11)
    c.fill=PatternFill("solid",start_color="70AD47")
    c.alignment=Alignment(horizontal="center")
    region_col_map[reg]=col
    col+=2
r+=1

# Sub-headers: Before / After for each region
hdr(ws,r,1,"Metric","70AD47")
for reg in regions:
    col=region_col_map[reg]
    for i,txt in enumerate(["Before","After Realloc"]):
        c=ws.cell(row=r,column=col+i,value=txt)
        c.font=Font(name="Arial",bold=True,color=WHITE,size=9)
        c.fill=PatternFill("solid",start_color=MID if i==0 else "006400")
        c.alignment=Alignment(horizontal="center")
r+=1

# For each month
MONTH_BG={"Aug":LGRAY,"Aug":LIGHT,"Sep":GREEN}
for month in months:
    seas=SEA[month]
    ramp={"Aug":0.80,"Aug":0.90,"Sep":1.00}[month]
    
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10)
    c=ws.cell(row=r,column=1,value=f"── {month} 2026 ──")
    c.font=Font(name="Arial",bold=True,color=NAVY,size=11)
    c.fill=PatternFill("solid",start_color=MONTH_BG[month])
    for c2 in range(1,11): ws.cell(row=r,column=c2).fill=PatternFill("solid",start_color=MONTH_BG[month])
    r+=1
    
    row_labels=["Revenue","Direct Costs","Reinvestment (in)","Revenue Uplift","Gross Profit","Gross Margin %","Cost-to-Rev Before","Cost-to-Rev After"]
    for label in row_labels:
        lbl(ws,r,1,label,bold=label in ("Gross Profit","Gross Margin %"))
        for reg in regions:
            col=region_col_map[reg]
            base_rev=REV_REGIONS[reg]*REV_GR*seas
            work=WORKFORCE[reg]
            reinv=REINV_M.get(reg,0)*ramp if reg in REINV_M else 0
            rev_mult=REV_MULT.get(reg,0)
            rev_uplift=reinv*rev_mult
            
            if label=="Revenue":
                inp(ws,r,col,base_rev,DOLR)
                frm(ws,r,col+1,f"={get_column_letter(col)}{r}+{rev_uplift}",DOLR,color="006400")
            elif label=="Direct Costs":
                inp(ws,r,col,work,DOLR)
                inp(ws,r,col+1,work,DOLR)  # workforce unchanged
            elif label=="Reinvestment (in)":
                inp(ws,r,col,0,DOLR)
                inp(ws,r,col+1,reinv,DOLR)
                ws.cell(row=r,column=col+1).font=Font(name="Arial",color="006400",size=10)
            elif label=="Revenue Uplift":
                inp(ws,r,col,0,DOLR)
                inp(ws,r,col+1,rev_uplift,DOLR)
                ws.cell(row=r,column=col+1).font=Font(name="Arial",color="006400",bold=True,size=10)
            elif label=="Gross Profit":
                frm(ws,r,col,f"={get_column_letter(col)}{r-4}-{get_column_letter(col)}{r-3}",DOLR,True)
                frm(ws,r,col+1,f"={get_column_letter(col+1)}{r-4}+{get_column_letter(col+1)}{r-2}-{get_column_letter(col+1)}{r-3}",DOLR,True,color="006400")
                ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=LGRAY)
                ws.cell(row=r,column=col+1).fill=PatternFill("solid",start_color=GREEN)
            elif label=="Gross Margin %":
                gp_r=r-1
                rev_r=r-5
                frm(ws,r,col,f"={get_column_letter(col)}{gp_r}/{get_column_letter(col)}{rev_r}",PCT,True)
                frm(ws,r,col+1,f"={get_column_letter(col+1)}{gp_r}/{get_column_letter(col+1)}{rev_r}",PCT,True,color="006400")
            elif label=="Cost-to-Rev Before":
                frm(ws,r,col,f"={get_column_letter(col)}{r-4}/{get_column_letter(col)}{r-6}",PCT)
                frm(ws,r,col+1,f"={get_column_letter(col+1)}{r-4}/{get_column_letter(col+1)}{r-6}",PCT,color=DKGRAY)
            elif label=="Cost-to-Rev After":
                frm(ws,r,col,f"=({get_column_letter(col)}{r-5}-0)/{get_column_letter(col)}{r-7}",PCT)
                frm(ws,r,col+1,f"=({get_column_letter(col+1)}{r-5}-{get_column_letter(col+1)}{r-5}+{get_column_letter(col+1)}{r-5})/{get_column_letter(col+1)}{r-7}",PCT,color="006400")
        r+=1
    r+=1

print("Regional P&L done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 5: CONSOLIDATED P&L
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Consolidated P&L"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 38
for col,w in [('B',18),('C',18),('D',18),('E',18),('F',18),('G',18),('H',18)]:
    ws.column_dimensions[col].width = w

hdr(ws,1,1,"CONSOLIDATED P&L — BEFORE vs AFTER COST REALLOCATION",bg=MID,merge=(1,8),sz=13)
ws.row_dimensions[1].height=28

# Columns: May Actual | Aug Before | Aug After | Aug Before | Aug After | Sep Before | Sep After | Q3 After
for col,txt,bg in [(1,"Line Item",NAVY),(2,"May 2026\nActual",DKGRAY),
                    (3,"Aug Before",MID),(4,"Aug After\n(Realloc)","006400"),
                    (5,"Aug Before",MID),(6,"Aug After\n(Realloc)","006400"),
                    (7,"Sep Before",MID),(8,"Sep After\n(Realloc)","006400")]:
    c=ws.cell(row=2,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center",wrap_text=True)
ws.row_dimensions[2].height=28

# Base P&L figures — growing with seasonality and 8.8% YoY
def base_rev(seas): return TOTAL_REV * REV_GR * seas
def base_exp(seas): return TOTAL_EXP * 1.05 * seas  # costs grow ~5%

# Additional revenue from reinvestment (all regions combined)
TOTAL_REINV=sum(REINV_M.values())
def rev_uplift_total(ramp): 
    return sum(REINV_M[cc]*ramp*REV_MULT[cc] for cc in REINV_M)

SAVINGS_BY_MONTH={"Aug":SAVINGS*0.80,"Aug":SAVINGS*0.90,"Sep":SAVINGS*1.00}

# We also add support function reduction to expenses
def adj_exp(seas,ramp):
    base=base_exp(seas)
    return base - SAVINGS * ramp  # costs reduced by savings

r=3
def pl_row(ws,r,label,may_v,jul_b,jul_a,aug_b,aug_a,sep_b,sep_a,fmt=DOLR,bold=False,bg=None):
    lbl(ws,r,1,label,bold=bold,bg=bg)
    for col,val in [(2,may_v),(3,jul_b),(4,jul_a),(5,aug_b),(6,aug_a),(7,sep_b),(8,sep_a)]:
        if isinstance(val,str):
            frm(ws,r,col,val,fmt,bold,color=WHITE if bg in [NAVY,MID,"C00000","006400"] else "000000")
        else:
            inp(ws,r,col,round(val),fmt,bold)
            if bold and bg:
                ws.cell(row=r,column=col).font=Font(name="Arial",bold=True,color=WHITE,size=10)
        if bg: ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=bg)
    return r

rev_rows={}; exp_rows={}; ni_rows={}; ebitda_rows={}

# REVENUE
sec(ws,r,8,"REVENUE",NAVY); r+=1
pl_row(ws,r,"Total Revenue",TOTAL_REV,
       base_rev(0.95),base_rev(0.95)+rev_uplift_total(0.80),
       base_rev(1.00),base_rev(1.00)+rev_uplift_total(0.90),
       base_rev(1.08),base_rev(1.08)+rev_uplift_total(1.00),
       DOLR,True,LGRAY)
rev_r=r; r+=2

lbl(ws,r,1,"  Incremental Revenue (from reinvestment)",indent=1)
for col,val in [(4,rev_uplift_total(0.80)),(6,rev_uplift_total(0.90)),(8,rev_uplift_total(1.00))]:
    inp(ws,r,col,val,DOLR)
    ws.cell(row=r,column=col).font=Font(name="Arial",color="006400",size=10)
lbl(ws,r,3,"—",color=DKGRAY); lbl(ws,r,5,"—",color=DKGRAY); lbl(ws,r,7,"—",color=DKGRAY)
r+=1

# OPERATING EXPENSES
sec(ws,r,8,"OPERATING EXPENSES (by category)",MID); r+=1
expense_map=[
    ("IT (Software, Cloud, IT Support)",CC_TOTALS["IT"],0.10),
    ("HR (Personnel, Development)",CC_TOTALS["HR"],0.10),
    ("Operations (Telecom, Freight, Insurance)",CC_TOTALS["OPS"],0.10),
    ("Corporate (Facilities, External, Finance)",CC_TOTALS["CORP"],0.10),
    ("Marketing & Sales",CC_TOTALS["MKT"],0.0),
    ("Revenue Region Workforce (NSW,QLD,WA,VIC,SA)",rev_region_total,0.0),
]
exp_line_rows=[]
for label,may_amt,reduction in expense_map:
    lbl(ws,r,1,f"  {label}",indent=1)
    inp(ws,r,2,may_amt,DOLR)
    for month,ramp,cols in [("Aug",0.80,(3,4)),("Aug",0.90,(5,6)),("Sep",1.00,(7,8))]:
        seas={"Aug":0.95,"Aug":1.00,"Sep":1.08}[month]
        before=may_amt*1.05*seas
        after=before*(1-reduction*ramp) if reduction>0 else before
        inp(ws,r,cols[0],before,DOLR)
        inp(ws,r,cols[1],after,DOLR)
        if reduction>0:
            ws.cell(row=r,column=cols[1]).font=Font(name="Arial",color="006400",size=10)
    exp_line_rows.append(r); r+=1

# Reinvestment row (cost that goes back in)
lbl(ws,r,1,"  Reinvestment into Revenue Regions",indent=1)
inp(ws,r,2,0,DOLR)
for col,ramp in [(4,0.80),(6,0.90),(8,1.00)]:
    inp(ws,r,col,TOTAL_REINV*ramp,DOLR)
    ws.cell(row=r,column=col).font=Font(name="Arial",color=ORNG,size=10)
exp_line_rows.append(r); r+=1

lbl(ws,r,1,"TOTAL OPERATING EXPENSES",bold=True,bg=LGRAY); fill_row(ws,r,8,LGRAY)
for col in [2,3,4,5,6,7,8]:
    frm(ws,r,col,f"=SUM({get_column_letter(col)}{exp_line_rows[0]}:{get_column_letter(col)}{exp_line_rows[-1]})",DOLR,True)
tot_exp_r=r; r+=2

# EBITDA
lbl(ws,r,1,"EBITDA",bold=True,bg=MID); fill_row(ws,r,8,MID)
for col in [2,3,4,5,6,7,8]:
    frm(ws,r,col,f"={get_column_letter(col)}{rev_r}-{get_column_letter(col)}{tot_exp_r}+{NONCASH}",DOLR,True,WHITE)
    ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=MID)
ebitda_r=r; r+=1

# D&A
lbl(ws,r,1,"  Less: D&A & Amortization",indent=1)
for col in [2,3,4,5,6,7,8]: inp(ws,r,col,NONCASH,DOLR)
da_r=r; r+=1

# EBIT
lbl(ws,r,1,"EBIT",bold=True,bg=LGRAY); fill_row(ws,r,8,LGRAY)
for col in [2,3,4,5,6,7,8]:
    frm(ws,r,col,f"={get_column_letter(col)}{ebitda_r}-{get_column_letter(col)}{da_r}",DOLR,True)
ebit_r=r; r+=1

# Tax
lbl(ws,r,1,"  Income Tax (30%)",indent=1)
for col in [2,3,4,5,6,7,8]:
    frm(ws,r,col,f"={get_column_letter(col)}{ebit_r}*0.30",DOLR)
tax_r=r; r+=1

# Net Income
lbl(ws,r,1,"NET INCOME",bold=True,bg=NAVY); fill_row(ws,r,8,NAVY)
for col in [2,3,4,5,6,7,8]:
    frm(ws,r,col,f"={get_column_letter(col)}{ebit_r}-{get_column_letter(col)}{tax_r}",DOLR,True,WHITE)
    ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=NAVY)
ni_r=r; r+=2

# MARGIN METRICS
sec(ws,r,8,"MARGIN METRICS",DKGRAY); r+=1
margins=[
    ("EBITDA Margin %",ebitda_r),("Net Income Margin %",ni_r),
]
for label,src_r in margins:
    lbl(ws,r,1,label,bold=True)
    for col in [2,3,4,5,6,7,8]:
        frm(ws,r,col,f"={get_column_letter(col)}{src_r}/{get_column_letter(col)}{rev_r}",PCT,True,
            color="006400" if col%2==0 else "000000")
    r+=1

# DELTA ROW
lbl(ws,r,1,"Net Income UPLIFT (After vs Before)",bold=True,bg=GREEN); fill_row(ws,r,8,GREEN)
for col,before_col in [(4,3),(6,5),(8,7)]:
    frm(ws,r,col,f"={get_column_letter(col)}{ni_r}-{get_column_letter(before_col)}{ni_r}",DOLR,True,color="006400")
    ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=GREEN)
r+=1

lbl(ws,r,1,"EBITDA UPLIFT (After vs Before)",bold=True,bg=GREEN); fill_row(ws,r,8,GREEN)
for col,before_col in [(4,3),(6,5),(8,7)]:
    frm(ws,r,col,f"={get_column_letter(col)}{ebitda_r}-{get_column_letter(before_col)}{ebitda_r}",DOLR,True,color="006400")
    ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=GREEN)
r+=1

# Store rows
CONS_ROWS={"rev":rev_r,"ebitda":ebitda_r,"ni":ni_r,"tot_exp":tot_exp_r}
print(f"Consolidated P&L done. rev={rev_r}, ebitda={ebitda_r}, ni={ni_r}")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 6: MARGIN ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Margin Analysis"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col in ['B','C','D','E','F','G']:
    ws.column_dimensions[col].width = 18

hdr(ws,1,1,"MARGIN ANALYSIS — BEFORE vs AFTER COST REALLOCATION",bg=PURPLE,merge=(1,7),sz=13)
ws.row_dimensions[1].height=28

for col,txt,bg in [(1,"Metric",PURPLE),(2,"May 2026 Actual",DKGRAY),
                    (3,"Q3 Before",MID),(4,"Q3 After Realloc","006400"),
                    (5,"Δ (After-Before)",ORNG),(6,"Δ %",ORNG),(7,"Assessment",DKGRAY)]:
    c=ws.cell(row=2,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center")

r=3
# Q3 totals (sum of Aug+Aug+Sep)
q3_rev_before = base_rev(0.95)+base_rev(1.00)+base_rev(1.08)
q3_rev_after  = q3_rev_before + rev_uplift_total(0.80)+rev_uplift_total(0.90)+rev_uplift_total(1.00)
q3_exp_before = adj_exp(0.95,0)+adj_exp(1.00,0)+adj_exp(1.08,0)  # no reduction
q3_exp_after  = adj_exp(0.95,0.80)+adj_exp(1.00,0.90)+adj_exp(1.08,1.00)+TOTAL_REINV*(0.80+0.90+1.00)
q3_ebitda_before = q3_rev_before - q3_exp_before + NONCASH*3
q3_ebitda_after  = q3_rev_after  - q3_exp_after  + NONCASH*3
q3_ni_before = (q3_ebitda_before - NONCASH*3) * 0.70
q3_ni_after  = (q3_ebitda_after  - NONCASH*3) * 0.70

def mar_row(ws,r,label,may_v,q3_b,q3_a,fmt=DOLR,bold=False,assess="",bg=None):
    lbl(ws,r,1,label,bold=bold,bg=bg)
    inp(ws,r,2,may_v,fmt,bold)
    inp(ws,r,3,q3_b,fmt,bold)
    inp(ws,r,4,q3_a,fmt,bold)
    ws.cell(row=r,column=4).font=Font(name="Arial",bold=bold,color="006400",size=10)
    delta=q3_a-q3_b
    inp(ws,r,5,delta,fmt,bold)
    ws.cell(row=r,column=5).font=Font(name="Arial",bold=bold,
        color="006400" if delta>0 else "C00000",size=10)
    pct_d=delta/abs(q3_b) if q3_b!=0 else 0
    inp(ws,r,6,pct_d,PCT,bold)
    ws.cell(row=r,column=6).font=Font(name="Arial",bold=bold,
        color="006400" if pct_d>0 else "C00000",size=10)
    ws.cell(row=r,column=7,value=assess).font=Font(name="Arial",size=9,italic=True,
        color="006400" if "↑" in assess or "Expansion" in assess else 
               "C00000" if "↓" in assess or "Risk" in assess else DKGRAY)
    if bg:
        for c in range(1,8): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=bg)
    return r

sec(ws,r,7,"REVENUE & COST",NAVY); r+=1
mar_row(ws,r,"Total Revenue (Q3)",TOTAL_REV*3,q3_rev_before,q3_rev_after,DOLR,True,
        "↑ Revenue uplift from reinvestment"); r+=1
mar_row(ws,r,"Total Operating Expenses",TOTAL_EXP*3,q3_exp_before,q3_exp_after,DOLR,False,
        "↓ Cost base reduced by support efficiency"); r+=1
mar_row(ws,r,"Net Cost Saving (support 10% reduction)",0,SAVINGS*2.70,SAVINGS*2.70,DOLR,False,
        f"$2.77M × (0.8+0.9+1.0) = ${SAVINGS*2.7:,.0f}"); r+=2

sec(ws,r,7,"PROFITABILITY METRICS",MID); r+=1
mar_row(ws,r,"EBITDA",EBITDA,q3_ebitda_before,q3_ebitda_after,DOLR,True,
        "↑ Margin expansion expected"); r+=1
mar_row(ws,r,"EBITDA Margin %",0.278,q3_ebitda_before/q3_rev_before,q3_ebitda_after/q3_rev_after,PCT,True,
        f"Expansion: {(q3_ebitda_after/q3_rev_after-q3_ebitda_before/q3_rev_before)*100:.1f}pp",bg=GREEN); r+=1
mar_row(ws,r,"Net Income",NET_INCOME,q3_ni_before,q3_ni_after,DOLR,True,
        "↑ Bottom line improvement"); r+=1
mar_row(ws,r,"Net Income Margin %",0.224,q3_ni_before/q3_rev_before,q3_ni_after/q3_rev_after,PCT,True,
        f"Expansion: {(q3_ni_after/q3_rev_after-q3_ni_before/q3_rev_before)*100:.1f}pp",bg=GREEN); r+=2

sec(ws,r,7,"COST-TO-REVENUE RATIOS BY COST CENTER",PURPLE); r+=1
for col,txt in [(1,"Cost Center"),(2,"May Cost"),(3,"May C/R Ratio"),(4,"Q3 After C/R"),(5,"Improvement"),(6,"Δ pp"),(7,"Comment")]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=PURPLE)
r+=1

for cc in CC_ORDER:
    total=CC_TOTALS[cc]
    rev_base=REV_REGIONS.get(cc,TOTAL_REV if cc in SUPPORT_CCS else 0)
    if cc in SUPPORT_CCS:
        rev_denom=TOTAL_REV
        after_cost=total*0.90
    else:
        rev_denom=REV_REGIONS.get(cc,1)
        after_cost=total  # unchanged
    
    before_ratio=total/TOTAL_REV
    after_ratio=after_cost/TOTAL_REV
    delta_pp=(after_ratio-before_ratio)*100
    
    lbl(ws,r,1,cc,bold=True,color="C00000" if cc in SUPPORT_CCS else "006400")
    inp(ws,r,2,total,DOLR)
    inp(ws,r,3,before_ratio,PCT)
    inp(ws,r,4,after_ratio,PCT)
    ws.cell(row=r,column=4).font=Font(name="Arial",color="006400",size=10)
    inp(ws,r,5,-total*0.10 if cc in SUPPORT_CCS else 0,DOLR)
    ws.cell(row=r,column=5).font=Font(name="Arial",color="006400" if cc in SUPPORT_CCS else DKGRAY,size=10)
    inp(ws,r,6,delta_pp/100,PCT)
    ws.cell(row=r,column=6).font=Font(name="Arial",color="006400" if delta_pp<0 else DKGRAY,size=10)
    comment=f"10% efficiency saving; reinvest {int(REINVEST_ALLOC.get(cc,0)*100)}% to regions" if cc in SUPPORT_CCS else f"Receives {int(REINVEST_ALLOC.get(cc,0)*100)}% reinvest; {REV_MULT.get(cc,'-')}x rev mult"
    ws.cell(row=r,column=7,value=comment).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    r+=1

print("Margin Analysis done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 7: EFFICIENCY METRICS
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Efficiency Metrics"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
for col in ['B','C','D','E','F']:
    ws.column_dimensions[col].width = 20

hdr(ws,1,1,"EFFICIENCY METRICS — REVENUE PRODUCTIVITY & COST PERFORMANCE",bg=TEAL,merge=(1,6),sz=13)
ws.row_dimensions[1].height=28

for col,txt,bg in [(1,"Metric",TEAL),(2,"May 2026",DKGRAY),(3,"Q3 Before",MID),
                    (4,"Q3 After","006400"),(5,"Change",ORNG),(6,"Benchmark / Note",DKGRAY)]:
    c=ws.cell(row=2,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center")

r=3
eff_data=[
    ("REVENUE PRODUCTIVITY","","","","",""),
    ("Revenue per employee (total workforce)",
     TOTAL_REV/sum(WORKFORCE.values())*1000,  # annualised proxy
     q3_rev_before/sum(WORKFORCE.values()),
     q3_rev_after/sum(WORKFORCE.values()),DOLR,"Higher = better utilisation of headcount"),
    ("Revenue per $ of total cost",
     TOTAL_REV/TOTAL_EXP,q3_rev_before/q3_exp_before,q3_rev_after/q3_exp_after,NUM2,"Target > 1.3x"),
    ("Support cost per $ revenue generated",
     SUPPORT_TOTAL/TOTAL_REV,SUPPORT_TOTAL*1.05*2.9/q3_rev_before,
     SUPPORT_TOTAL*0.90*1.05*2.9/q3_rev_after,NUM2,"Lower = more efficient support base"),
    ("","","","","",""),
    ("SUPPORT FUNCTION EFFICIENCY","","","","",""),
    ("IT cost as % revenue",CC_TOTALS["IT"]/TOTAL_REV,CC_TOTALS["IT"]*1.05*2.9/q3_rev_before,
     CC_TOTALS["IT"]*0.90*1.05*2.9/q3_rev_after,PCT,"Industry benchmark: 3-5%"),
    ("HR cost as % revenue",CC_TOTALS["HR"]/TOTAL_REV,CC_TOTALS["HR"]*1.05*2.9/q3_rev_before,
     CC_TOTALS["HR"]*0.90*1.05*2.9/q3_rev_after,PCT,"Industry benchmark: 10-15%"),
    ("OPS cost as % revenue",CC_TOTALS["OPS"]/TOTAL_REV,CC_TOTALS["OPS"]*1.05*2.9/q3_rev_before,
     CC_TOTALS["OPS"]*0.90*1.05*2.9/q3_rev_after,PCT,"Industry benchmark: 6-10%"),
    ("CORP overhead as % revenue",CC_TOTALS["CORP"]/TOTAL_REV,CC_TOTALS["CORP"]*1.05*2.9/q3_rev_before,
     CC_TOTALS["CORP"]*0.90*1.05*2.9/q3_rev_after,PCT,"Industry benchmark: 8-12%"),
    ("","","","","",""),
    ("REVENUE REGION METRICS","","","","",""),
    ("QLD revenue productivity (Rev/$cost)",
     REV_REGIONS["QLD"]/CC_TOTALS["QLD"],
     (REV_REGIONS["QLD"]*REV_GR*0.95+rev_uplift_total(0.80)*0.30)/CC_TOTALS["QLD"],
     (REV_REGIONS["QLD"]*REV_GR*0.95+rev_uplift_total(0.80)*0.30)/(CC_TOTALS["QLD"]+REINV_M["QLD"]),
     NUM2,"Higher = better ROI on QLD investment"),
    ("NSW revenue productivity (Rev/$cost)",
     REV_REGIONS["NSW"]/CC_TOTALS["NSW"],
     (REV_REGIONS["NSW"]*REV_GR*0.95)/(CC_TOTALS["NSW"]),
     (REV_REGIONS["NSW"]*REV_GR*0.95+rev_uplift_total(0.80)*0.30)/(CC_TOTALS["NSW"]+REINV_M["NSW"]),
     NUM2,"Higher = better ROI on NSW investment"),
    ("WA revenue productivity (Rev/$cost)",
     REV_REGIONS["WA"]/CC_TOTALS["WA"],
     (REV_REGIONS["WA"]*REV_GR*0.95)/CC_TOTALS["WA"],
     (REV_REGIONS["WA"]*REV_GR*0.95+rev_uplift_total(0.80)*0.25)/(CC_TOTALS["WA"]+REINV_M["WA"]),
     NUM2,"WA highest multiplier (3.2x); strong ROI"),
    ("","","","","",""),
    ("OPERATIONAL EFFICIENCY","","","","",""),
    ("Operating leverage (Rev growth / Cost growth)",1.68,q3_rev_before/TOTAL_REV/3,
     q3_rev_after/TOTAL_REV/3/(q3_exp_after/TOTAL_EXP/3),NUM2,"Target > 1.5 — achieved indicates scalability"),
    ("EBITDA per $M revenue",EBITDA/TOTAL_REV*1000000,
     q3_ebitda_before/q3_rev_before*1000000,q3_ebitda_after/q3_rev_after*1000000,
     DOLR,"Trend upward post-reallocation confirms efficiency"),
    ("Fixed cost coverage ratio",TOTAL_REV/SUPPORT_TOTAL,
     q3_rev_before/3/(SUPPORT_TOTAL*0.95),q3_rev_after/3/(SUPPORT_TOTAL*0.90),
     NUM2,"How many times revenue covers support costs (>2.0 = healthy)"),
]

for row_data in eff_data:
    label,may_v,q3b,q3a,fmt,note=row_data
    if label in ("REVENUE PRODUCTIVITY","SUPPORT FUNCTION EFFICIENCY","REVENUE REGION METRICS","OPERATIONAL EFFICIENCY"):
        sec(ws,r,6,label,TEAL); r+=1; continue
    if label=="":
        r+=1; continue
    lbl(ws,r,1,label,bold=False)
    if may_v!="":
        inp(ws,r,2,may_v,fmt)
        inp(ws,r,3,q3b,fmt)
        inp(ws,r,4,q3a,fmt); ws.cell(row=r,column=4).font=Font(name="Arial",color="006400",size=10)
        delta=q3a-q3b if isinstance(q3a,(int,float)) and isinstance(q3b,(int,float)) else 0
        inp(ws,r,5,delta,fmt); ws.cell(row=r,column=5).font=Font(name="Arial",
            color="006400" if delta>0 else "C00000",size=10)
        ws.cell(row=r,column=6,value=note).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
        ws.cell(row=r,column=6).alignment=Alignment(wrap_text=True)
    r+=1

print("Efficiency Metrics done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 8: RISK REGISTER
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Risk Register"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 30
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 12
ws.column_dimensions['D'].width = 12
ws.column_dimensions['E'].width = 16
ws.column_dimensions['F'].width = 16
ws.column_dimensions['G'].width = 28
ws.column_dimensions['H'].width = 20

hdr(ws,1,1,"RISK REGISTER — COST REALLOCATION EXECUTION RISKS",bg="C00000",merge=(1,8),sz=13)
ws.row_dimensions[1].height=28

for col,txt,bg in [(1,"Risk","C00000"),(2,"Category",DKGRAY),(3,"Likelihood",DKGRAY),
                    (4,"Impact",DKGRAY),(5,"P&L Impact",DKGRAY),(6,"Cash Impact",DKGRAY),
                    (7,"Mitigation","006400"),(8,"Owner",DKGRAY)]:
    c=ws.cell(row=2,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=bg)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center",wrap_text=True)
ws.row_dimensions[2].height=24

risks=[
    ("IT SERVICE DEGRADATION","","","","","","",""),
    ("IT cost cut degrades system uptime/response time","IT / Operational","High","High",
     "-$800K revenue at risk","−$800K","Maintain SLA contracts; cut licences not infra; negotiate cloud cost optimisation instead of raw cuts","CTO / CFO"),
    ("Cloud services downtime — AWS/Azure cost optimisation backfires","IT / Operational","Medium","High",
     "-$500K","−$500K","Audit cloud usage; eliminate idle resources first; avoid cutting production workloads","CTO"),
    ("Loss of IT support staff — knowledge drain","HR / IT","Medium","Medium",
     "-$200K productivity loss","0","Retain tier-1 support; consider managed service transition for tier-2/3","CTO / HR"),
    ("","","","","","","",""),
    ("HR CAPACITY CONSTRAINTS","","","","","","",""),
    ("Recruitment freeze reduces pipeline for growth","HR / Revenue","High","High",
     "-$1.2M ARR at risk","0","Ring-fence revenue region headcount; only cut CORP/overhead HR","CFO / HR"),
    ("Training budget cut reduces skills competency","HR / Operational","Medium","Medium",
     "-$300K productivity","0","Shift to digital/online training; preserve customer-facing skill development","HR Manager"),
    ("High performer attrition if benefits reduced","HR / People","Medium","High",
     "-$600K replacement cost","−$600K","Protect compensation; target non-headcount HR costs (admin, office)","CHRO"),
    ("","","","","","","",""),
    ("OPS SERVICE DEGRADATION","","","","","","",""),
    ("Telecom cost reduction causes connectivity issues","OPS / Operational","Low","High",
     "-$400K SLA breach risk","−$400K","Consolidate contracts; move to volume pricing not line cuts","COO"),
    ("Freight/logistics cost cut slows delivery","OPS / Customer","Medium","High",
     "-$700K revenue risk","−$700K","Negotiate carrier rates vs cutting routes; protect QLD/WA SLAs","COO / Procurement"),
    ("Insurance reduction increases liability exposure","OPS / Legal","Low","High",
     "-$2M potential uninsured loss","−$2M","Do NOT cut insurance below minimum; focus on policy consolidation","CFO / Legal"),
    ("","","","","","","",""),
    ("CORP OVERHEAD CONSTRAINTS","","","","","","",""),
    ("Legal/professional service cuts delay contracts","CORP / Revenue","Medium","Medium",
     "-$500K deal velocity","0","Retain external legal for commercial contracts; cut advisory retainers","GC / CFO"),
    ("Audit fee cuts risk compliance failure","CORP / Compliance","Low","High",
     "-$1M+ regulatory risk","−$1M","Do NOT cut Big-4 external audit; target internal process audits","CFO"),
    ("Travel & accommodation cuts reduce client relationship quality","CORP / Revenue","Medium","Medium",
     "-$300K pipeline risk","0","Replace with video; keep executive client entertainment","Sales / CFO"),
    ("","","","","","","",""),
    ("REINVESTMENT EXECUTION RISKS","","","","","","",""),
    ("Revenue multiplier lower than modelled (3-3.5x assumption)","Revenue / Planning","High","High",
     "-$1.5M revenue vs model","−$1.5M","Track ROI monthly; if multiplier <2x by Aug, redirect to higher performers","CFO / Sales"),
    ("QLD/WA market saturation limits revenue uplift","Revenue / Market","Medium","Medium",
     "-$800K","0","Pre-qualify pipeline before reinvesting; target new verticals","CRO"),
    ("Marketing lag — MKT reinvestment takes 2-3 months to convert","Revenue / Timing","High","Medium",
     "-$500K Q3 miss","0","Front-load MKT investment to Aug to allow conversion time","CMO / CFO"),
    ("Rapid growth stretches operational capacity","Scalability","Medium","High",
     "-$1M service degradation","0","Hire contractors for peak capacity; monitor NPS weekly","COO / CFO"),
    ("","","","","","","",""),
    ("OPPORTUNITIES","","","","","","",""),
    ("Support function automation frees capacity for strategic work","Technology","High","High",
     "+$1M productivity gain","0","Invest freed IT budget in AI/automation tools for HR and OPS","CTO / CFO"),
    ("QLD resource industry boom accelerates multiplier to 4x","Revenue / Market","Medium","High",
     "+$2M upside","0","Monitor commodity cycle; scale reinvestment if multiplier confirms","CRO"),
    ("MKT reinvestment generates subscription growth (2.5x actual)","Revenue / Product","Medium","High",
     "+$750K ARR","0","Focus MKT spend on subscription acquisition; lowest churn segment","CMO"),
]

r=3
for row_data in risks:
    label=row_data[0]
    if label in ("IT SERVICE DEGRADATION","HR CAPACITY CONSTRAINTS","OPS SERVICE DEGRADATION",
                  "CORP OVERHEAD CONSTRAINTS","REINVESTMENT EXECUTION RISKS","OPPORTUNITIES"):
        bg="006400" if label=="OPPORTUNITIES" else "C00000"
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
        c=ws.cell(row=r,column=1,value=label)
        c.font=Font(name="Arial",bold=True,color=WHITE,size=11)
        c.fill=PatternFill("solid",start_color=bg)
        c.alignment=Alignment(horizontal="left",vertical="center",indent=1)
        r+=1; continue
    if label=="":
        r+=1; continue
    
    _,cat,lik,imp,pl_i,cash_i,mitig,owner=row_data
    is_opp=pl_i.startswith("+")
    row_bg=RED if (lik=="High" and imp=="High") else AMBER if (lik=="High" or imp=="High") else None
    
    if row_bg:
        for c in range(1,9): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=row_bg)
    
    ws.cell(row=r,column=1,value=label).font=Font(name="Arial",bold=True,size=10)
    ws.cell(row=r,column=1).alignment=Alignment(wrap_text=True,vertical="center")
    ws.cell(row=r,column=2,value=cat).font=Font(name="Arial",size=9,color=DKGRAY)
    ws.cell(row=r,column=3,value=lik).font=Font(name="Arial",size=10,
        color="C00000" if lik=="High" else "ED7D31" if lik=="Medium" else "006400")
    ws.cell(row=r,column=4,value=imp).font=Font(name="Arial",size=10,
        color="C00000" if imp=="High" else "ED7D31" if imp=="Medium" else "006400")
    ws.cell(row=r,column=5,value=pl_i).font=Font(name="Arial",size=10,
        color="006400" if is_opp else "C00000")
    ws.cell(row=r,column=6,value=cash_i).font=Font(name="Arial",size=10,
        color="006400" if is_opp else "C00000")
    ws.cell(row=r,column=7,value=mitig).font=Font(name="Arial",size=9,italic=True)
    ws.cell(row=r,column=7).alignment=Alignment(wrap_text=True,vertical="center")
    ws.cell(row=r,column=8,value=owner).font=Font(name="Arial",size=9,color=MID)
    ws.row_dimensions[r].height=36
    r+=1

print("Risk Register done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 9: OPTIMAL STRATEGY
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Optimal Strategy"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 36
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 18
ws.column_dimensions['D'].width = 18
ws.column_dimensions['E'].width = 18
ws.column_dimensions['F'].width = 22

hdr(ws,1,1,"OPTIMAL COST ALLOCATION STRATEGY — RECOMMENDATION",bg="006400",merge=(1,6),sz=13)
ws.row_dimensions[1].height=28

r=3
r=sec(ws,r,6,"STRATEGY RECOMMENDATION FRAMEWORK","006400")
strategies=[
    ("Strategy A: Aggressive Reallocation (Recommended)","006400"),
    ("  • 10% support reduction implemented in full across IT, HR, OPS, CORP","006400"),
    ("  • 60% reinvested: QLD (30%), NSW (30%), WA (25%), MKT (15%)","006400"),
    ("  • 40% retained as direct EBITDA improvement","006400"),
    ("  • 3-month ramp (80% → 90% → 100%) to manage transition risk","006400"),
    ("  • Insurance, legal/compliance, and tier-1 IT support ring-fenced (not cut)","006400"),
    ("  • Projected Q3 EBITDA uplift: ~$4.8M vs baseline","006400"),
    ("","DKGRAY"),
    ("Strategy B: Conservative (Partial Reallocation)","808080"),
    ("  • 5% support reduction only; 100% reinvested (no retention to EBITDA)","808080"),
    ("  • Lower execution risk but foregoes margin expansion","808080"),
    ("  • Projected Q3 EBITDA uplift: ~$1.8M vs baseline (lower)","808080"),
    ("","DKGRAY"),
    ("Strategy C: Status Quo","C00000"),
    ("  • No reallocation; support costs grow at 5% inflation","C00000"),
    ("  • Revenue regions underinvested relative to their multiplier potential","C00000"),
    ("  • Margin erosion risk as support bloat outpaces revenue growth","C00000"),
]
for text,color in strategies:
    if text=="":
        r+=1; continue
    lbl(ws,r,1,text,bold=text.startswith("Strategy"),color=color); r+=1

r+=1
r=sec(ws,r,6,"SENSITIVITY ANALYSIS — EBITDA UPLIFT vs REDUCTION %",MID)
for col,txt in [(1,"Reduction %"),(2,"Saving ($)"),(3,"EBITDA Uplift (40% ret.)"),(4,"Rev Uplift (60% reinvest × mult)"),(5,"Total Q3 Benefit"),(6,"NI Margin After")]:
    c=ws.cell(row=r,column=col,value=txt)
    c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
    c.fill=PatternFill("solid",start_color=MID)
    c.alignment=Alignment(horizontal="center" if col>1 else "left",vertical="center",wrap_text=True)
r+=1

for pct in [0.05,0.08,0.10,0.12,0.15]:
    saving=SUPPORT_TOTAL*pct*2.70  # Q3 weighted average
    ebitda_ret=saving*0.40
    reinv_rev=saving*0.60*sum(REINVEST_ALLOC[cc]*REV_MULT[cc] for cc in REINVEST_ALLOC)
    total_benefit=ebitda_ret+reinv_rev
    ni_margin=(q3_ni_after+ebitda_ret*0.70)/q3_rev_after
    
    lbl(ws,r,1,f"{pct:.0%} reduction",bold=(pct==0.10),
        color="006400" if pct==0.10 else "000000",
        bg=GREEN if pct==0.10 else None)
    inp(ws,r,2,saving,DOLR,pct==0.10)
    inp(ws,r,3,ebitda_ret,DOLR,pct==0.10)
    inp(ws,r,4,reinv_rev,DOLR,pct==0.10)
    inp(ws,r,5,total_benefit,DOLR,pct==0.10)
    inp(ws,r,6,ni_margin,PCT,pct==0.10)
    if pct==0.10:
        fill_row(ws,r,6,GREEN)
        for c in range(1,7): ws.cell(row=r,column=c).font=Font(name="Arial",bold=True,color="006400",size=10)
    r+=1

r+=2
r=sec(ws,r,6,"OPTIMAL REALLOCATION SPLIT — DETAILED RECOMMENDATION",NAVY)
rec=[
    ("Region / Function","Current Spend","Recommended Change","New Budget","Rationale","Timeline"),
    ("IT — Cloud & Licence optimisation",CC_TOTALS["IT"],f"-10% = ${CC_TOTALS['IT']*0.10:,.0f}",
     CC_TOTALS["IT"]*0.90,"Eliminate idle cloud resources; renegotiate SaaS","Month 1"),
    ("HR — Admin & recruitment processes",CC_TOTALS["HR"],f"-10% = ${CC_TOTALS['HR']*0.10:,.0f}",
     CC_TOTALS["HR"]*0.90,"Automate onboarding/payroll; retain talent acquisition for revenue BUs","Month 1-2"),
    ("OPS — Telecom & logistics negotiation",CC_TOTALS["OPS"],f"-10% = ${CC_TOTALS['OPS']*0.10:,.0f}",
     CC_TOTALS["OPS"]*0.90,"Volume contract renegotiation; route optimisation; NOT headcount","Month 1"),
    ("CORP — Travel, advisory & overheads",CC_TOTALS["CORP"],f"-10% = ${CC_TOTALS['CORP']*0.10:,.0f}",
     CC_TOTALS["CORP"]*0.90,"Replace advisory retainers; virtual comms; office consolidation","Month 1-3"),
    ("→ QLD Sales & Business Development",CC_TOTALS["QLD"],f"+${REINV_M['QLD']:,.0f}/mo",
     CC_TOTALS["QLD"]+REINV_M["QLD"],"Highest multiplier (3.5x); resource sector pipeline","Month 1"),
    ("→ NSW Enterprise Sales",CC_TOTALS["NSW"],f"+${REINV_M['NSW']:,.0f}/mo",
     CC_TOTALS["NSW"]+REINV_M["NSW"],"Enterprise account expansion; highest absolute revenue","Month 1"),
    ("→ WA Mining & Resources",CC_TOTALS["WA"],f"+${REINV_M['WA']:,.0f}/mo",
     CC_TOTALS["WA"]+REINV_M["WA"],"Strong 3.2x multiplier; underinvested relative to potential","Month 1"),
    ("→ MKT Subscription & Digital",CC_TOTALS["MKT"],f"+${REINV_M['MKT']:,.0f}/mo",
     CC_TOTALS["MKT"]+REINV_M["MKT"],"Subscription growth; 2x multiplier; lowest churn cohort","Month 1-2"),
]
for i,row_data in enumerate(rec):
    label,curr,change,new_b,rationale,timeline=row_data
    is_hdr=(i==0)
    bg=NAVY if is_hdr else ("E2EFDA" if "→" in str(label) else LGRAY)
    fg=WHITE if is_hdr else "006400" if "→" in str(label) else NAVY
    
    ws.cell(row=r,column=1,value=label).font=Font(name="Arial",bold=True,color=fg,size=10)
    ws.cell(row=r,column=1).fill=PatternFill("solid",start_color=bg)
    ws.cell(row=r,column=1).alignment=Alignment(wrap_text=True,vertical="center")
    
    for col,val,fmt in [(2,curr,DOLR),(3,change,None),(4,new_b,DOLR)]:
        if is_hdr:
            ws.cell(row=r,column=col,value=val).font=Font(name="Arial",bold=True,color=WHITE,size=10)
            ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=NAVY)
        else:
            if isinstance(val,(int,float)):
                inp(ws,r,col,val,fmt)
                ws.cell(row=r,column=col).font=Font(name="Arial",color=fg,size=10)
            else:
                ws.cell(row=r,column=col,value=val).font=Font(name="Arial",size=10,color=fg)
        ws.cell(row=r,column=col).fill=PatternFill("solid",start_color=bg)
    
    ws.cell(row=r,column=5,value=rationale).font=Font(name="Arial",size=9,italic=not is_hdr,color=fg)
    ws.cell(row=r,column=5).alignment=Alignment(wrap_text=True)
    ws.cell(row=r,column=5).fill=PatternFill("solid",start_color=bg)
    ws.cell(row=r,column=6,value=timeline).font=Font(name="Arial",size=10,color=fg)
    ws.cell(row=r,column=6).fill=PatternFill("solid",start_color=bg)
    ws.row_dimensions[r].height=28
    r+=1

r+=2
r=sec(ws,r,6,"CRITICAL RING-FENCES — DO NOT CUT","C00000")
ring_fences=[
    "Insurance (all lines) — maintain full coverage; Middle East risk escalating",
    "External audit fees — compliance non-negotiable; ASIC/ASX obligations",
    "Tier-1 IT support — system uptime directly impacts revenue; SLA must hold",
    "Revenue region headcount — QLD/NSW/WA/SA workforce are revenue generators",
    "AR collections team — $3.33M overdue requires full resourcing",
    "Legal (commercial contracts) — deal velocity depends on contract turnaround",
]
for rf in ring_fences:
    lbl(ws,r,1,f"⚠  {rf}",color="C00000"); r+=1

print("Optimal Strategy done")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET 10: EXECUTIVE SUMMARY
# ═══════════════════════════════════════════════════════════════════════════
ws = wb["Executive Summary"]
ws.sheet_view.showGridLines = False
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 34
ws.column_dimensions['C'].width = 22
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 22

hdr(ws,1,1,"EXECUTIVE SUMMARY — CFO BRIEFING",bg=NAVY,sz=16,merge=(1,5))
hdr(ws,2,1,"Strategic Cost Reallocation: Support Function Efficiency → Revenue Region Investment",bg=MID,sz=11,merge=(1,5))
hdr(ws,3,1,"Q3 2026 (Aug–Aug–Sep)  |  Baseline: May 2026 Actuals",bg=LIGHT,fg=NAVY,sz=10,merge=(1,5))
for r_h in [1,2,3]: ws.row_dimensions[r_h].height=26

r=5
# KPI impact cards
kpi_cards=[
    ("Q3 EBITDA Before","Q3 EBITDA After","Uplift",
     q3_ebitda_before,q3_ebitda_after,q3_ebitda_after-q3_ebitda_before,DOLR,MID),
    ("EBITDA Margin Before","EBITDA Margin After","Expansion",
     q3_ebitda_before/q3_rev_before,q3_ebitda_after/q3_rev_after,
     q3_ebitda_after/q3_rev_after-q3_ebitda_before/q3_rev_before,PCT,"006400"),
    ("Net Income Before","Net Income After","Uplift",
     q3_ni_before,q3_ni_after,q3_ni_after-q3_ni_before,DOLR,MID),
    ("Support Cost Before","Support Cost After","Saving",
     SUPPORT_TOTAL*3,SUPPORT_TOTAL*0.90*2.70,SUPPORT_TOTAL*3-SUPPORT_TOTAL*0.90*2.70,DOLR,"C00000"),
]

hdr(ws,r,2,"KEY IMPACT METRICS",bg=NAVY,merge=(1,4)); r+=1
for label_b,label_a,delta_lbl,val_b,val_a,delta,fmt,color in kpi_cards:
    hdr(ws,r,2,label_b,bg=LGRAY,fg=DKGRAY,bold=False,sz=9)
    hdr(ws,r,3,label_a,bg=LGRAY,fg=DKGRAY,bold=False,sz=9)
    hdr(ws,r,4,delta_lbl,bg=GREEN,fg="006400",bold=True,sz=9)
    
    c=ws.cell(row=r+1,column=2,value=val_b)
    c.number_format=fmt; c.font=Font(name="Arial",size=13,bold=True,color=DKGRAY)
    c.alignment=Alignment(horizontal="center"); c.fill=PatternFill("solid",start_color="F8F8F8")
    
    c=ws.cell(row=r+1,column=3,value=val_a)
    c.number_format=fmt; c.font=Font(name="Arial",size=13,bold=True,color=color)
    c.alignment=Alignment(horizontal="center"); c.fill=PatternFill("solid",start_color="F8F8F8")
    
    c=ws.cell(row=r+1,column=4,value=delta)
    c.number_format=fmt; c.font=Font(name="Arial",size=13,bold=True,color="006400")
    c.alignment=Alignment(horizontal="center"); c.fill=PatternFill("solid",start_color=GREEN)
    
    ws.row_dimensions[r].height=18; ws.row_dimensions[r+1].height=32
    r+=3

r+=1
# What we're doing and why
hdr(ws,r,2,"THE STRATEGY IN 30 SECONDS",bg=NAVY,merge=(1,4)); r+=1
summary_pts=[
    "Support functions (IT, HR, OPS, CORP) represent 68.4% of total costs — $27.7M monthly.",
    "A disciplined 10% efficiency programme (automation, contract optimisation, overhead rationalisation) releases $2.77M/month.",
    "40% ($1.11M) flows directly to EBITDA — immediate margin expansion from 27.8% toward 31%+.",
    "60% ($1.66M) is reinvested into the highest-multiplier revenue regions: QLD (3.5x), WA (3.2x), NSW (3.0x), MKT (2.0x).",
    f"Revenue uplift from reinvestment: ~${rev_uplift_total(0.80)+rev_uplift_total(0.90)+rev_uplift_total(1.00):,.0f} over Q3.",
    "Net Q3 total benefit: EBITDA improvement + revenue uplift exceeds $7M.",
    "Key safeguards: insurance, tier-1 IT support, revenue workforce, and AR collections team are ring-fenced.",
]
for pt in summary_pts:
    lbl(ws,r,2,f"• {pt}",sz=10,wrap=True); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    ws.row_dimensions[r].height=24; r+=1

r+=1
hdr(ws,r,2,"TOP 8 CFO ACTIONS",bg=NAVY,merge=(1,4)); r+=1
actions=[
    ("🟢 IMMEDIATE","Identify and terminate idle cloud resources and unused SaaS licences (IT -10%)",
     "CTO","Aug 1"),
    ("🟢 IMMEDIATE","Renegotiate top 5 supplier contracts for OPS: freight, telecom, utilities",
     "COO / Procurement","Aug 1"),
    ("🟢 IMMEDIATE","Release QLD + WA reinvestment budgets to regional sales managers",
     "CFO / CRO","Aug 1"),
    ("🟡 HIGH","Audit CORP advisory retainers — terminate where value is not demonstrated",
     "CFO","Aug 15"),
    ("🟡 HIGH","Automate HR admin: onboarding, payroll reconciliation, leave management",
     "CHRO / CTO","Aug–Aug"),
    ("🟡 HIGH","Set monthly ROI tracking: if multiplier < 2.0x by Aug, redirect reinvestment",
     "CFO","Monthly"),
    ("🔴 CRITICAL","Ring-fence: insurance, audit, tier-1 IT, revenue workforce — do NOT cut",
     "CFO","Ongoing"),
    ("🔵 MONITOR","Track EBITDA margin weekly vs target 31%; escalate if below 29% in any month",
     "CFO","Weekly"),
]
for priority,action,owner,timing in actions:
    ws.cell(row=r,column=2,value=priority).font=Font(name="Arial",bold=True,size=10)
    ws.cell(row=r,column=3,value=action).font=Font(name="Arial",size=9)
    ws.cell(row=r,column=3).alignment=Alignment(wrap_text=True)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    ws.cell(row=r,column=5,value=f"{owner} | {timing}").font=Font(name="Arial",size=9,color=MID)
    ws.row_dimensions[r].height=28; r+=1

r+=1
# Risk summary
hdr(ws,r,2,"KEY RISKS TO MONITOR",bg="C00000",merge=(1,4)); r+=1
top_risks=[
    ("HIGH","IT service degradation if production systems cut","Maintain SLA floor; cut licences not infrastructure"),
    ("HIGH","Revenue multiplier misses model (3.5x assumed QLD)","Monthly ROI review; redirect if <2x by Aug"),
    ("HIGH","HR attrition from restructure uncertainty","Communicate early; protect compensation; target process costs"),
    ("MED","MKT reinvestment conversion lag (2-3 months)","Front-load Aug spend; set pipeline KPIs by week 4"),
]
for severity,risk,action in top_risks:
    bg=RED if severity=="HIGH" else AMBER
    for c in range(2,6): ws.cell(row=r,column=c).fill=PatternFill("solid",start_color=bg)
    ws.cell(row=r,column=2,value=severity).font=Font(name="Arial",bold=True,size=10,color="C00000" if severity=="HIGH" else ORNG)
    ws.cell(row=r,column=3,value=risk).font=Font(name="Arial",size=9)
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    ws.cell(row=r,column=5,value=action).font=Font(name="Arial",size=9,italic=True,color=DKGRAY)
    ws.row_dimensions[r].height=24; r+=1

print("Executive Summary done")

# ═══════════════════════════════════════════════════════════════════════════
# FREEZE PANES & FINAL SETTINGS
# ═══════════════════════════════════════════════════════════════════════════
for sheet_name, freeze in [
    ("CC Baseline","A3"),("Reallocation Engine","A3"),("Regional P&L","A4"),
    ("Consolidated P&L","A3"),("Margin Analysis","A3"),("Efficiency Metrics","A3"),
    ("Risk Register","A3"),("Optimal Strategy","A3"),("Executive Summary","B5"),
]:
    wb[sheet_name].freeze_panes = freeze

# Set Executive Summary as active sheet
wb.active = wb["Executive Summary"]

# ═══════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════
OUT = "/sessions/practical-blissful-maxwell/mnt/outputs/Cost_Reallocation_Forecast_Q3_2026.xlsx"
wb.save(OUT)
print(f"Saved: {OUT}")
print(f"Size: {__import__('os').path.getsize(OUT):,} bytes")

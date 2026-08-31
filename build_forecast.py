"""
Quarterly Forecast Model - Q3 2026
Three Scenarios: Base, Upside (+20% Revenue), Risk-Adjusted (Geopolitical)
Source: May 2026 actual financial data
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DARK_NAVY   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "BDD7EE"
ACCENT_GOLD = "C9A235"
GREEN_FILL  = "E2EFDA"
RED_FILL    = "FFE5E5"
AMBER_FILL  = "FFF2CC"
WHITE       = "FFFFFF"
LIGHT_GREY  = "F2F2F2"
DARK_GREY   = "595959"
BLUE_INPUT  = "0000FF"
BLACK_CALC  = "000000"
GREEN_LINK  = "007030"

ACT = {
    "product_sales":    29865684.74,
    "service_rev":       6339650.50,
    "sub_rev":           3498642.58,
    "interest_income":    616333.56,
    "fx_gain":            225865.04,
    "total_rev":        40546176.42,
    "salaries":          1110732.19,
    "benefits":          1019737.24,
    "contractors":        924312.86,
    "rent":               868729.19,
    "utilities":          914299.66,
    "office_supplies":   1030030.90,
    "software":          1107364.55,
    "cloud":             1354642.51,
    "it_support":         958460.94,
    "prof_services":      958056.17,
    "legal":             1371577.63,
    "audit":             1340571.27,
    "travel":            1021794.25,
    "accommodation":     1122559.57,
    "meals":             1221692.84,
    "marketing":         1391033.95,
    "advertising":       1237476.07,
    "events":            1098494.59,
    "telecom":           1188371.93,
    "postage_courier":    924546.02,
    "insurance":         1310206.50,
    "training":          1060256.02,
    "recruitment":       1310481.05,
    "depreciation":      1330907.62,
    "amortization":       884697.36,
    "bank_charges":      1285755.28,
    "interest_exp":       966892.39,
    "fx_loss":           1157387.93,
    "total_exp":        31471068.48,
    "net_income":        9075107.94,
    "cash":             12911223.00,
    "ar_total":          9249700.00,
    "ar_overdue":        3326800.00,
}

NUM_FMT = '#,##0;(#,##0);"-"'
PCT_FMT = '0.0%;(0.0%);"-"'

def sc(cell, bold=False, fc=BLACK_CALC, bg=None, align="right", fmt=None, size=10, italic=False, wrap=False):
    cell.font = Font(name="Arial", bold=bold, color=fc, size=size, italic=italic)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fmt:
        cell.number_format = fmt

def thin_border():
    s = Side(style="thin")
    return Border(bottom=s)

def thick_border():
    s = Side(style="medium")
    return Border(bottom=s)

def col_w(ws, col, w):
    ws.column_dimensions[get_column_letter(col)].width = w

wb = Workbook()
ws_cover   = wb.active; ws_cover.title = "Cover"
ws_assum   = wb.create_sheet("Assumptions")
ws_pl      = wb.create_sheet("P&L Forecast")
ws_wc      = wb.create_sheet("Working Capital")
ws_cf      = wb.create_sheet("Cash Flow")
ws_risk    = wb.create_sheet("Risk Register")
ws_cfo     = wb.create_sheet("CFO Actions")
ws_dash    = wb.create_sheet("Dashboard")

for ws, tc in [(ws_cover,"1F3864"),(ws_assum,"2E75B6"),(ws_pl,"375623"),
               (ws_wc,"7030A0"),(ws_cf,"C55A11"),(ws_risk,"FF0000"),
               (ws_cfo,"C9A235"),(ws_dash,"1F3864")]:
    ws.sheet_properties.tabColor = tc

# ── COVER ────────────────────────────────────────────────────────────────────
ws = ws_cover
ws.sheet_view.showGridLines = False
col_w(ws,1,3); col_w(ws,2,55); col_w(ws,3,25)
ws.merge_cells("B2:C2"); c=ws["B2"]; c.value="QUARTERLY FINANCIAL FORECAST"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=18); ws.row_dimensions[2].height=36
ws.merge_cells("B3:C3"); c=ws["B3"]; c.value="Q3 2026 — Three-Scenario Sensitivity Analysis"
sc(c,bold=False,fc=WHITE,bg=MID_BLUE,align="center",size=12); ws.row_dimensions[3].height=22
ws.merge_cells("B4:C4"); c=ws["B4"]; c.value="Geopolitical Risk | Revenue Growth | CFO Action Plan"
sc(c,bold=False,fc=ACCENT_GOLD,bg=DARK_NAVY,align="center",size=10)
for i,(lbl,val) in enumerate([
    ("Prepared by:","CFO Office — Financial Planning & Analysis"),
    ("Baseline Period:","May 2026 Actuals"),
    ("Forecast Period:","Q3 2026 (Aug–Sep 2026)"),
    ("Model Date:","August 2026"),
    ("Classification:","CONFIDENTIAL — Internal Use Only"),
],6):
    ws.row_dimensions[i].height=18
    c=ws.cell(row=i,column=2,value=lbl); sc(c,bold=True,bg=LIGHT_GREY,align="left",size=10)
    c=ws.cell(row=i,column=3,value=val); sc(c,bg=LIGHT_GREY,align="left",size=10)
ws.merge_cells("B12:C12"); c=ws["B12"]; c.value="SCENARIO SUMMARY"
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="center",size=11)
for i,(sc_name,desc,bg_c) in enumerate([
    ("BASE","May 2026 cost base, moderate geopolitical cost increase (+5% logistics/insurance)",LIGHT_BLUE),
    ("UPSIDE","+20% revenue growth; operating leverage drives margin expansion; logistics headwinds managed",GREEN_FILL),
    ("RISK-ADJUSTED","Geopolitical escalation: +15% fuel/logistics, +10% supplier pricing, FX headwinds, delayed AR",RED_FILL),
],13):
    ws.row_dimensions[i].height=22
    ws.merge_cells(f"B{i}:C{i}")
    c=ws.cell(row=i,column=2,value=f"  {sc_name}: {desc}")
    sc(c,bg=bg_c,align="left",size=10,wrap=True)

# ── ASSUMPTIONS ───────────────────────────────────────────────────────────────
ws = ws_assum
ws.sheet_view.showGridLines = False
for i,w in enumerate([38,16,18,18,35],1): col_w(ws,i,w)
ws.merge_cells("A1:E1"); c=ws["A1"]; c.value="MODEL ASSUMPTIONS — Q3 2026 QUARTERLY FORECAST"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=12); ws.row_dimensions[1].height=28
ws.merge_cells("A2:E2"); c=ws["A2"]; c.value="Blue = Hardcoded Input  |  Source: May 2026 Actuals (PL_Statement_May2025_Comparative.csv)"
sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="center",size=9,italic=True)
for j,h in enumerate(["Assumption","Base","Upside (+20%)","Risk-Adjusted","Notes"],1):
    c=ws.cell(row=3,column=j,value=h); sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="center",size=10)

sections = [
    ("── REVENUE ASSUMPTIONS ────────────────────────────",None,None,None,""),
    ("May 2026 Baseline Revenue (AUD)",40546176,40546176,40546176,"Source: PL_Statement_May2025_Comparative.csv"),
    ("Revenue Growth Rate (quarterly)",0.00,0.20,0.10,"Upside=+20%; Risk=+10% volume constrained"),
    ("Product Sales Mix (%)",0.737,0.737,0.737,"Constant across scenarios"),
    ("Service Revenue Mix (%)",0.156,0.156,0.156,""),
    ("Subscription Revenue Mix (%)",0.086,0.086,0.086,"Recurring; most resilient"),
    ("","","","",""),
    ("── GEOPOLITICAL COST OVERLAYS ──────────────────────",None,None,None,""),
    ("Logistics / Courier uplift (%)",0.05,0.05,0.15,"Risk: Middle East re-routing, fuel surcharge"),
    ("Insurance (Marine/Cargo) uplift (%)",0.05,0.05,0.12,"War-risk clause activation"),
    ("Fuel-linked costs (Telecom+Utilities) uplift (%)",0.03,0.03,0.10,"Oil price pass-through"),
    ("Supplier pricing pressure (COGS proxy) (%)",0.02,0.02,0.10,"Input cost inflation"),
    ("FX Loss uplift (%)",0.05,0.05,0.20,"USD safe-haven demand weakens AUD"),
    ("","","","",""),
    ("── PERSONNEL & OPERATING COSTS ─────────────────────",None,None,None,""),
    ("Salary growth rate (%)",0.02,0.025,0.02,"Wage inflation; Upside adds headcount"),
    ("Benefits growth rate (%)",0.02,0.025,0.02,""),
    ("Contractor growth rate (%)",0.01,0.030,0.01,""),
    ("Marketing uplift (%)",0.03,0.08,0.03,"Upside: growth investment"),
    ("T&E uplift (%)",0.02,0.04,0.02,""),
    ("","","","",""),
    ("── WORKING CAPITAL ─────────────────────────────────",None,None,None,""),
    ("Debtor Days (DSO) - days",45,45,55,"Risk: +10 days geopolitical delay"),
    ("Inventory buffer build (AUD)",0,0,500000,"Risk: safety stock pre-build"),
    ("Opening Cash Balance (AUD)",12911223,12911223,12911223,"Source: GL_Cash_Balances_May2026.csv"),
    ("Capex Q3 forecast (AUD)",1500000,2000000,1200000,"Upside: growth; Risk: deferred"),
    ("Tax rate (%)",0.30,0.30,0.30,"Australian corporate tax rate"),
    ("Min cash buffer target (AUD)",5000000,5000000,7000000,"Risk: higher contingency buffer"),
]

ASM_ROW = {}
r = 4
for item in sections:
    lbl,base,up,risk,note = item
    if str(base)=="" and lbl=="":
        ws.row_dimensions[r].height=8; r+=1; continue
    if base is None:
        ws.merge_cells(f"A{r}:E{r}")
        c=ws.cell(row=r,column=1,value=lbl)
        sc(c,bold=True,fc=WHITE,bg=DARK_GREY,align="left",size=9)
        ws.row_dimensions[r].height=16; r+=1; continue
    ws.row_dimensions[r].height=17
    c=ws.cell(row=r,column=1,value=lbl); sc(c,bg=WHITE,align="left",size=10)
    for j,(col,val) in enumerate(zip(["B","C","D"],[base,up,risk]),2):
        c=ws.cell(row=r,column=j,value=val)
        fmt=PCT_FMT if isinstance(val,float) and val<2 else NUM_FMT
        sc(c,fc=BLUE_INPUT,bg=AMBER_FILL if j==2 else WHITE,fmt=fmt,size=10)
    c=ws.cell(row=r,column=5,value=note); sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="left",size=9,italic=True)
    ASM_ROW[lbl]=r; r+=1

# Row references in Assumptions sheet (by row number)
# Row 5 = baseline rev, Row 6 = growth rate, Row 7-9 = mix
# Row 12 = logistics uplift, Row 13 = insurance, Row 14 = fuel/utilities
# Row 15 = supplier pricing, Row 16 = FX loss uplift
# Row 19 = salary, Row 20 = benefits, Row 21 = contractors, Row 22 = marketing, Row 23 = T&E
# Row 26 = DSO, Row 27 = inventory buffer, Row 28 = opening cash, Row 29 = capex, Row 30 = tax, Row 31 = min buffer

# ── P&L FORECAST ─────────────────────────────────────────────────────────────
ws = ws_pl
ws.sheet_view.showGridLines = False
ws.freeze_panes = "B5"
for i,w in enumerate([42,18,18,20,20,26],1): col_w(ws,i,w)
ws.merge_cells("A1:F1"); c=ws["A1"]; c.value="Q3 2026 QUARTERLY P&L FORECAST — THREE SCENARIOS"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=13); ws.row_dimensions[1].height=28
ws.merge_cells("A2:F2"); c=ws["A2"]
c.value="Amounts in AUD | Source baseline: May 2026 Actuals (PL_Statement_May2025_Comparative.csv)"
sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="center",size=9,italic=True)
for j,(h,bg,fg) in enumerate(zip(
    ["P&L LINE ITEM","May 2026\nActual","Q3 2026\nBase","Q3 2026\nUpside (+20%)","Q3 2026\nRisk-Adjusted","Commentary"],
    [DARK_NAVY,MID_BLUE,LIGHT_BLUE,GREEN_FILL,RED_FILL,DARK_NAVY],
    [WHITE,WHITE,BLACK_CALC,BLACK_CALC,BLACK_CALC,WHITE]),1):
    c=ws.cell(row=4,column=j,value=h); sc(c,bold=True,fc=fg,bg=bg,align="center",size=10,wrap=True)
ws.row_dimensions[4].height=30

def sec_hdr(ws,r,txt):
    ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value=txt)
    sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; return r+1

def plr(ws,r,lbl,act,base,up,risk,com="",bold=False,total=False,indent=False):
    bg=LIGHT_BLUE if total else (LIGHT_GREY if r%2==0 else WHITE)
    ws.row_dimensions[r].height=16
    c=ws.cell(row=r,column=1,value=("    " if indent else "")+lbl)
    sc(c,bold=bold,bg=bg,align="left",size=10)
    c.border=thick_border() if total else thin_border()
    for j,val in enumerate([act,base,up,risk],2):
        c=ws.cell(row=r,column=j,value=val)
        fc=GREEN_LINK if isinstance(val,str) and "!" in val else BLACK_CALC if isinstance(val,str) else BLUE_INPUT
        sc(c,bold=bold,fc=fc,bg=bg,fmt=NUM_FMT,size=10)
        c.border=thick_border() if total else thin_border()
    c=ws.cell(row=r,column=6,value=com); sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="left",size=9,italic=True,wrap=True)
    return r+1

def pct_r(ws,r,lbl,act,base,up,risk,com=""):
    ws.row_dimensions[r].height=14
    c=ws.cell(row=r,column=1,value="    "+lbl); sc(c,fc=DARK_GREY,bg=AMBER_FILL,align="left",size=9,italic=True)
    for j,val in enumerate([act,base,up,risk],2):
        c=ws.cell(row=r,column=j,value=val); sc(c,fc=BLACK_CALC,bg=AMBER_FILL,fmt=PCT_FMT,size=9)
    c=ws.cell(row=r,column=6,value=com); sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="left",size=9,italic=True)
    return r+1

r=5
# REVENUE
r=sec_hdr(ws,r,"REVENUE")
r=plr(ws,r,"Product Sales Revenue",ACT["product_sales"],
    "=Assumptions!B5*(1+Assumptions!B6)*Assumptions!B7",
    "=Assumptions!C5*(1+Assumptions!C6)*Assumptions!C7",
    "=Assumptions!D5*(1+Assumptions!D6)*Assumptions!D7","73.7% of total; dominant channel",indent=True)
r=plr(ws,r,"Service Revenue",ACT["service_rev"],
    "=Assumptions!B5*(1+Assumptions!B6)*Assumptions!B8",
    "=Assumptions!C5*(1+Assumptions!C6)*Assumptions!C8",
    "=Assumptions!D5*(1+Assumptions!D6)*Assumptions!D8","Professional services",indent=True)
r=plr(ws,r,"Subscription Revenue",ACT["sub_rev"],
    "=Assumptions!B5*(1+Assumptions!B6)*Assumptions!B9",
    "=Assumptions!C5*(1+Assumptions!C6)*Assumptions!C9",
    "=Assumptions!D5*(1+Assumptions!D6)*Assumptions!D9","Recurring; most resilient to disruption",indent=True)
r=plr(ws,r,"Interest Income",ACT["interest_income"],
    ACT["interest_income"]*1.02,ACT["interest_income"]*1.02,ACT["interest_income"]*0.98,
    "Risk: slight reduction as cash deployed for WC",indent=True)
r=plr(ws,r,"FX Gain",ACT["fx_gain"],
    ACT["fx_gain"],ACT["fx_gain"]*1.05,ACT["fx_gain"]*0.80,
    "Risk: geopolitical volatility compresses FX gains",indent=True)
total_rev_r=r
r=plr(ws,r,"TOTAL REVENUE",ACT["total_rev"],
    f"=SUM(B{r-5}:B{r-1})",f"=SUM(C{r-5}:C{r-1})",f"=SUM(D{r-5}:D{r-1})",
    "",bold=True,total=True)
rev_row=total_rev_r

r=pct_r(ws,r,"Revenue Growth vs May 2026",0.0,0.0,0.20,0.10,"Upside=+20% uplift; Risk=+10% volume constrained by demand softness")
r+=1

# SUPPLY CHAIN (primary geopolitical exposure)
r=sec_hdr(ws,r,"SUPPLY CHAIN & LOGISTICS  ◄ PRIMARY GEOPOLITICAL EXPOSURE")
r=plr(ws,r,"Postage, Freight & Courier",ACT["postage_courier"],
    f"={ACT['postage_courier']}*(1+Assumptions!B12)",
    f"={ACT['postage_courier']}*(1+Assumptions!C12)",
    f"={ACT['postage_courier']}*(1+Assumptions!D12)",
    "Risk: +15% fuel surcharge & Red Sea re-routing",indent=True)
r=plr(ws,r,"Insurance (Marine/Cargo/War-Risk)",ACT["insurance"],
    f"={ACT['insurance']}*(1+Assumptions!B13)",
    f"={ACT['insurance']}*(1+Assumptions!C13)",
    f"={ACT['insurance']}*(1+Assumptions!D13)",
    "Risk: +12% war-risk clause activation & route diversion premiums",indent=True)
telecom_util=ACT["telecom"]+ACT["utilities"]
r=plr(ws,r,"Telecom & Utilities (Fuel-linked)",telecom_util,
    f"={telecom_util}*(1+Assumptions!B14)",
    f"={telecom_util}*(1+Assumptions!C14)",
    f"={telecom_util}*(1+Assumptions!D14)",
    "Risk: +10% energy cost pass-through from oil price spike",indent=True)
r=plr(ws,r,"Supplier Pricing (Input Cost Proxy)",ACT["office_supplies"],
    f"={ACT['office_supplies']}*(1+Assumptions!B15)",
    f"={ACT['office_supplies']}*(1+Assumptions!C15)",
    f"={ACT['office_supplies']}*(1+Assumptions!D15)",
    "Risk: +10% raw material & component cost inflation",indent=True)
r=plr(ws,r,"FX Loss (Geopolitical AUD/USD Volatility)",ACT["fx_loss"],
    f"={ACT['fx_loss']}*(1+Assumptions!B16)",
    f"={ACT['fx_loss']}*(1+Assumptions!C16)",
    f"={ACT['fx_loss']}*(1+Assumptions!D16)",
    "Risk: +20% — USD safe-haven demand weakens AUD",indent=True)
log_tot_r=r
log_act=ACT["postage_courier"]+ACT["insurance"]+telecom_util+ACT["office_supplies"]+ACT["fx_loss"]
r=plr(ws,r,"TOTAL LOGISTICS & SUPPLY CHAIN",log_act,
    f"=SUM(B{r-5}:B{r-1})",f"=SUM(C{r-5}:C{r-1})",f"=SUM(D{r-5}:D{r-1})",
    "",bold=True,total=True)
log_tot_r=r-1; r+=1

# PERSONNEL
r=sec_hdr(ws,r,"PERSONNEL COSTS")
r=plr(ws,r,"Salaries & Wages",ACT["salaries"],
    f"={ACT['salaries']}*(1+Assumptions!B19)",
    f"={ACT['salaries']}*(1+Assumptions!C19)",
    f"={ACT['salaries']}*(1+Assumptions!D19)","",indent=True)
r=plr(ws,r,"Employee Benefits",ACT["benefits"],
    f"={ACT['benefits']}*(1+Assumptions!B20)",
    f"={ACT['benefits']}*(1+Assumptions!C20)",
    f"={ACT['benefits']}*(1+Assumptions!D20)","",indent=True)
r=plr(ws,r,"Contractor Fees",ACT["contractors"],
    f"={ACT['contractors']}*(1+Assumptions!B21)",
    f"={ACT['contractors']}*(1+Assumptions!C21)",
    f"={ACT['contractors']}*(1+Assumptions!D21)","",indent=True)
pers_tot_r=r
r=plr(ws,r,"TOTAL PERSONNEL",ACT["salaries"]+ACT["benefits"]+ACT["contractors"],
    f"=SUM(B{r-3}:B{r-1})",f"=SUM(C{r-3}:C{r-1})",f"=SUM(D{r-3}:D{r-1})","",bold=True,total=True)
pers_tot_r=r-1; r+=1

# FACILITIES
r=sec_hdr(ws,r,"FACILITIES")
r=plr(ws,r,"Office Rent",ACT["rent"],ACT["rent"],ACT["rent"],ACT["rent"],"Fixed lease",indent=True)
fac_tot_r=r
r=plr(ws,r,"TOTAL FACILITIES",ACT["rent"],
    f"=B{r-1}",f"=C{r-1}",f"=D{r-1}","",bold=True,total=True)
fac_tot_r=r-1; r+=1

# IT
r=sec_hdr(ws,r,"IT & TECHNOLOGY")
r=plr(ws,r,"Software Licenses",ACT["software"],
    ACT["software"]*1.03,ACT["software"]*1.05,ACT["software"]*1.03,"",indent=True)
r=plr(ws,r,"Cloud Services",ACT["cloud"],
    ACT["cloud"]*1.04,ACT["cloud"]*1.08,ACT["cloud"]*1.04,"Upside: elastic compute for growth",indent=True)
r=plr(ws,r,"IT Support",ACT["it_support"],
    ACT["it_support"]*1.02,ACT["it_support"]*1.02,ACT["it_support"]*1.02,"",indent=True)
it_tot_r=r
r=plr(ws,r,"TOTAL IT",ACT["software"]+ACT["cloud"]+ACT["it_support"],
    f"=SUM(B{r-3}:B{r-1})",f"=SUM(C{r-3}:C{r-1})",f"=SUM(D{r-3}:D{r-1})","",bold=True,total=True)
it_tot_r=r-1; r+=1

# PROFESSIONAL SERVICES
r=sec_hdr(ws,r,"PROFESSIONAL & EXTERNAL SERVICES")
r=plr(ws,r,"Professional Services",ACT["prof_services"],
    ACT["prof_services"]*1.02,ACT["prof_services"]*1.04,ACT["prof_services"]*1.02,"",indent=True)
r=plr(ws,r,"Legal Fees",ACT["legal"],
    ACT["legal"]*1.03,ACT["legal"]*1.03,ACT["legal"]*1.08,
    "Risk: +8% sanctions review, trade compliance legal work",indent=True)
r=plr(ws,r,"Audit Fees",ACT["audit"],
    ACT["audit"]*1.02,ACT["audit"]*1.02,ACT["audit"]*1.02,"",indent=True)
ext_tot_r=r
r=plr(ws,r,"TOTAL PROFESSIONAL SERVICES",ACT["prof_services"]+ACT["legal"]+ACT["audit"],
    f"=SUM(B{r-3}:B{r-1})",f"=SUM(C{r-3}:C{r-1})",f"=SUM(D{r-3}:D{r-1})","",bold=True,total=True)
ext_tot_r=r-1; r+=1

# SALES & MARKETING
r=sec_hdr(ws,r,"SALES & MARKETING")
r=plr(ws,r,"Marketing",ACT["marketing"],
    f"={ACT['marketing']}*(1+Assumptions!B22)",
    f"={ACT['marketing']}*(1+Assumptions!C22)",
    f"={ACT['marketing']}*(1+Assumptions!D22)","Upside: +8% growth investment",indent=True)
r=plr(ws,r,"Advertising",ACT["advertising"],
    ACT["advertising"]*1.02,ACT["advertising"]*1.06,ACT["advertising"]*1.02,"",indent=True)
r=plr(ws,r,"Events",ACT["events"],
    ACT["events"]*1.01,ACT["events"]*1.04,ACT["events"]*0.95,"Risk: event cancellations reduce cost",indent=True)
mkt_tot_r=r
r=plr(ws,r,"TOTAL SALES & MARKETING",ACT["marketing"]+ACT["advertising"]+ACT["events"],
    f"=SUM(B{r-3}:B{r-1})",f"=SUM(C{r-3}:C{r-1})",f"=SUM(D{r-3}:D{r-1})","",bold=True,total=True)
mkt_tot_r=r-1; r+=1

# T&E
r=sec_hdr(ws,r,"TRAVEL & ENTERTAINMENT")
r=plr(ws,r,"Travel",ACT["travel"],
    f"={ACT['travel']}*(1+Assumptions!B23)",
    f"={ACT['travel']}*(1+Assumptions!C23)",
    f"={ACT['travel']}*(1+Assumptions!D23)","",indent=True)
r=plr(ws,r,"Accommodation",ACT["accommodation"],
    f"={ACT['accommodation']}*(1+Assumptions!B23)",
    f"={ACT['accommodation']}*(1+Assumptions!C23)",
    f"={ACT['accommodation']}*(1+Assumptions!D23)","",indent=True)
r=plr(ws,r,"Meals & Entertainment",ACT["meals"],
    f"={ACT['meals']}*(1+Assumptions!B23)",
    f"={ACT['meals']}*(1+Assumptions!C23)",
    f"={ACT['meals']}*(1+Assumptions!D23)","",indent=True)
te_tot_r=r
r=plr(ws,r,"TOTAL T&E",ACT["travel"]+ACT["accommodation"]+ACT["meals"],
    f"=SUM(B{r-3}:B{r-1})",f"=SUM(C{r-3}:C{r-1})",f"=SUM(D{r-3}:D{r-1})","",bold=True,total=True)
te_tot_r=r-1; r+=1

# DEVELOPMENT
r=sec_hdr(ws,r,"PEOPLE DEVELOPMENT")
r=plr(ws,r,"Training",ACT["training"],ACT["training"]*1.02,ACT["training"]*1.04,ACT["training"]*1.02,"",indent=True)
r=plr(ws,r,"Recruitment",ACT["recruitment"],ACT["recruitment"]*1.02,ACT["recruitment"]*1.05,ACT["recruitment"]*1.02,"",indent=True)
dev_tot_r=r
r=plr(ws,r,"TOTAL DEVELOPMENT",ACT["training"]+ACT["recruitment"],
    f"=SUM(B{r-2}:B{r-1})",f"=SUM(C{r-2}:C{r-1})",f"=SUM(D{r-2}:D{r-1})","",bold=True,total=True)
dev_tot_r=r-1; r+=1

# NON-CASH & FINANCE
r=sec_hdr(ws,r,"NON-CASH & FINANCE COSTS")
r=plr(ws,r,"Depreciation",ACT["depreciation"],ACT["depreciation"],ACT["depreciation"]*1.02,ACT["depreciation"],"",indent=True)
r=plr(ws,r,"Amortization",ACT["amortization"],ACT["amortization"],ACT["amortization"],ACT["amortization"],"",indent=True)
r=plr(ws,r,"Bank Charges",ACT["bank_charges"],ACT["bank_charges"]*1.02,ACT["bank_charges"]*1.03,ACT["bank_charges"]*1.04,"",indent=True)
r=plr(ws,r,"Interest Expense",ACT["interest_exp"],ACT["interest_exp"],ACT["interest_exp"],ACT["interest_exp"]*1.05,"Risk: higher borrowing cost if credit line drawn",indent=True)
nc_tot_r=r
r=plr(ws,r,"TOTAL NON-CASH & FINANCE",ACT["depreciation"]+ACT["amortization"]+ACT["bank_charges"]+ACT["interest_exp"],
    f"=SUM(B{r-4}:B{r-1})",f"=SUM(C{r-4}:C{r-1})",f"=SUM(D{r-4}:D{r-1})","",bold=True,total=True)
nc_tot_r=r-1; r+=1

# TOTALS
ws.row_dimensions[r].height=8; r+=1
exp_tot_r=r
r=plr(ws,r,"TOTAL OPERATING EXPENSES",ACT["total_exp"],
    f"=B{log_tot_r}+B{pers_tot_r}+B{fac_tot_r}+B{it_tot_r}+B{ext_tot_r}+B{mkt_tot_r}+B{te_tot_r}+B{dev_tot_r}+B{nc_tot_r}",
    f"=C{log_tot_r}+C{pers_tot_r}+C{fac_tot_r}+C{it_tot_r}+C{ext_tot_r}+C{mkt_tot_r}+C{te_tot_r}+C{dev_tot_r}+C{nc_tot_r}",
    f"=D{log_tot_r}+D{pers_tot_r}+D{fac_tot_r}+D{it_tot_r}+D{ext_tot_r}+D{mkt_tot_r}+D{te_tot_r}+D{dev_tot_r}+D{nc_tot_r}",
    "",bold=True,total=True)
exp_tot_r=r-1

ni_r=r
r=plr(ws,r,"NET INCOME",ACT["net_income"],
    f"=B{rev_row}-B{exp_tot_r}",
    f"=C{rev_row}-C{exp_tot_r}",
    f"=D{rev_row}-D{exp_tot_r}","",bold=True,total=True)
ni_r=r-1

r=pct_r(ws,r,"Net Profit Margin",ACT["net_income"]/ACT["total_rev"],
    f"=B{ni_r}/B{rev_row}",f"=C{ni_r}/C{rev_row}",f"=D{ni_r}/D{rev_row}",
    "Upside: operating leverage expansion; Risk: compressed by logistics + FX cost inflation")

da_total=ACT["depreciation"]+ACT["amortization"]
ebitda_act=ACT["net_income"]+ACT["interest_exp"]+da_total
ebitda_r=r
r=plr(ws,r,"EBITDA",ebitda_act,
    f"=B{ni_r}+{ACT['interest_exp']}+{ACT['depreciation']}+{ACT['amortization']}",
    f"=C{ni_r}+{ACT['interest_exp']}+{ACT['depreciation']*1.02}+{ACT['amortization']}",
    f"=D{ni_r}+{ACT['interest_exp']*1.05}+{ACT['depreciation']}+{ACT['amortization']}",
    "",bold=True,total=True)
ebitda_r=r-1

r=pct_r(ws,r,"EBITDA Margin",ebitda_act/ACT["total_rev"],
    f"=B{ebitda_r}/B{rev_row}",f"=C{ebitda_r}/C{rev_row}",f"=D{ebitda_r}/D{rev_row}",
    "Risk scenario: est. ~4-6pp compression vs Upside from combined cost headwinds")

# ── WORKING CAPITAL ───────────────────────────────────────────────────────────
ws=ws_wc
ws.sheet_view.showGridLines=False
for i,w in enumerate([38,18,18,18,18,26],1): col_w(ws,i,w)
ws.merge_cells("A1:F1"); c=ws["A1"]; c.value="WORKING CAPITAL & LIQUIDITY ANALYSIS — Q3 2026"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=13); ws.row_dimensions[1].height=28
ws.merge_cells("A2:F2"); c=ws["A2"]
c.value="Source: AR_Subledger_May2026.csv | GL_Cash_Balances_May2026.csv | Geopolitical risk overlay"
sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="center",size=9,italic=True)
for j,(h,bg,fg) in enumerate(zip(
    ["WORKING CAPITAL ITEM","May 2026\nActual","Q3 2026\nBase","Q3 2026\nUpside","Q3 2026\nRisk","Commentary"],
    [DARK_NAVY,MID_BLUE,LIGHT_BLUE,GREEN_FILL,RED_FILL,DARK_NAVY],
    [WHITE,WHITE,BLACK_CALC,BLACK_CALC,BLACK_CALC,WHITE]),1):
    c=ws.cell(row=4,column=j,value=h); sc(c,bold=True,fc=fg,bg=bg,align="center",size=10,wrap=True)
ws.row_dimensions[4].height=30

def wcr(ws,r,lbl,act,base,up,risk,com="",bold=False,total=False,indent=False):
    bg=LIGHT_BLUE if total else (LIGHT_GREY if r%2==0 else WHITE)
    ws.row_dimensions[r].height=16
    c=ws.cell(row=r,column=1,value=("    " if indent else "")+lbl)
    sc(c,bold=bold,bg=bg,align="left",size=10)
    for j,val in enumerate([act,base,up,risk],2):
        c=ws.cell(row=r,column=j,value=val)
        fc=GREEN_LINK if isinstance(val,str) and "!" in val else BLACK_CALC if isinstance(val,str) else BLUE_INPUT
        sc(c,bold=bold,fc=fc,bg=bg,fmt=NUM_FMT,size=10)
    c=ws.cell(row=r,column=6,value=com); sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="left",size=9,italic=True,wrap=True)
    return r+1

r=5
ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value="CASH POSITION")
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; r+=1

r=wcr(ws,r,"Opening Cash Balance (May 2026 GL)",ACT["cash"],ACT["cash"],ACT["cash"],ACT["cash"],
    "Source: GL_Cash_Balances_May2026.csv — AUD equivalent across AUS, NZL, SGP, UK entities")
cash_open_r=r-1
r=wcr(ws,r,"  + Net Income (Q3 Forecast)",ACT["net_income"],
    f"='P&L Forecast'!B{ni_r}",f"='P&L Forecast'!C{ni_r}",f"='P&L Forecast'!D{ni_r}",
    "Linked from P&L Forecast",indent=True)
ni_wc_r=r-1
r=wcr(ws,r,"  + D&A Add-back (non-cash)",da_total,da_total,da_total,da_total,
    "Non-cash expense add-back to operating cash flow",indent=True)
da_wc_r=r-1
r=wcr(ws,r,"  - Capex (Q3 Forecast)",0,-1500000,-2000000,-1200000,
    "Upside: growth investments; Risk: deferred non-critical capex",indent=True)
capex_wc_r=r-1
r=wcr(ws,r,"  - Tax Provision (30%)",0,
    f"=-'P&L Forecast'!B{ni_r}*0.3",
    f"=-'P&L Forecast'!C{ni_r}*0.3",
    f"=-'P&L Forecast'!D{ni_r}*0.3","Estimated quarterly tax provision",indent=True)
tax_wc_r=r-1
r=wcr(ws,r,"  +/- Working Capital Movement",0,-200000,-500000,-1700000,
    "Risk: AR delay + inventory build + supplier term tightening",indent=True)
wc_mov_r=r-1
closing_cash_r=r
r=wcr(ws,r,"CLOSING CASH BALANCE (Q3 2026 Est.)",
    ACT["cash"]+ACT["net_income"],
    f"=B{cash_open_r}+B{ni_wc_r}+B{da_wc_r}+B{capex_wc_r}+B{tax_wc_r}+B{wc_mov_r}",
    f"=C{cash_open_r}+C{ni_wc_r}+C{da_wc_r}+C{capex_wc_r}+C{tax_wc_r}+C{wc_mov_r}",
    f"=D{cash_open_r}+D{ni_wc_r}+D{da_wc_r}+D{capex_wc_r}+D{tax_wc_r}+D{wc_mov_r}",
    "",bold=True,total=True)
closing_cash_r=r-1
r=wcr(ws,r,"  Minimum Cash Buffer Target",5000000,5000000,5000000,7000000,
    "Risk scenario: higher buffer for geopolitical contingency",indent=True)
min_buf_r=r-1
r=wcr(ws,r,"  Cash Headroom vs Buffer",ACT["cash"]-5000000,
    f"=B{closing_cash_r}-B{min_buf_r}",
    f"=C{closing_cash_r}-C{min_buf_r}",
    f"=D{closing_cash_r}-D{min_buf_r}",
    "⚠ Monitor Risk scenario — potential liquidity squeeze if headroom < $2M")

r+=1
ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value="ACCOUNTS RECEIVABLE — COLLECTION RISK")
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; r+=1

r=wcr(ws,r,"Total AR Balance",ACT["ar_total"],
    ACT["ar_total"]*1.02,ACT["ar_total"]*1.20,ACT["ar_total"]*1.15,
    "Source: AR_Subledger_May2026.csv | Upside: higher revenue = higher AR; Risk: elevated + slower")
ar_tot_r=r-1
r=wcr(ws,r,"  Overdue AR Balance",ACT["ar_overdue"],
    ACT["ar_overdue"],ACT["ar_overdue"]*1.05,ACT["ar_overdue"]*1.80,
    "Risk: +80% overdue — geopolitical disruption delays customer payments",indent=True)
ar_over_r=r-1
r=wcr(ws,r,"  Doubtful Debt Provision (est.)",0,0,0,500000,
    "Risk: 15% provision on elevated overdue balance",indent=True)
ws.row_dimensions[r-1].height=15
# DSO row (days format)
c=ws.cell(row=r,column=1,value="    Debtor Days (DSO)")
sc(c,fc=DARK_GREY,bg=AMBER_FILL,align="left",size=9,italic=True)
for j,val in enumerate([45,45,45,55],2):
    c=ws.cell(row=r,column=j,value=val)
    sc(c,fc=BLUE_INPUT,bg=AMBER_FILL,fmt='0 "days"',size=9)
c=ws.cell(row=r,column=6,value="Risk: +10 days — delayed payment from resources, logistics, aviation sectors")
sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="left",size=9,italic=True); r+=1

r+=1
ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value="SUPPLY CHAIN WORKING CAPITAL IMPACTS")
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; r+=1
r=wcr(ws,r,"Inventory Pre-Build (Safety Stock)",0,0,0,500000,"Risk: pre-build to buffer supply disruption — cash absorbed immediately")
r=wcr(ws,r,"Supplier Term Tightening",0,0,0,-200000,"Risk: key suppliers move from 30→15 day terms, accelerating outflows")
r=wcr(ws,r,"Net Working Capital Impact (Q3)",0,-200000,-500000,-1700000,"Negative = cash absorbed; Risk scenario most capital-intensive")

# ── CASH FLOW ─────────────────────────────────────────────────────────────────
ws=ws_cf
ws.sheet_view.showGridLines=False
for i,w in enumerate([38,18,18,18,18,26],1): col_w(ws,i,w)
ws.merge_cells("A1:F1"); c=ws["A1"]; c.value="CASH FLOW BRIDGE — Q3 2026 SCENARIO ANALYSIS (INDIRECT METHOD)"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=13); ws.row_dimensions[1].height=28
ws.merge_cells("A2:F2"); c=ws["A2"]; c.value="Three scenarios | Indirect method | Linked from P&L Forecast and Working Capital sheets"
sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="center",size=9,italic=True)
for j,(h,bg,fg) in enumerate(zip(
    ["CASH FLOW ITEM","May 2026\nActual","Q3 2026\nBase","Q3 2026\nUpside","Q3 2026\nRisk","Commentary"],
    [DARK_NAVY,MID_BLUE,LIGHT_BLUE,GREEN_FILL,RED_FILL,DARK_NAVY],
    [WHITE,WHITE,BLACK_CALC,BLACK_CALC,BLACK_CALC,WHITE]),1):
    c=ws.cell(row=4,column=j,value=h); sc(c,bold=True,fc=fg,bg=bg,align="center",size=10,wrap=True)
ws.row_dimensions[4].height=30

def cfr(ws,r,lbl,act,base,up,risk,com="",bold=False,total=False,indent=False):
    bg=LIGHT_BLUE if total else (LIGHT_GREY if r%2==0 else WHITE)
    ws.row_dimensions[r].height=16
    c=ws.cell(row=r,column=1,value=("    " if indent else "")+lbl)
    sc(c,bold=bold,bg=bg,align="left",size=10)
    for j,val in enumerate([act,base,up,risk],2):
        c=ws.cell(row=r,column=j,value=val)
        fc=GREEN_LINK if isinstance(val,str) and "!" in val else BLACK_CALC if isinstance(val,str) else BLUE_INPUT
        sc(c,bold=bold,fc=fc,bg=bg,fmt=NUM_FMT,size=10)
    c=ws.cell(row=r,column=6,value=com); sc(c,fc=DARK_GREY,bg=LIGHT_GREY,align="left",size=9,italic=True,wrap=True)
    return r+1

r=5
ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value="A.  OPERATING ACTIVITIES")
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; r+=1

r=cfr(ws,r,"Net Income",ACT["net_income"],
    f"='P&L Forecast'!B{ni_r}",f"='P&L Forecast'!C{ni_r}",f"='P&L Forecast'!D{ni_r}","Linked from P&L Forecast")
ni_cf_r=r-1
r=cfr(ws,r,"  Add: Depreciation & Amortization",da_total,da_total,da_total,da_total,"Non-cash expense add-back",indent=True)
da_cf_r=r-1
r=cfr(ws,r,"  Less: Increase in AR",0,-188000,-1049970,-1499960,"Risk: AR grows + DSO extends = more cash tied up",indent=True)
ar_cf_r=r-1
r=cfr(ws,r,"  Less: Inventory Build",0,0,0,-500000,"Risk: safety stock absorbs WC",indent=True)
r=cfr(ws,r,"  Less: Prepayments / Deposits",0,-50000,-100000,-250000,"Risk: suppliers demand upfront payments",indent=True)
r=cfr(ws,r,"  Add: Increase in Trade Payables",0,150000,250000,-200000,"Risk: suppliers tighten terms",indent=True)
ocf_r=r
r=cfr(ws,r,"NET CASH FROM OPERATIONS",ACT["net_income"]+da_total,
    f"=SUM(B{ni_cf_r}:B{ocf_r-1})",f"=SUM(C{ni_cf_r}:C{ocf_r-1})",f"=SUM(D{ni_cf_r}:D{ocf_r-1})",
    "",bold=True,total=True)
ocf_r=r-1; r+=1

ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value="B.  INVESTING ACTIVITIES")
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; r+=1
r=cfr(ws,r,"  Capital Expenditure",0,-1500000,-2000000,-1200000,"Upside: growth capex; Risk: deferred",indent=True)
capex_cf_r=r-1
icf_r=r
r=cfr(ws,r,"NET CASH FROM INVESTING",0,
    f"=B{capex_cf_r}",f"=C{capex_cf_r}",f"=D{capex_cf_r}","",bold=True,total=True)
icf_r=r-1; r+=1

ws.merge_cells(f"A{r}:F{r}"); c=ws.cell(row=r,column=1,value="C.  FINANCING ACTIVITIES")
sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=10); ws.row_dimensions[r].height=16; r+=1
r=cfr(ws,r,"  Debt Drawdown (revolving credit)",0,0,0,1000000,"Risk: precautionary drawdown on RCF",indent=True)
r=cfr(ws,r,"  Dividends / Distributions",0,0,0,0,"No dividends forecast Q3",indent=True)
fcf_r=r
r=cfr(ws,r,"NET CASH FROM FINANCING",0,0,0,1000000,"",bold=True,total=True)
fcf_r=r-1; r+=1

net_cf_r=r
r=cfr(ws,r,"NET CHANGE IN CASH",ACT["net_income"]+da_total,
    f"=B{ocf_r}+B{icf_r}+B{fcf_r}",
    f"=C{ocf_r}+C{icf_r}+C{fcf_r}",
    f"=D{ocf_r}+D{icf_r}+D{fcf_r}","",bold=True,total=True)
net_cf_r=r-1
r=cfr(ws,r,"Opening Cash Balance (May 2026)",ACT["cash"],ACT["cash"],ACT["cash"],ACT["cash"],"")
open_cf_r=r-1
closing_cf_r=r
r=cfr(ws,r,"CLOSING CASH BALANCE",ACT["cash"],
    f"=B{net_cf_r}+B{open_cf_r}",
    f"=C{net_cf_r}+C{open_cf_r}",
    f"=D{net_cf_r}+D{open_cf_r}","",bold=True,total=True)
closing_cf_r=r-1
r=cfr(ws,r,"  Free Cash Flow (OCF – Capex)",ACT["net_income"]+da_total,
    f"=B{ocf_r}+B{icf_r}",
    f"=C{ocf_r}+C{icf_r}",
    f"=D{ocf_r}+D{icf_r}",
    "Key metric: Risk scenario FCF compression is primary CFO watch item",indent=True)

# ── RISK REGISTER ─────────────────────────────────────────────────────────────
ws=ws_risk
ws.sheet_view.showGridLines=False
for i,w in enumerate([4,26,16,38,12,12,12,16,18,32],1): col_w(ws,i,w)
ws.merge_cells("A1:J1"); c=ws["A1"]; c.value="FINANCIAL & OPERATIONAL RISK REGISTER — Q3 2026"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=13); ws.row_dimensions[1].height=28
ws.merge_cells("A2:J2"); c=ws["A2"]
c.value="Geopolitical Risk Focus: Middle East conflict → fuel, logistics, FX, supply chain, credit, cyber"
sc(c,bold=False,fc=DARK_GREY,bg=AMBER_FILL,align="center",size=10,italic=True)
for j,h in enumerate(["#","Risk","Category","Description / Trigger","Like-\nlihood\n(1-5)","Impact\n(1-5)","Score","Financial\nImpact\n(AUD)","Owner","Mitigation Actions"],1):
    c=ws.cell(row=4,column=j,value=h); sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="center",size=9,wrap=True)
ws.row_dimensions[4].height=36

risks=[
    (1,"Fuel & Energy Cost Spike","Supply Chain","Middle East conflict drives Brent crude >$100/bbl; fuel surcharges on all sea and air freight routes",4,4,138682,"COO","Lock fixed-rate freight contracts; fuel hedges; rail diversification"),
    (2,"Shipping Delays & Port Congestion","Logistics","Red Sea re-routing adds 10-14 days transit; port backlogs compound",4,4,500000,"COO","Pre-build safety stock; dual-source suppliers; alternative port routing"),
    (3,"Supplier Pricing Pressure","Procurement","Input costs +8-12%; energy and logistics inflation passed through by suppliers",4,4,103003,"CPO","Long-term supply agreements; local supplier qualification; cost programs"),
    (4,"FX Volatility (AUD/USD)","Treasury","USD safe-haven demand; AUD depreciates 5-8%; import cost and USD obligation increase",4,3,231478,"CFO","FX forward contracts; natural hedging; USD revenue alignment"),
    (5,"Delayed AR Collections","Credit","Customers in exposed sectors delay payments; DSO extends 45→55 days",3,4,2658240,"CFO","AR acceleration; early payment discounts; credit insurance review"),
    (6,"Inventory Build-Up / Obsolescence","Operations","$500K+ safety stock tied up; obsolescence risk if disruption resolves quickly",3,3,500000,"COO","Demand-driven replenishment; JIT on stable items; write-down policy"),
    (7,"Marine Insurance Surcharge","Risk","War-risk clauses activated; premiums +10-15% for Middle East routes",5,2,157225,"CFO","Broker renegotiation; route diversification; captive insurance review"),
    (8,"Sanctions & Trade Compliance","Legal","New sanctions regimes require compliance review of supplier and customer base",2,5,250000,"GC","Sanctions screening tool; supplier due diligence; trade compliance team"),
    (9,"Credit Covenant Breach","Financial","EBITDA compression in Risk scenario may push leverage near bank covenant thresholds",2,5,0,"CFO","Covenant headroom analysis; proactive bank communication; waiver prep"),
    (10,"Cyber Risk (State-Sponsored)","Technology","Elevated cyber threat from state actors associated with Middle East conflict",3,4,500000,"CTO","Enhanced monitoring; penetration testing; incident response plan review"),
]
for i,(num,name,cat,desc,lhood,impact,fin,owner,mitig) in enumerate(risks):
    row=5+i; ws.row_dimensions[row].height=44
    score=lhood*impact
    bg=RED_FILL if score>=15 else (AMBER_FILL if score>=9 else GREEN_FILL)
    for j,val in enumerate([num,name,cat,desc,lhood,impact,f"={get_column_letter(5)}{row}*{get_column_letter(6)}{row}",fin,owner,mitig],1):
        c=ws.cell(row=row,column=j,value=val)
        fmt=NUM_FMT if j==8 else None
        sc(c,fc=BLACK_CALC,bg=bg,align="left" if j>1 else "center",size=9,wrap=(j in [4,10]),fmt=fmt)
ws.row_dimensions[15].height=16
ws.merge_cells("A15:J15"); c=ws.cell(row=15,column=1)
c.value="LEGEND:  Green = Low (score 1-8)   |   Amber = Medium (score 9-14)   |   Red = High (score 15-25)   |   Score = Likelihood × Impact"
sc(c,bold=True,fc=DARK_GREY,bg=LIGHT_GREY,align="center",size=9)

# ── CFO ACTIONS ───────────────────────────────────────────────────────────────
ws=ws_cfo
ws.sheet_view.showGridLines=False
for i,w in enumerate([4,28,18,12,46,16,14,18],1): col_w(ws,i,w)
ws.merge_cells("A1:H1"); c=ws["A1"]
c.value="CFO ACTION PLAN — MARGIN PROTECTION | WORKING CAPITAL OPTIMISATION | CASH PRESERVATION"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=13); ws.row_dimensions[1].height=28
ws.merge_cells("A2:H2"); c=ws["A2"]
c.value="Targeted actions to protect margins, optimise working capital, and preserve cash while sustaining the +20% growth trajectory"
sc(c,fc=WHITE,bg=MID_BLUE,align="center",size=10,italic=True); ws.row_dimensions[2].height=20
for j,h in enumerate(["#","CFO Action","Theme","Priority","Detail / Implementation Notes","Est. Benefit (AUD)","Timeline","Owner"],1):
    c=ws.cell(row=4,column=j,value=h); sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="center",size=10,wrap=True)
ws.row_dimensions[4].height=24

actions=[
    (1,"Freight Contract Re-Negotiation","Margin Protection","H",
     "Immediately renegotiate freight contracts to lock fixed rates for Q3-Q4 2026; target 3-6 month hedges on key lanes. Reduces fuel surcharge exposure by 30-40%. Engage top 3 logistics partners this week.",
     180000,"< 2 weeks","COO + Procurement"),
    (2,"Fuel Price Hedging","Margin Protection","H",
     "Enter fuel price swap agreements on jet/marine fuel components. Target 50% hedge coverage on logistics-linked fuel exposure. Coordinate with Group Treasury and banking syndicate.",
     138682,"< 4 weeks","CFO + Treasury"),
    (3,"FX Forward Contracts","Margin Protection","H",
     "Enter USD/AUD forward contracts to hedge 70-80% of USD-denominated payables and USD invoice exposure for Q3 2026. Reduces estimated FX loss variance by $231K vs Risk scenario.",
     231478,"< 2 weeks","CFO + Treasury"),
    (4,"Insurance Premium Renegotiation","Margin Protection","M",
     "Engage brokers to review war-risk endorsements. Explore route diversification to reduce war-zone exposure and cap premium increase at 5%. Evaluate captive insurance contribution.",
     78612,"< 6 weeks","CFO + Risk"),
    (5,"Supplier Cost Renegotiation","Margin Protection","H",
     "Audit top 20 suppliers for commodity/energy exposure. Negotiate cost-sharing or volume rebates. Identify alternative suppliers outside conflict-impacted zones. Target: cap COGS uplift at 5%.",
     103003,"4-8 weeks","CPO + Finance"),
    (6,"Discretionary Cost Freeze","Margin Protection","M",
     "Freeze non-critical discretionary spend: T&E above policy, non-essential events, low-ROI marketing. Estimated 10% reduction across T&E and events = $330K quarterly saving.",
     330000,"Immediate","CFO + All BUs"),
    (7,"AR Acceleration Programme","Working Capital","H",
     "Deploy dedicated AR team: contact all >30-day balances. Offer 1.5% early payment discount for settlement within 14 days. Prioritise resources and logistics sector customers. Target: DSO from 55 to 45 days.",
     1324500,"Immediate","CFO + AR Team"),
    (8,"Credit Insurance Review","Working Capital","H",
     "Review credit insurance coverage for top 20 customers. Increase limits for exposed sectors (resources, logistics, aviation). File claims on overdue balances where applicable.",
     332680,"< 3 weeks","CFO + Credit"),
    (9,"Supplier Payment Terms Extension","Working Capital","M",
     "Renegotiate payment terms with key suppliers from 30 to 45-60 days where relationships allow. Improves DPO and preserves cash. Target $300K-500K DPO improvement.",
     400000,"4-6 weeks","CPO + Finance"),
    (10,"Inventory Optimisation","Working Capital","M",
     "Establish inventory review committee. Use demand forecasting to right-size safety stock. Pre-build only long lead-time critical items. Avoid blanket build-up to limit WC absorption.",
     200000,"< 4 weeks","COO + Finance"),
    (11,"Revolving Credit Facility Pre-Activation","Cash Preservation","H",
     "Proactively confirm RCF availability with bank syndicate. Ensure documentation current for 48-hour drawdown capability if required. Do not draw unless headroom falls below $5M buffer.",
     0,"< 1 week","CFO + Treasurer"),
    (12,"Capex Deferral Review","Cash Preservation","M",
     "Review Q3 capex pipeline: defer non-critical projects one quarter. Prioritise revenue-generating capex. Estimated $300K deferral in Risk scenario preserves liquidity.",
     300000,"< 2 weeks","CFO + COO"),
    (13,"Weekly 13-Week Cash Flow Forecasting","Cash Preservation","H",
     "Upgrade from monthly to weekly cash flow forecasting during geopolitical risk period. Implement rolling 13-week model. Present to ExCo weekly with sensitivity analysis.",
     0,"Immediate","CFO + FP&A"),
    (14,"Intercompany Cash Pooling Optimisation","Cash Preservation","M",
     "Optimise cash pooling across AUS, NZL, SGP, UK entities. Repatriate idle cash from subsidiaries to central treasury. Target $500K liquidity improvement from intercompany netting.",
     500000,"< 3 weeks","Group Treasury"),
    (15,"Early Revenue Lock-In (Upside Enabler)","Growth Sustainability","M",
     "Accelerate Q4 contract negotiations; push subscription cross-sell; lock in advance payments where possible. Revenue certainty reduces cash flow volatility in any scenario.",
     500000,"Q3 2026","CEO + Sales"),
]
theme_col={"Margin Protection":LIGHT_BLUE,"Working Capital":GREEN_FILL,"Cash Preservation":AMBER_FILL,"Growth Sustainability":"E8D5F5"}
pri_col={"H":RED_FILL,"M":AMBER_FILL,"L":GREEN_FILL}
for i,(num,action,theme,pri,detail,benefit,timeline,owner) in enumerate(actions):
    row=5+i; ws.row_dimensions[row].height=50
    bg=theme_col.get(theme,WHITE)
    for j,val in enumerate([num,action,theme,pri,detail,benefit,timeline,owner],1):
        c=ws.cell(row=row,column=j,value=val)
        cell_bg=pri_col.get(pri,bg) if j==4 else bg
        fmt=NUM_FMT if j==6 else None
        sc(c,fc=BLACK_CALC,bg=cell_bg,align="left" if j>1 else "center",size=9,wrap=(j in [2,5,8]),fmt=fmt)
tot_r=5+len(actions)
ws.row_dimensions[tot_r].height=18
ws.merge_cells(f"A{tot_r}:E{tot_r}")
c=ws.cell(row=tot_r,column=1,value="TOTAL ESTIMATED FINANCIAL BENEFIT (AUD)")
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="right",size=10)
c=ws.cell(row=tot_r,column=6,value=f"=SUM(F5:F{tot_r-1})")
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,fmt=NUM_FMT,size=10)
ws.merge_cells(f"G{tot_r}:H{tot_r}")
c=ws.cell(row=tot_r,column=7,value="Q3 2026 Cumulative")
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=9)

# ── DASHBOARD ─────────────────────────────────────────────────────────────────
ws=ws_dash
ws.sheet_view.showGridLines=False
for i,w in enumerate([3,30,18,18,20,20,18,20],1): col_w(ws,i,w)
ws.merge_cells("A1:H1"); c=ws["A1"]; c.value="EXECUTIVE DASHBOARD — Q3 2026 SCENARIO SUMMARY"
sc(c,bold=True,fc=WHITE,bg=DARK_NAVY,align="center",size=14); ws.row_dimensions[1].height=36
ws.merge_cells("A2:H2"); c=ws["A2"]
c.value="Three-Scenario Financial Forecast | May 2026 Baseline | Geopolitical Risk Overlay | Source: Internal FP&A"
sc(c,fc=WHITE,bg=MID_BLUE,align="center",size=10); ws.row_dimensions[2].height=18

for j,(h,bg,fg) in enumerate(zip(
    ["KPI","May 2026\nActual","Q3 2026\nBase","Q3 2026\nUpside\n(+20%)","Q3 2026\nRisk-Adj.","Upside\nvs Base","Risk\nvs Base","Key Insight"],
    [DARK_NAVY,MID_BLUE,LIGHT_BLUE,GREEN_FILL,RED_FILL,GREEN_FILL,RED_FILL,ACCENT_GOLD],
    [WHITE,WHITE,BLACK_CALC,BLACK_CALC,BLACK_CALC,BLACK_CALC,BLACK_CALC,WHITE]),1):
    c=ws.cell(row=4,column=j,value=h); sc(c,bold=True,fc=fg,bg=bg,align="center",size=10,wrap=True)
ws.row_dimensions[4].height=36

def dash_sec(ws,r,txt):
    ws.merge_cells(f"A{r}:H{r}"); c=ws.cell(row=r,column=1,value=txt)
    sc(c,bold=True,fc=WHITE,bg=MID_BLUE,align="left",size=9)
    ws.row_dimensions[r].height=14; return r+1

def dash_kpi(ws,r,lbl,act,base,up,risk,up_v,risk_v,insight,fmt=NUM_FMT):
    bg=LIGHT_GREY if r%2==0 else WHITE
    ws.row_dimensions[r].height=16
    c=ws.cell(row=r,column=2,value=lbl); sc(c,bg=bg,align="left",size=10)
    vals=[(3,act,MID_BLUE),(4,base,LIGHT_BLUE),(5,up,GREEN_FILL),(6,risk,RED_FILL)]
    for col,val,cell_bg in vals:
        c=ws.cell(row=r,column=col,value=val)
        fc=GREEN_LINK if isinstance(val,str) and "!" in val else BLACK_CALC if isinstance(val,str) else BLUE_INPUT
        sc(c,fc=fc,bg=cell_bg if col==3 else bg,fmt=fmt,size=10)
    # variance cols
    for col,val,v_bg in [(7,up_v,GREEN_FILL),(8,risk_v,RED_FILL)]:
        c=ws.cell(row=r,column=col,value=val)
        sc(c,fc=BLACK_CALC,bg=v_bg,fmt='+#,##0;(#,##0);"-"' if fmt==NUM_FMT else '+0.0%;(0.0%);"-"',size=10)
    c=ws.cell(row=r,column=9 if ws.max_column>=9 else 8,value=insight)
    # insight in col H (8)
    c=ws.cell(row=r,column=8,value=insight); sc(c,fc=DARK_GREY,bg=AMBER_FILL,align="left",size=9,italic=True,wrap=True)
    return r+1

r=5
r=dash_sec(ws,r,"REVENUE")
r=dash_kpi(ws,r,"Total Revenue (AUD)",
    f"='P&L Forecast'!B{rev_row}",f"='P&L Forecast'!B{rev_row}",
    f"='P&L Forecast'!C{rev_row}",f"='P&L Forecast'!D{rev_row}",
    f"='P&L Forecast'!C{rev_row}-'P&L Forecast'!B{rev_row}",
    f"='P&L Forecast'!D{rev_row}-'P&L Forecast'!B{rev_row}",
    "Upside adds $8.1M; Risk adds $4.1M — growth constrained by demand softness")

r=dash_sec(ws,r,"PROFITABILITY")
r=dash_kpi(ws,r,"Net Income (AUD)",
    f"='P&L Forecast'!B{ni_r}",f"='P&L Forecast'!B{ni_r}",
    f"='P&L Forecast'!C{ni_r}",f"='P&L Forecast'!D{ni_r}",
    f"='P&L Forecast'!C{ni_r}-'P&L Forecast'!B{ni_r}",
    f"='P&L Forecast'!D{ni_r}-'P&L Forecast'!B{ni_r}",
    "Risk scenario: logistics+FX cost spikes erode ~40% of Base net income uplift")
r=dash_kpi(ws,r,"Net Profit Margin",
    f"='P&L Forecast'!B{ni_r}/'P&L Forecast'!B{rev_row}",
    f"='P&L Forecast'!B{ni_r}/'P&L Forecast'!B{rev_row}",
    f"='P&L Forecast'!C{ni_r}/'P&L Forecast'!C{rev_row}",
    f"='P&L Forecast'!D{ni_r}/'P&L Forecast'!D{rev_row}",
    None, None,
    "Operating leverage in Upside drives ~3pp margin expansion",fmt=PCT_FMT)
r=dash_kpi(ws,r,"EBITDA (AUD)",
    f"='P&L Forecast'!B{ebitda_r}",f"='P&L Forecast'!B{ebitda_r}",
    f"='P&L Forecast'!C{ebitda_r}",f"='P&L Forecast'!D{ebitda_r}",
    f"='P&L Forecast'!C{ebitda_r}-'P&L Forecast'!B{ebitda_r}",
    f"='P&L Forecast'!D{ebitda_r}-'P&L Forecast'!B{ebitda_r}",
    "Primary covenant metric — monitor Risk scenario EBITDA closely")
r=dash_kpi(ws,r,"EBITDA Margin",
    f"='P&L Forecast'!B{ebitda_r}/'P&L Forecast'!B{rev_row}",
    f"='P&L Forecast'!B{ebitda_r}/'P&L Forecast'!B{rev_row}",
    f"='P&L Forecast'!C{ebitda_r}/'P&L Forecast'!C{rev_row}",
    f"='P&L Forecast'!D{ebitda_r}/'P&L Forecast'!D{rev_row}",
    None, None,
    "Risk: estimated 4-6pp compression vs Upside from cost headwinds",fmt=PCT_FMT)

r=dash_sec(ws,r,"SUPPLY CHAIN & LOGISTICS")
r=dash_kpi(ws,r,"Total Logistics & Supply Chain Costs (AUD)",
    f"='P&L Forecast'!B{log_tot_r}",f"='P&L Forecast'!B{log_tot_r}",
    f"='P&L Forecast'!C{log_tot_r}",f"='P&L Forecast'!D{log_tot_r}",
    f"='P&L Forecast'!C{log_tot_r}-'P&L Forecast'!B{log_tot_r}",
    f"='P&L Forecast'!D{log_tot_r}-'P&L Forecast'!B{log_tot_r}",
    "Risk: $1.2M+ additional cost vs Base — primary geopolitical exposure")

r=dash_sec(ws,r,"CASH & LIQUIDITY")
r=dash_kpi(ws,r,"Closing Cash Balance (AUD)",
    ACT["cash"],
    f"='Working Capital'!B{closing_cash_r}",
    f"='Working Capital'!C{closing_cash_r}",
    f"='Working Capital'!D{closing_cash_r}",
    f"='Working Capital'!C{closing_cash_r}-'Working Capital'!B{closing_cash_r}",
    f"='Working Capital'!D{closing_cash_r}-'Working Capital'!B{closing_cash_r}",
    "Risk: cash headroom vs $7M buffer is key watch item")
r=dash_kpi(ws,r,"Accounts Receivable (AUD)",
    ACT["ar_total"],ACT["ar_total"]*1.02,ACT["ar_total"]*1.20,ACT["ar_total"]*1.15,
    (ACT["ar_total"]*1.20)-(ACT["ar_total"]*1.02),
    (ACT["ar_total"]*1.15)-(ACT["ar_total"]*1.02),
    "Risk: higher AR + slower collection = $1.3M+ WC absorbed")
r=dash_kpi(ws,r,"Overdue AR (AUD)",ACT["ar_overdue"],ACT["ar_overdue"],ACT["ar_overdue"]*1.05,ACT["ar_overdue"]*1.80,
    ACT["ar_overdue"]*0.05,ACT["ar_overdue"]*0.80,
    "Risk: +80% overdue — doubtful debt provision required; credit insurance essential")
r=dash_kpi(ws,r,"Free Cash Flow (AUD)",
    ACT["net_income"]+da_total,
    f"='Cash Flow'!B{ocf_r}+'Cash Flow'!B{icf_r}",
    f"='Cash Flow'!C{ocf_r}+'Cash Flow'!C{icf_r}",
    f"='Cash Flow'!D{ocf_r}+'Cash Flow'!D{icf_r}",
    f"='Cash Flow'!C{ocf_r}+'Cash Flow'!C{icf_r}-('Cash Flow'!B{ocf_r}+'Cash Flow'!B{icf_r})",
    f"='Cash Flow'!D{ocf_r}+'Cash Flow'!D{icf_r}-('Cash Flow'!B{ocf_r}+'Cash Flow'!B{icf_r})",
    "FCF compression in Risk is the primary signal for CFO action trigger")

# Key watch banner
ws.row_dimensions[r+1].height=22
ws.merge_cells(f"A{r+1}:H{r+1}")
c=ws.cell(row=r+1,column=1)
c.value="KEY CFO WATCH ITEMS:  Cash headroom vs buffer  |  AR collection velocity  |  Logistics cost trajectory  |  FX forward coverage  |  DSO trend  |  EBITDA covenant headroom"
sc(c,bold=True,fc=WHITE,bg=ACCENT_GOLD,align="left",size=9)

# Save
wb.save("/sessions/practical-blissful-maxwell/mnt/outputs/Q3_2026_Quarterly_Forecast.xlsx")
print("DONE")

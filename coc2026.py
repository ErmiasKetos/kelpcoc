"""
KELP Chain-of-Custody (CoC) Generator v3
- Landscape (matching original 792x612)
- Tall narrow vertical analysis column headers
- KELP logo image
- KELP Ordering ID
- Light blue shading on KELP Use Only areas (matching original)
- Proper sidebar layout
"""
import streamlit as st
import io, os
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.colors import HexColor, black, white
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ── Colours ──
KELP_TEAL = HexColor("#007272")
DARK_BLUE = HexColor("#1F4E79")
HDR_BG = HexColor("#1F4E79")
LIGHT_BLUE = HexColor("#D6E4F0")  # light blue from original
BORDER = black
PW, PH = landscape(letter)  # 792 x 612
ML, MR, MT, MB = 28, 28, 24, 18
CW = PW - ML - MR

# ── Test Catalogue ──
TEST_CATALOGUE = {
    "PHYSICAL/GENERAL CHEMISTRY": {
        "pH": {"pres":"None","cont":"250mL Plastic","hold":"15 min","m":{"Potable":"EPA 150.1","Non-Potable":"EPA 150.2"}},
        "Temperature": {"pres":"None","cont":"1L Plastic","hold":"Immediate","m":{"Non-Potable":"SM 2550 B"}},
        "Dissolved Oxygen": {"pres":"None","cont":"300mL BOD","hold":"15 min","m":{"Potable":"SM 4500-O"}},
        "Turbidity": {"pres":"None","cont":"250mL Plastic","hold":"48 hrs","m":{"Potable":"EPA 180.1","Non-Potable":"EPA 180.1"}},
        "Conductivity": {"pres":"None","cont":"500mL Plastic","hold":"28 days","m":{"Potable":"SM 2510B","Non-Potable":"EPA 120.1"}},
        "Alkalinity": {"pres":"None; Cool<=6C","cont":"250mL Plastic","hold":"14 days","m":{"Potable":"SM 2320B"}},
        "TDS": {"pres":"Cool<=6C","cont":"500mL Plastic","hold":"7 days","m":{"Potable":"SM 2540C","Non-Potable":"SM 2540C"}},
        "TSS": {"pres":"Cool<=6C","cont":"500mL Plastic","hold":"7 days","m":{"Non-Potable":"SM 2540D"}},
        "Total Solids": {"pres":"Cool<=6C","cont":"250mL Plastic","hold":"7 days","m":{"Potable":"SM 2540B","Non-Potable":"SM 2540B"}},
        "COD": {"pres":"H2SO4, Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Non-Potable":"EPA 410.4"}},
        "BOD 5-day": {"pres":"Cool<=6C","cont":"1L Plastic","hold":"48 hrs","m":{"Non-Potable":"SM 5210B"}},
    },
    "METALS": {
        "Hardness - Total": {"pres":"HNO3, pH<2","cont":"250mL Plastic","hold":"180 days","m":{"Potable":"SM 2340C","Non-Potable":"EPA 130.1"}},
        "Calcium - Total": {"pres":"HNO3, pH<2","cont":"250mL Plastic","hold":"180 days","m":{"Potable":"SM 3500-CaB","Non-Potable":"EPA 200.8"}},
        "Magnesium - Total": {"pres":"HNO3, pH<2","cont":"250mL Plastic","hold":"180 days","m":{"Potable":"SM 3500-MgB","Non-Potable":"EPA 200.8"}},
        "Chromium (VI)": {"pres":"Cool<=6C, pH 9.3-9.7","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 218.6","Non-Potable":"EPA 218.6"}},
        "Metals - EPA 200.8": {"pres":"HNO3, pH<2","cont":"250mL Plastic","hold":"180 days","m":{"Potable":"EPA 200.8"}},
        "Metals - EPA 6020B": {"pres":"HNO3, pH<2","cont":"500mL Plastic","hold":"180 days","m":{"Non-Potable":"EPA 6020B"}},
        "RCRA 8 Metals": {"pres":"HNO3, pH<2","cont":"500mL Plastic","hold":"180 days","m":{"Non-Potable":"EPA 6020B"}},
    },
    "INORGANICS": {
        "Bromide": {"pres":"None; Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Bromate": {"pres":"EDA, Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Chloride": {"pres":"None; Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Chlorite": {"pres":"EDA, Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Chlorate": {"pres":"EDA, Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Fluoride": {"pres":"None; Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Nitrate": {"pres":"Cool<=6C","cont":"250mL Plastic","hold":"48 hrs","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Nitrite": {"pres":"Cool<=6C","cont":"250mL Plastic","hold":"48 hrs","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Phosphate, Ortho": {"pres":"Cool<=6C","cont":"250mL Plastic","hold":"48 hrs","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Sulfate": {"pres":"Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 300.1","Non-Potable":"EPA 300.1"}},
        "Perchlorate": {"pres":"None; Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 314.2","Non-Potable":"EPA 314.0"}},
        "Cyanide, Total": {"pres":"NaOH, pH>12","cont":"1L Plastic","hold":"14 days","m":{"Potable":"SM 4500-CNE","Non-Potable":"SW-846 9012B"}},
        "Sulfide": {"pres":"NaOH+Zn Acetate","cont":"500mL Plastic","hold":"7 days","m":{"Non-Potable":"SM 4500-S2-D"}},
        "Sulfite": {"pres":"None; Immediate","cont":"250mL Plastic","hold":"Immediate","m":{"Potable":"SM 4500-SO3","Non-Potable":"SM 4500-SO3"}},
    },
    "NUTRIENTS": {
        "Ammonia as N": {"pres":"H2SO4, pH<2, Cool<=6C","cont":"500mL Plastic","hold":"28 days","m":{"Potable":"EPA 350.1","Non-Potable":"EPA 350.1"}},
        "TKN": {"pres":"H2SO4, pH<2, Cool<=6C","cont":"500mL Plastic","hold":"28 days","m":{"Potable":"EPA 351.2","Non-Potable":"EPA 351.2"}},
        "Total Phosphorus": {"pres":"H2SO4, pH<2, Cool<=6C","cont":"250mL Plastic","hold":"28 days","m":{"Potable":"EPA 365.1","Non-Potable":"EPA 365.1"}},
    },
    "ORGANICS": {
        "Surfactants MBAS": {"pres":"Cool<=6C","cont":"500mL Glass","hold":"48 hrs","m":{"Potable":"SM 5540C","Non-Potable":"SM 5540C"}},
        "DOC": {"pres":"H2SO4/HCl, pH<2","cont":"250mL Amber","hold":"28 days","m":{"Potable":"EPA 415.3","Non-Potable":"EPA 415.3"}},
        "TOC": {"pres":"H2SO4/HCl, pH<2","cont":"250mL Amber","hold":"28 days","m":{"Potable":"EPA 415.3","Non-Potable":"EPA 415.1"}},
    },
    "DISINFECTION": {
        "Free Chlorine": {"pres":"None; Immediate","cont":"250mL Plastic","hold":"15 min","m":{"Potable":"SM 4500-Cl G"}},
        "Total Chlorine DPD": {"pres":"None; Immediate","cont":"250mL Plastic","hold":"15 min","m":{"Potable":"SM 4500-Cl F","Non-Potable":"SM 4500-Cl F"}},
        "Combined Chlorine": {"pres":"None; Immediate","cont":"250mL Plastic","hold":"15 min","m":{"Potable":"SM 4500-Cl G"}},
        "Chlorine Dioxide": {"pres":"None; Immediate","cont":"250mL Plastic(amb)","hold":"Immediate","m":{"Potable":"SM 4500-ClO2 E"}},
    },
    "PFAS TESTING": {
        "PFAS 3-Compound": {"pres":"Trizma, Cool<=6C","cont":"250mL HDPE","hold":"14d(28d w/pres)","m":{"Potable":"EPA 537.1","Non-Potable":"EPA 1633"}},
        "PFAS 14-Compound": {"pres":"Trizma, Cool<=6C","cont":"250mL HDPE","hold":"14d(28d w/pres)","m":{"Potable":"EPA 537.1"}},
        "PFAS 18-Compound": {"pres":"Trizma, Cool<=6C","cont":"250mL HDPE","hold":"14d(28d w/pres)","m":{"Potable":"EPA 537.1","Non-Potable":"EPA 1633"}},
        "PFAS 25-Compound": {"pres":"Trizma, Cool<=6C","cont":"250mL HDPE","hold":"14d(28d w/pres)","m":{"Potable":"EPA 533","Non-Potable":"EPA 1633"}},
        "PFAS 40-Compound": {"pres":"Trizma, Cool<=6C","cont":"250mL HDPE","hold":"14d(28d w/pres)","m":{"Non-Potable":"EPA 1633"}},
    },
    "PACKAGES": {
        "Essential Home Water Test": {"pres":"HNO3/None","cont":"2x250mL","hold":"Varies","m":{"Potable":"EPA 200.8/300.1"}},
        "Complete Homeowner": {"pres":"HNO3/None","cont":"3x250mL","hold":"Varies","m":{"Potable":"EPA 200.8/300.1/SM"}},
        "PFAS Home Safety": {"pres":"Trizma/HNO3","cont":"HDPE+Plastic","hold":"Varies","m":{"Potable":"EPA 1633/200.8/300.1"}},
        "Real Estate Well Water": {"pres":"HNO3/None","cont":"3x250mL","hold":"Varies","m":{"Potable":"EPA 200.8/300.1/SM"}},
        "Conventional Loan Testing": {"pres":"HNO3/None","cont":"3x250mL","hold":"Varies","m":{"Potable":"EPA 200.8/300.1/SM"}},
        "Food & Beverage Water": {"pres":"HNO3/None","cont":"3x500mL","hold":"Varies","m":{"Non-Potable":"EPA 200.8/300.1/SM"}},
        "Agricultural Irrigation": {"pres":"HNO3/None","cont":"3x500mL","hold":"Varies","m":{"Non-Potable":"EPA 200.8/300.1/SM"}},
    },
}
MATRIX_CODES = {"Drinking Water (DW)":"DW","Ground Water (GW)":"GW","Wastewater (WW)":"WW","Surface Water (SW)":"SW","Product (P)":"P","Other (OT)":"OT"}
DATA_DELIVERABLES = ["Level I (Std)","Level II","Level III","Level IV","Other"]
RUSH_OPTIONS = ["Standard (5-10 Day)","Same Day","1 Day","2 Day","3 Day","4 Day","5 Day"]
TIME_ZONES = ["PT","MT","CT","ET","AK"]
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kelp_logo.png")

# ═══ DRAWING HELPERS ═══
def _cb(c,x,y,checked=False,sz=7):
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,sz,sz,fill=0)
    if checked:
        c.setFont("Helvetica-Bold",sz);c.setFillColor(black);c.drawCentredString(x+sz/2,y+1,"X")

def _cell(c,x,y,w,h,text="",fs=7,al="left",bold=False,bg=None):
    if bg:
        c.setFillColor(bg);c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,w,h,fill=0,stroke=1)
    if text:
        fn="Helvetica-Bold" if bold else "Helvetica"
        c.setFont(fn,fs);c.setFillColor(black)
        ty=y+(h-fs)/2+1
        if al=="center":c.drawCentredString(x+w/2,ty,str(text))
        elif al=="right":c.drawRightString(x+w-2,ty,str(text))
        else:c.drawString(x+2,ty,str(text))

def _lcell(c,x,y,w,h,label,value="",lfs=6,vfs=7):
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,w,h,fill=0,stroke=1)
    c.setFont("Helvetica",lfs);c.setFillColor(black);c.drawString(x+2,y+(h-lfs)/2+1,label)
    if value:
        lw=len(label)*lfs*0.48+4
        c.setFont("Helvetica",vfs);c.drawString(x+lw,y+(h-vfs)/2+1,str(value))

def _hcell(c,x,y,w,h,text="",fs=6):
    c.setFillColor(HDR_BG);c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,w,h,fill=0,stroke=1)
    if text:
        c.setFont("Helvetica-Bold",fs);c.setFillColor(white)
        c.drawCentredString(x+w/2,y+(h-fs)/2+1,str(text))

# ═══ PDF GENERATION ═══
def generate_coc_pdf(data, logo_path=None):
    buf=io.BytesIO()
    c=canvas.Canvas(buf,pagesize=landscape(letter))
    c.setTitle("KELP Chain-of-Custody")
    y=PH-MT

    # === SIDEBAR DIMENSIONS (right side, runs full height like original) ===
    sb_w = 130  # KELP Use Only sidebar width
    main_w = CW - sb_w  # main content width

    # ═══ HEADER ═══
    hh=42
    c.setStrokeColor(KELP_TEAL);c.setLineWidth(2);c.line(ML,y,PW-MR,y)
    # Logo
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(ImageReader(logo_path),ML+3,y-hh+4,width=105,height=hh-8,preserveAspectRatio=True,mask='auto')
        except: pass
    # Title
    c.setFont("Helvetica-Bold",13);c.setFillColor(black)
    c.drawCentredString(ML+main_w/2,y-16,"CHAIN-OF-CUSTODY")
    c.setFont("Helvetica",7);c.drawCentredString(ML+main_w/2,y-27,"Chain-of-Custody is a LEGAL DOCUMENT - Complete all relevant fields")
    # KELP USE ONLY box (top right)
    kx=ML+main_w
    c.setFillColor(LIGHT_BLUE);c.rect(kx,y-hh,sb_w,hh,fill=1,stroke=0)
    c.setStrokeColor(black);c.setLineWidth(0.5);c.rect(kx,y-hh,sb_w,hh,fill=0,stroke=1)
    c.setFont("Helvetica-Bold",7);c.setFillColor(black)
    c.drawCentredString(kx+sb_w/2,y-10,"KELP USE ONLY")
    c.setFont("Helvetica",6);c.drawString(kx+4,y-22,"KELP Ordering ID:")
    c.setLineWidth(0.3);c.line(kx+68,y-24,kx+sb_w-4,y-24)
    kid=data.get("kelp_ordering_id","")
    if kid:
        c.setFont("Helvetica-Bold",7);c.drawString(kx+70,y-22,kid)
    c.setFont("Helvetica",5);c.drawString(kx+4,y-32,"Format: KELP-MMDDYY-####")
    c.setStrokeColor(KELP_TEAL);c.setLineWidth(1.5);c.line(ML,y-hh,PW-MR,y-hh)
    y-=hh

    # ═══ CLIENT INFO (6 rows) ═══
    # Original layout: Left=x20..237, Center=x239..523/663, Right sidebar=x665..783
    # Our proportional: Left=ML to ML+lw, Center=ML+lw to ML+main_w, Sidebar=ML+main_w to ML+CW
    rh=12
    lw=230  # left column (Company, Street, Project Name etc)
    cw2=main_w-lw  # center column (Contact, Phone, Email, Invoice etc)

    rows=[
        ("Company Name: ","company_name","Contact/Report To: ","contact_name"),
        ("Street Address: ","street_address","Phone #: ","phone"),
        ("","","E-Mail: ","email"),
        ("","","Cc E-Mail: ","cc_email"),
        ("Customer Project #: ","project_number","Invoice to: ","invoice_to"),
    ]
    for ll,lk,rl,rk in rows:
        if ll:_lcell(c,ML,y-rh,lw,rh,ll,data.get(lk,""))
        else:_cell(c,ML,y-rh,lw,rh)
        _lcell(c,ML+lw,y-rh,cw2,rh,rl,data.get(rk,""))
        y-=rh

    # Row 6: Project Name | Invoice E-mail | Specify Container Size | Container Size legend
    _lcell(c,ML,y-rh,lw,rh,"Project Name: ",data.get("project_name",""))
    inv_w=cw2*0.55  # Invoice E-mail portion
    scs_w=cw2*0.45  # Specify Container Size portion
    _lcell(c,ML+lw,y-rh,inv_w,rh,"Invoice E-mail: ",data.get("invoice_email",""))
    _cell(c,ML+lw+inv_w,y-rh,scs_w,rh)
    c.setFont("Helvetica",7);c.setFillColor(black)
    c.drawCentredString(ML+lw+inv_w+scs_w/2,y-rh+3,"Specify Container Size")
    # Container Size legend in sidebar (spans 2 rows)
    _cell(c,ML+main_w,y-rh*2,sb_w,rh*2)
    c.setFont("Helvetica-Bold",5);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-4,"Container Size: (1) 1L, (2) 500mL, (3) 250mL,")
    c.drawString(ML+main_w+3,y-10,"(4) 125mL, (5) 100mL, (6) Other")
    y-=rh

    # Row 7: Site Collection | Purchase Order | Identify Container Preservative Type
    site_w=lw  # left: Site Collection
    po_w=inv_w  # center-left: Purchase Order
    icp_w=scs_w  # center-right: Identify Container Preservative Type
    _lcell(c,ML,y-rh,site_w,rh,"Site Collection Info/Facility ID (as applicable): ",data.get("site_info",""),lfs=5)
    _lcell(c,ML+site_w,y-rh,po_w,rh,"Purchase Order (if applicable): ",data.get("purchase_order",""),lfs=5.5)
    _cell(c,ML+site_w+po_w,y-rh,icp_w,rh)
    c.setFont("Helvetica",7);c.setFillColor(black)
    c.drawCentredString(ML+site_w+po_w+icp_w/2,y-rh+3,"Identify Container Preservative Type")
    # Preservative legend in sidebar (spans 3 rows)
    _cell(c,ML+main_w,y-rh*3,sb_w,rh*3)
    c.setFont("Helvetica-Bold",5);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-4,"Preservative Types: (1) None, (2) HNO3, (3)")
    c.drawString(ML+main_w+3,y-11,"H2SO4, (4) HCl, (5) NaOH, (6) Zn Acetate, (7)")
    c.drawString(ML+main_w+3,y-18,"NaHSO4, (8) Sod.Thiosulfate, (9) Ascorbic Acid,")
    c.drawString(ML+main_w+3,y-25,"(10) MeOH, (11) Other")
    y-=rh

    # Row 8: (empty left) | (if applicable): continued | Quote # | Analysis Requested area
    _cell(c,ML,y-rh,site_w,rh)  # empty left
    _cell(c,ML+site_w,y-rh,po_w,rh)  # empty or continuation
    c.setFont("Helvetica",5.5);c.setFillColor(black)
    # Quote # in center area
    _lcell(c,ML+site_w,y-rh,po_w,rh,"Quote #: ",data.get("quote_number",""))
    # Analysis Requested in right area
    _cell(c,ML+site_w+po_w,y-rh,icp_w,rh)
    c.setFont("Helvetica",7);c.setFillColor(black)
    c.drawCentredString(ML+site_w+po_w+icp_w/2,y-rh+3,"Analysis Requested")
    y-=rh

    # KELP Use Only sidebar starts here (spanning down)
    kelp_sb_top = y

    # ═══ TIME ZONE ROW ═══
    tzh=11
    tz_w=main_w*0.30  # time zone checkboxes area
    cs_w=main_w*0.70  # county/state area
    _cell(c,ML,y-tzh,tz_w,tzh)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+2,y-tzh+3,"Sample Collection Time Zone :")
    st_tz=data.get("time_zone","PT")
    for i,tz in enumerate(["AK","PT","MT","CT","ET"]):
        cx=ML+118+i*26
        _cb(c,cx,y-tzh+2,checked=(tz==st_tz),sz=6)
        c.setFont("Helvetica",6);c.setFillColor(black);c.drawString(cx+8,y-tzh+3,tz)
    _lcell(c,ML+tz_w,y-tzh,cs_w,tzh,"County / State origin of sample(s): ",data.get("county_state",""),lfs=5.5)
    # Sidebar: Project Mgr
    _cell(c,ML+main_w,y-tzh,sb_w,tzh,bg=LIGHT_BLUE)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-tzh+3,"Project Mgr.: "+data.get("project_manager",""))
    y-=tzh

    # ═══ DATA DELIVERABLES / REGULATORY / RUSH block ═══
    # Original: 4 rows total
    # Row 1: Data Deliverables (left) | Regulatory Program ... Reportable Yes No (center+right)
    # Row 2: (cont Level III) | Rush (Pre-approval required): Same Day 1Day 2Day 3Day 4Day
    # Row 3: (cont Level IV) | 5 Day / Other | DW PWSID # or WW Permit # as applicable
    # Row 4: (cont Other) | Other_ | Field Filtered Yes No | Analysis:
    rr = 11  # row height for this block
    dd_w = 130  # Data Deliverables column width

    # Row 1: Data Deliverables + Regulatory + Reportable
    _cell(c,ML,y-rr*4,dd_w,rr*4)  # Data Deliverables spans 4 rows
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+2,y-5,"Data Deliverables:")
    sd=data.get("data_deliverable","Level I (Std)")
    dd_items = [("Level I (Std)",0),("Level II",0),("Level III",1),("Level IV",2),("Other",3)]
    for label,row_idx in dd_items:
        if row_idx < 2:
            # Row 0: Level I and Level II side by side
            if label=="Level I (Std)":
                bx=ML+4; by=y-14
            else:
                bx=ML+60; by=y-14
        else:
            bx=ML+4; by=y-14-(row_idx)*rr
        _cb(c,bx,by,checked=(label==sd),sz=6)
        c.setFont("Helvetica",5.5);c.setFillColor(black)
        c.drawString(bx+8,by+1,label)

    # Regulatory Program row (center area)
    reg_w = main_w - dd_w
    _cell(c,ML+dd_w,y-rr,reg_w,rr)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+dd_w+2,y-rr+3,f"Regulatory Program (DW, RCRA, etc.) as applicable: {data.get('regulatory_program','')}")
    # Reportable checkboxes
    rpx = ML+dd_w+reg_w*0.60
    c.drawString(rpx,y-rr+3,"Reportable")
    rv=data.get("reportable","No")=="Yes"
    _cb(c,rpx+40,y-rr+2,checked=rv,sz=7);c.drawString(rpx+49,y-rr+3,"Yes")
    _cb(c,rpx+68,y-rr+2,checked=not rv,sz=7);c.drawString(rpx+77,y-rr+3,"No")

    # AcctNum sidebar
    _cell(c,ML+main_w,y-rr,sb_w,rr,bg=LIGHT_BLUE)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-rr+3,"AcctNum / Client ID: "+data.get("acct_num",""))

    # Row 2: Rush
    _cell(c,ML+dd_w,y-rr*2,reg_w,rr)
    c.setFont("Helvetica-Bold",6);c.setFillColor(black)
    c.drawString(ML+dd_w+2,y-rr*2+3,"Rush (Pre-approval required):")
    sr=data.get("rush","Standard (5-10 Day)")
    rush_items = ["Same Day","1 Day","2 Day","3 Day","4 Day"]
    for i,ro in enumerate(rush_items):
        rx=ML+dd_w+130+i*42
        _cb(c,rx,y-rr*2+2,checked=(ro==sr),sz=7)
        c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(rx+9,y-rr*2+3,ro)

    # Table # sidebar
    _cell(c,ML+main_w,y-rr*2,sb_w,rr,bg=LIGHT_BLUE)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-rr*2+3,"Table #: "+data.get("table_number",""))

    # Row 3: 5 Day + Other | DW PWSID
    half_reg = reg_w * 0.55
    _cell(c,ML+dd_w,y-rr*3,half_reg,rr)
    _cb(c,ML+dd_w+4,y-rr*3+2,checked=("5 Day"==sr),sz=7)
    c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(ML+dd_w+14,y-rr*3+3,"5 Day")
    c.drawString(ML+dd_w+60,y-rr*3+3,"Other ___________")
    # DW PWSID
    _cell(c,ML+dd_w+half_reg,y-rr*3,reg_w-half_reg,rr)
    c.drawString(ML+dd_w+half_reg+2,y-rr*3+3,f"DW PWSID # or WW Permit # as applicable: {data.get('pwsid','')}")

    # Profile / Template sidebar
    _cell(c,ML+main_w,y-rr*3,sb_w,rr,bg=LIGHT_BLUE)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-rr*3+3,"Profile / Template: "+data.get("profile_template",""))

    # Row 4: Other_ | Field Filtered + Analysis:
    half_reg2 = reg_w * 0.40
    _cell(c,ML+dd_w,y-rr*4,half_reg2,rr)
    c.setFont("Helvetica",5.5);c.setFillColor(black)
    c.drawString(ML+dd_w+2,y-rr*4+3,"Other ___________")
    # Field Filtered + Analysis in same cell
    ff_w = reg_w * 0.60
    _cell(c,ML+dd_w+half_reg2,y-rr*4,ff_w,rr)
    c.drawString(ML+dd_w+half_reg2+2,y-rr*4+3,"Field Filtered (if applicable):")
    ff=data.get("field_filtered","No")
    ffx=ML+dd_w+half_reg2+100
    _cb(c,ffx,y-rr*4+2,checked=(ff=="Yes"),sz=7);c.drawString(ffx+9,y-rr*4+3,"Yes")
    _cb(c,ffx+28,y-rr*4+2,checked=(ff=="No"),sz=7);c.drawString(ffx+37,y-rr*4+3,"No")
    # Analysis: field (right portion of same row area)
    c.drawString(ML+dd_w+half_reg2+ff_w*0.55+2,y-rr*4+3,"Analysis:")

    # Prelog sidebar (blank row completing sidebar)
    _cell(c,ML+main_w,y-rr*4,sb_w,rr,bg=LIGHT_BLUE)
    c.setFont("Helvetica",6);c.setFillColor(black)
    y -= rr*4

    # ═══ MATRIX LEGEND ═══
    lgh=9
    _cell(c,ML,y-lgh,main_w,lgh)
    c.setFont("Helvetica",5);c.setFillColor(black)
    c.drawString(ML+2,y-lgh+2,"* Insert in Matrix box below: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water (SW), Other (OT)")
    # Prelog sidebar
    _cell(c,ML+main_w,y-lgh,sb_w,lgh,bg=LIGHT_BLUE)
    c.setFont("Helvetica-Bold",5.5);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-lgh+2,"Prelog / Bottle Ord. ID: "+data.get("prelog_id",""))
    y-=lgh

    # ═══ SAMPLE TABLE ═══
    # Original column widths from template analysis (total table width = 769px in original = 17..786)
    # We scale to our main_w + sb_w = CW
    # Original columns:
    #   Customer Sample ID: 161px, Matrix: 32px, Comp/Grab: 27px
    #   Composite Start Date: 55px, Time: 35px
    #   Collected/End Date: 50px, Time: 35px
    #   # Cont: 20px, Res Chlorine Result: 23px, Units: 21px
    #   10 Analysis columns: 19px each = 190px total
    #   Sample Comment: 89px (col 22: x=674..763)
    #   Pres NC sidebar: 23px (col 23: x=763..786)
    # Total: 161+32+27+55+35+50+35+20+23+21+190+89+23 = 761
    
    samples=data.get("samples",[])
    acols=data.get("analysis_columns",[])
    
    # Scale factor: our drawing width / original width
    OW = 769.0  # original total width
    TW = main_w + sb_w  # our total width (main + sidebar)
    sf = TW / OW
    
    # Fixed columns (scaled from original)
    cid  = round(161 * sf)  # Customer Sample ID
    cmx  = round(32 * sf)   # Matrix
    ccg  = round(27 * sf)   # Comp/Grab
    csd  = round(55 * sf)   # Composite Start Date
    cst  = round(35 * sf)   # Composite Start Time
    ced  = round(50 * sf)   # Collected/End Date
    cet  = round(35 * sf)   # Collected/End Time
    cnc  = round(20 * sf)   # # Cont
    crr  = round(23 * sf)   # Residual Chlorine Result
    cru  = round(21 * sf)   # Residual Chlorine Units
    fixed_total = cid+cmx+ccg+csd+cst+ced+cet+cnc+crr+cru
    
    # Analysis columns: ALWAYS 10 columns (matching original template)
    NUM_ANALYSIS_COLS = 10
    acw_each = round(19 * sf)  # each analysis column width
    analysis_total = acw_each * NUM_ANALYSIS_COLS
    
    # Sample Comment column
    comment_w = round(89 * sf)
    
    # Preservation NC column (part of sidebar)
    pnc_w = round(23 * sf)
    
    # Adjust comment_w to absorb any rounding difference
    used = fixed_total + analysis_total + comment_w + pnc_w
    comment_w += (TW - used)  # absorb rounding
    
    # ═══ TABLE HEADERS ═══
    # Original: headers are ~45pt tall (y=258..305 = ~47pt)
    vert_hdr_h = 45
    hdr_top = y
    
    # -- Fixed column headers --
    cx = ML
    _hcell(c,cx,y-vert_hdr_h,cid,vert_hdr_h,"Customer Sample ID",fs=6)
    cx+=cid
    _hcell(c,cx,y-vert_hdr_h,cmx,vert_hdr_h,"Matrix *",fs=5)
    cx+=cmx
    _hcell(c,cx,y-vert_hdr_h,ccg,vert_hdr_h,"Comp /\nGrab",fs=5)
    cx+=ccg
    
    # Composite Start (grouped header)
    csw=csd+cst
    h1=12
    _hcell(c,cx,y-h1,csw,h1,"Composite  Start",fs=5)
    _hcell(c,cx,y-vert_hdr_h,csd,vert_hdr_h-h1,"Date",fs=5)
    _hcell(c,cx+csd,y-vert_hdr_h,cst,vert_hdr_h-h1,"Time",fs=5)
    cx+=csw
    
    # Collected/End (grouped header)
    cew=ced+cet
    _hcell(c,cx,y-h1,cew,h1,"Collected or Composite\nEnd",fs=4.5)
    _hcell(c,cx,y-vert_hdr_h,ced,vert_hdr_h-h1,"Date",fs=5)
    _hcell(c,cx+ced,y-vert_hdr_h,cet,vert_hdr_h-h1,"Time",fs=4.5)
    cx+=cew
    
    _hcell(c,cx,y-vert_hdr_h,cnc,vert_hdr_h,"#\nCont.",fs=5)
    cx+=cnc
    
    # Residual Chlorine (grouped header)
    rcw=crr+cru
    _hcell(c,cx,y-h1,rcw,h1,"Residual\nChlorine",fs=5)
    _hcell(c,cx,y-vert_hdr_h,crr,vert_hdr_h-h1,"Result",fs=5)
    _hcell(c,cx+crr,y-vert_hdr_h,cru,vert_hdr_h-h1,"Units",fs=5)
    cx+=rcw
    
    # -- 10 Analysis Requested columns (vertical rotated text) --
    # Fill first N columns with user's analysis categories, rest blank
    for ai in range(NUM_ANALYSIS_COLS):
        label = acols[ai] if ai < len(acols) else ""
        # Draw the header cell with dark blue background
        c.setFillColor(HDR_BG);c.rect(cx,y-vert_hdr_h,acw_each,vert_hdr_h,fill=1,stroke=0)
        c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(cx,y-vert_hdr_h,acw_each,vert_hdr_h,fill=0,stroke=1)
        if label:
            c.saveState()
            # Fit text: truncate if needed, use smaller font for long names
            fs = 5.5 if len(label) < 18 else 5 if len(label) < 22 else 4.5
            c.setFont("Helvetica-Bold",fs);c.setFillColor(white)
            c.translate(cx+acw_each/2+1.5,y-vert_hdr_h+3);c.rotate(90)
            c.drawString(0,0,label[:28])  # hard limit
            c.restoreState()
        cx+=acw_each
    
    # Sample Comment header
    _hcell(c,cx,y-vert_hdr_h,comment_w,vert_hdr_h,"Sample Comment",fs=6)
    cx+=comment_w
    
    # Preservation NC column header (rotated, part of sidebar)
    c.setFillColor(LIGHT_BLUE);c.rect(cx,y-vert_hdr_h,pnc_w,vert_hdr_h,fill=1,stroke=0)
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(cx,y-vert_hdr_h,pnc_w,vert_hdr_h,fill=0,stroke=1)
    c.saveState()
    c.setFont("Helvetica",3.5);c.setFillColor(black)
    c.translate(cx+pnc_w/2+1,y-vert_hdr_h+2);c.rotate(90)
    c.drawString(0,0,"Preservation non-conformance")
    c.restoreState()
    c.saveState()
    c.setFont("Helvetica",3.5);c.setFillColor(black)
    c.translate(cx+pnc_w/2-4,y-vert_hdr_h+2);c.rotate(90)
    c.drawString(0,0,"identified for sample.")
    c.restoreState()
    
    y -= vert_hdr_h
    
    # ═══ SAMPLE DATA ROWS ═══
    srh=13  # row height (matches original ~19pt scaled)
    max_rows=10  # original has 10 data rows
    sample_block_top = y
    
    for ri in range(max_rows):
        s=samples[ri] if ri<len(samples) else {}
        cx=ML
        
        # Fixed columns
        vals=[
            (s.get("sample_id",""),cid,"left",6),
            (s.get("matrix",""),cmx,"center",6),
            (s.get("comp_grab",""),ccg,"center",5.5),
            (s.get("start_date",""),csd,"center",5.5),
            (s.get("start_time",""),cst,"center",5.5),
            (s.get("end_date",""),ced,"center",5.5),
            (s.get("end_time",""),cet,"center",5.5),
            (s.get("num_containers",""),cnc,"center",6),
            (s.get("res_cl_result",""),crr,"center",5.5),
            (s.get("res_cl_units",""),cru,"center",5),
        ]
        for val,w,al,fs in vals:
            _cell(c,cx,y-srh,w,srh,str(val),fs=fs,al=al)
            cx+=w
        
        # 10 Analysis columns - mark X where applicable
        s_analyses = s.get("analyses",[])
        for ai in range(NUM_ANALYSIS_COLS):
            ac_name = acols[ai] if ai < len(acols) else ""
            chk = "X" if ac_name and ac_name in s_analyses else ""
            _cell(c,cx,y-srh,acw_each,srh,chk,fs=7,al="center",bold=True)
            cx+=acw_each
        
        # Sample Comment
        cmt = s.get("comment","")
        _cell(c,cx,y-srh,comment_w,srh,cmt,fs=5.5)
        cx+=comment_w
        
        # Preservation NC column (blank data cells)
        _cell(c,cx,y-srh,pnc_w,srh)
        
        y-=srh
    
    sample_block_bottom = y
    
    # ═══ SIDEBAR for sample table area (KELP Use Only + Sample Comment label) ═══
    # The sidebar in the original spans from the header down through the sample rows
    # It contains: "KELP Use Only" rotated vertically, "Sample Comment" label
    # In our layout, the sidebar is already drawn as part of the Pres NC column
    # But we need the "KELP Use Only" text and "Sample Comment" text in the sidebar area
    # These are on the LEFT side of the Pres NC column area
    # Actually - looking at original: cols 21-22 = x=664..674 (10px gap) + x=674..763 (Sample Comment)
    # The KELP Use Only and Sample Comment are in the sidebar header area
    # The Pres NC text is on the far right edge

    # ═══ ADDITIONAL INSTRUCTIONS / REMARKS ═══
    rmkh=26
    hw=main_w/2
    for side,lbl,key in [
        (ML,"Additional Instructions from KELP:","additional_instructions"),
        (ML+hw,"Customer Remarks / Special Conditions / Possible Hazards:","customer_remarks"),
    ]:
        _cell(c,side,y-rmkh,hw,rmkh)
        c.setFont("Helvetica-Bold",5.5);c.setFillColor(black);c.drawString(side+2,y-6,lbl)
        txt=data.get(key,"")
        if txt:
            c.setFont("Helvetica",5.5)
            words=txt.split();lines=[];cur=""
            for w in words:
                t=f"{cur} {w}".strip()
                if len(t)>85:
                    if cur:lines.append(cur)
                    cur=w
                else:cur=t
            if cur:lines.append(cur)
            for i,ln in enumerate(lines[:3]):c.drawString(side+2,y-14-i*7,ln)
    y-=rmkh

    # ═══ COOLER/TEMP ═══
    tmph=12;qw=main_w/5
    _lcell(c,ML,y-tmph,qw,tmph,"# Coolers: ",data.get("num_coolers",""),lfs=5.5)
    _lcell(c,ML+qw,y-tmph,qw,tmph,"Thermometer ID: ",data.get("thermometer_id",""),lfs=5.5)
    _lcell(c,ML+qw*2,y-tmph,qw*0.8,tmph,"Temp. (C): ",data.get("temperature",""),lfs=5.5)
    ix=ML+qw*2.8;iw=main_w-qw*2.8
    _cell(c,ix,y-tmph,iw,tmph)
    c.setFont("Helvetica",6);c.setFillColor(black);c.drawString(ix+2,y-tmph+3,"Sample Received on ice:")
    oi=data.get("received_on_ice","Yes")
    _cb(c,ix+85,y-tmph+2,checked=(oi=="Yes"),sz=6);c.drawString(ix+93,y-tmph+3,"Yes")
    _cb(c,ix+112,y-tmph+2,checked=(oi=="No"),sz=6);c.drawString(ix+120,y-tmph+3,"No")
    y-=tmph

    # ═══ SIGNATURES ═══
    sgh=12;sc1=main_w*0.28;sc2=main_w*0.14;sc3=main_w*0.28;sc4=main_w*0.30
    dm=data.get("delivery_method","FedEX")
    for ri in range(3):
        cx=ML
        _cell(c,cx,y-sgh,sc1,sgh);c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(cx+2,y-sgh+3,"Relinquished by/Company: (Signature)")
        cx+=sc1;_cell(c,cx,y-sgh,sc2,sgh);c.drawString(cx+2,y-sgh+3,"Date/Time:")
        cx+=sc2;_cell(c,cx,y-sgh,sc3,sgh);c.drawString(cx+2,y-sgh+3,"Received by/Company: (Signature)")
        cx+=sc3;_cell(c,cx,y-sgh,sc4,sgh);c.drawString(cx+2,y-sgh+3,"Date/Time:")
        # Right column
        sx=ML+main_w;_cell(c,sx,y-sgh,sb_w,sgh)
        if ri==0:c.drawString(sx+3,y-sgh+3,f"Tracking Number: {data.get('tracking_number','')}")
        elif ri==1:
            c.drawString(sx+3,y-sgh+3,"Delivered by:")
            for j,d in enumerate(["In-Person","Courier"]):
                _cb(c,sx+50+j*38,y-sgh+2,checked=(d==dm),sz=6);c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(sx+58+j*38,y-sgh+3,d)
        else:
            for j,s2 in enumerate(["FedEX","UPS","Other"]):
                _cb(c,sx+8+j*35,y-sgh+2,checked=(s2==dm),sz=6);c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(sx+16+j*35,y-sgh+3,s2)
        y-=sgh

    # ═══ FOOTER ═══
    c.setFont("Helvetica",5);c.setFillColor(black)
    c.drawCentredString(PW/2,y-8,"Submitting a sample via this chain of custody constitutes acknowledgment and acceptance of the KELP's Terms and Conditions")

    # ═══ PAGE 2: INSTRUCTIONS ═══
    # Original uses: Calibri-Bold 14pt title, Calibri 14pt subtitle,
    # 12pt body text, 10pt bullets, ~15pt line spacing, two columns
    # Left col: x=30..370, Right col: x=400..762
    # Sections 1+2 left; continuation of 2 + sections 3+4 right
    c.showPage()
    y2 = PH - 52  # title at y=52 from top in original

    # Title: centered, bold, 14pt
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(black)
    c.drawCentredString(PW / 2, y2, "Chain of Custody (COC) Instructions")
    y2 -= 17

    # Subtitle: centered, regular, 14pt
    c.setFont("Helvetica", 12)
    c.drawCentredString(PW / 2, y2, "Complete all relevant fields on the COC form. Incomplete information may cause delays.")
    y2 -= 30

    # ── Layout constants ──
    LX = 30       # left column x
    RX = 400      # right column x
    COL_W = 355   # usable width per column
    LS = 15       # line spacing (matching original ~15pt between baselines)
    BFS = 12      # body font size
    HFS = 12      # heading font size
    BUL = 18      # bullet indent from column x
    TXT = 36      # text indent from column x (after bullet)

    def draw_heading(c, x, y, text):
        """Draw section heading in bold."""
        c.setFont("Helvetica-Bold", HFS)
        c.setFillColor(black)
        c.drawString(x, y, text)
        return y - LS - 3

    def draw_bullet(c, x, y, label, desc, col_w):
        """Draw bullet point with bold label and regular description.
        Returns new y position after all wrapped lines."""
        # Bullet character
        c.setFont("Helvetica", 10)
        c.setFillColor(black)
        c.drawString(x + BUL, y + 1, "\u2022")

        # Build the full text and figure out wrapping
        if label:
            full = label + " " + desc
        else:
            full = desc

        # Wrap text to fit column
        max_w = col_w - TXT - 6
        lines = []
        words = full.split()
        cur = ""
        for w in words:
            test = (cur + " " + w).strip()
            tw = len(test) * BFS * 0.48  # approximate width
            if tw > max_w and cur:
                lines.append(cur)
                cur = w
            else:
                cur = test
        if cur:
            lines.append(cur)

        # Draw each line
        for i, line in enumerate(lines):
            ly = y - i * LS
            tx = x + TXT
            if i > 0:
                tx += 6  # continuation indent

            if i == 0 and label:
                # First line: bold label then regular desc
                c.setFont("Helvetica-Bold", BFS)
                c.setFillColor(black)
                lw = c.stringWidth(label, "Helvetica-Bold", BFS)
                c.drawString(x + TXT, ly, label)
                # Get remainder of first line after label
                rest = line[len(label):].strip()
                if rest:
                    c.setFont("Helvetica", BFS)
                    c.drawString(x + TXT + lw + 3, ly, rest)
            else:
                c.setFont("Helvetica", BFS)
                c.setFillColor(black)
                c.drawString(tx, ly, line)

        return y - len(lines) * LS

    # ═══ LEFT COLUMN ═══
    ly = y2

    # Section 1: Client & Project Information
    ly = draw_heading(c, LX, ly, "1. Client & Project Information:")

    sec1 = [
        ("Company Name:", "Your company\u2019s name."),
        ("Street Address:", "Your mailing address."),
        ("City, State, Zip:", "Your city, state, and zip code."),
        ("Contact/Report To:", "Person designated to receive results."),
        ("Customer Project # and Project Name:", "Your project reference number and name."),
        ("Site Collection Info/Facility ID:", "Project location or facility ID."),
        ("Time Zone:", "Sample collection time zone (e.g., AK, PT, MT, CT, ET) for accurate hold times."),
        ("Purchase Order #:", "Your PO number for invoicing, if applicable."),
        ("Invoice To:", "Contact person for the invoice."),
        ("Invoice Email:", "Email address for the invoice."),
        ("Phone #:", "Your contact phone number."),
        ("E-mail:", "Your email for correspondence and the final report."),
        ("Data Deliverable:", "Required data deliverable level (Standard, Level II III IV, Other)."),
        ("Field Filtered:", "Indicate if samples were filtered in the field (Yes/No)."),
        ("Quote #:", "Quote number, if applicable."),
        ("DW PWSID # or WW Permit #:", "Relevant drinking water or wastewater permit numbers, if applicable."),
    ]
    for label, desc in sec1:
        ly = draw_bullet(c, LX, ly, label, desc, COL_W)

    # Section 2: Sample Information
    ly -= 6
    ly = draw_heading(c, LX, ly, "2. Sample Information:")

    sec2_left = [
        ("Customer Sample ID:", "Unique sample identifier for the report."),
        ("Collected Date:", "Sample collection date (provide start and end dates for composites)."),
        ("Collected Time:", "Sample collection time (provide start and end times for composites)."),
        ("Comp/Grab:", "\"GRAB\" for single-point collection; \"COMP\" for combined samples."),
        ("Matrix:", "Sample type (e.g., DW, GW, WW, P, SW, OT)."),
        ("Container Size:", "Specify size (e.g., 1L, 500ml, Other)."),
    ]
    for label, desc in sec2_left:
        ly = draw_bullet(c, LX, ly, label, desc, COL_W)

    # ═══ RIGHT COLUMN ═══
    ry = y2

    # Continuation of Section 2
    sec2_right = [
        ("Container Preservation Type:", "Specify preservative (e.g., None, HNO3, H2SO4, Other)."),
        ("Analysis Requested:", "List tests or method numbers and check boxes for applicable samples."),
        ("Sample Comment:", "Notes about individual samples; identify MS/MSD samples here."),
        ("Residual Chlorine:", "Record results and units if measured."),
    ]
    for label, desc in sec2_right:
        ry = draw_bullet(c, RX, ry, label, desc, COL_W)

    # Section 3: Additional Information & Instructions
    ry -= 6
    ry = draw_heading(c, RX, ry, "3. Additional Information & Instructions:")

    sec3 = [
        ("Customer Remarks/Special Conditions/Possible Hazards:", "Note special instructions, potential hazards (attach SDS if possible), or requests for extra report copies."),
        ("Rush Request:", "For expedited results, circle an option (Same Day to 5 Day) and note the due date. Pre-approval from the lab is required for all rush requests, and surcharges apply."),
        ("Collected By:", "Printed name of the sample collector."),
        ("Collected By Signature:", "Signature of the sample collector."),
        ("Relinquished By/Received By:", "Sign and date at each transfer of custody."),
    ]
    for label, desc in sec3:
        ry = draw_bullet(c, RX, ry, label, desc, COL_W)

    # Section 4: Sample Acceptance Policy Summary
    ry -= 6
    ry = draw_heading(c, RX, ry, "4. Sample Acceptance Policy Summary:")

    c.setFont("Helvetica", BFS)
    c.setFillColor(black)
    c.drawString(RX + 4, ry, "For samples to be accepted, ensure:")
    ry -= LS

    sec4 = [
        "Complete COC documentation.",
        "Readable, unique sample ID on containers (indelible ink).",
        "Appropriate containers and sufficient volume.",
        "Receipt within holding time and temperature requirements.",
        "Containers are in good condition, seals intact (if used).",
        "Proper preservation, no headspace in volatile water samples.",
        "Adequate volume for MS/MSD if required.",
    ]
    for item in sec4:
        ry = draw_bullet(c, RX, ry, "", item, COL_W)

    # Closing paragraph
    ry -= 8
    c.setFont("Helvetica", BFS)
    c.setFillColor(black)
    closing = [
        "Failure to meet these may result in data qualifiers. A detailed policy is",
        "available from your Project Manager. Submitting samples implies",
        "acceptance of KELP Terms and Conditions.",
    ]
    for line in closing:
        c.drawString(RX + 4, ry, line)
        ry -= LS
    c.save();buf.seek(0);return buf

# ═══ STREAMLIT UI ═══
def main():
    st.set_page_config(page_title="KELP CoC Generator",page_icon="🧪",layout="wide")
    st.markdown("""<style>.main .block-container{padding-top:1rem;max-width:1200px}h1{color:#007272}h2,h3{color:#1F4E79}div[data-testid="stExpander"]{border:1px solid #007272;border-radius:8px}</style>""",unsafe_allow_html=True)
    c1,c2=st.columns([1,5])
    with c1:
        lf=os.path.join(os.path.dirname(os.path.abspath(__file__)),"kelp_logo.png")
        if os.path.exists(lf):st.image(lf,width=100)
        else:st.markdown('<div style="background:#007272;color:white;padding:10px;border-radius:8px;text-align:center"><b>KELP</b></div>',unsafe_allow_html=True)
    with c2:st.title("Chain-of-Custody Generator");st.caption("EPA/TNI-compliant CoC PDF generator • Landscape format")

    tab1,tab2,tab3,tab4,tab5=st.tabs(["👤 Client Info","🧪 Test Selection","🧫 Sample Details","📋 Additional Info","📄 Generate PDF"])

    with tab1:
        st.subheader("Client & Project Information")
        ca,cb2=st.columns(2)
        with ca:st.text_input("Company Name *",key="company_name");st.text_input("Street Address",key="street_address");st.text_input("Phone #",key="phone");st.text_input("E-Mail *",key="email");st.text_input("Cc E-Mail",key="cc_email")
        with cb2:st.text_input("Contact/Report To *",key="contact_name");st.text_input("Customer Project #",key="project_number");st.text_input("Project Name",key="project_name");st.text_input("Invoice To",key="invoice_to");st.text_input("Invoice E-Mail",key="invoice_email")
        st.divider();st.subheader("Site & Regulatory Details")
        cc2,cd=st.columns(2)
        with cc2:st.text_input("Site Collection Info / Facility ID",key="site_info");st.text_input("Purchase Order #",key="purchase_order");st.text_input("Quote #",key="quote_number");st.text_input("County / State",key="county_state")
        with cd:st.selectbox("Time Zone",TIME_ZONES,index=0,key="time_zone");st.selectbox("Data Deliverables",DATA_DELIVERABLES,key="data_deliverable");st.text_input("Regulatory Program",key="regulatory_program");st.radio("Reportable",["Yes","No"],index=1,horizontal=True,key="reportable")
        ce,cf=st.columns(2)
        with ce:st.selectbox("Turnaround Time",RUSH_OPTIONS,key="rush")
        with cf:st.text_input("DW PWSID # or WW Permit #",key="pwsid");st.radio("Field Filtered?",["Yes","No"],index=1,horizontal=True,key="field_filtered")

    with tab2:
        st.subheader("Select Tests / Analyses")
        st.info("Select tests. Method, preservative, container, hold time auto-determined by water type.")
        wt=st.radio("Water Type",["Potable","Non-Potable"],horizontal=True,key="water_type")
        if "selected_tests_dict" not in st.session_state:st.session_state.selected_tests_dict={}
        sel={}
        for cat,tests in TEST_CATALOGUE.items():
            with st.expander(f"📂 {cat}",expanded=False):
                for tn,ti in tests.items():
                    av=ti.get("m",{})
                    if wt not in av:continue
                    method=av[wt]
                    c1t,c2t,c3t,c4t=st.columns([3,2,2,2])
                    with c1t:checked=st.checkbox(tn,key=f"t_{cat}_{tn}")
                    with c2t:st.caption(f"📋 {method}")
                    with c3t:st.caption(f"🧪 {ti.get('pres','')}")
                    with c4t:st.caption(f"⏱️ {ti.get('hold','')}")
                    if checked:sel[tn]={"method":method,"pres":ti.get("pres",""),"cont":ti.get("cont",""),"hold":ti.get("hold","")}
        st.session_state.selected_tests_dict=sel
        if sel:
            st.success(f"✅ {len(sel)} test(s) selected")
            with st.expander("📋 Summary",expanded=True):
                for n,i in sel.items():st.markdown(f"- **{n}** → {i['method']} | {i['pres']} | {i['hold']}")

    with tab3:
        st.subheader("Sample Information")
        ns=st.number_input("Number of Samples",1,14,1,key="num_samples")
        s2=st.session_state.get("selected_tests_dict",{})
        an=list(s2.keys()) if s2 else ["Analysis"]
        ac=[n[:20]+"..." if len(n)>20 else n for n in an]
        samples=[]
        for i in range(ns):
            with st.expander(f"🧫 Sample {i+1}",expanded=(i==0)):
                c1s,c2s,c3s,c4s=st.columns(4)
                with c1s:sid=st.text_input("Sample ID *",key=f"sid_{i}",placeholder="e.g. TAP-001")
                with c2s:mx=st.selectbox("Matrix",list(MATRIX_CODES.keys()),key=f"mx_{i}")
                with c3s:cg=st.selectbox("Comp/Grab",["Grab","Comp"],key=f"cg_{i}")
                with c4s:nc=st.number_input("# Containers",1,20,1,key=f"nc_{i}")
                c5s,c6s,c7s,c8s=st.columns(4)
                with c5s:cd2=st.date_input("Collection Date",key=f"cd_{i}")
                with c6s:ct2=st.time_input("Collection Time",key=f"ct_{i}")
                with c7s:ed=st.date_input("End Date (composite)",key=f"ed_{i}",value=None)
                with c8s:et=st.time_input("End Time (composite)",key=f"et_{i}",value=None)
                c9s,c10s=st.columns(2)
                with c9s:rcl=st.text_input("Residual Chlorine",key=f"rcl_{i}");rcu=st.selectbox("Units",["","mg/L","ppm"],key=f"rcu_{i}")
                with c10s:cmt=st.text_area("Comment",key=f"cmt_{i}",height=60)
                sa=[]
                if an and an[0]!="Analysis":
                    st.markdown("**Analyses:**")
                    ac2=st.columns(min(len(an),4))
                    for j,a in enumerate(an):
                        with ac2[j%min(len(an),4)]:
                            if st.checkbox(a[:30],value=True,key=f"a_{i}_{j}"):sa.append(a[:20]+"..." if len(a)>20 else a)
                else:sa=ac
                samples.append({"sample_id":sid,"matrix":MATRIX_CODES.get(mx,"DW"),"comp_grab":cg[:4].upper(),"start_date":cd2.strftime("%m/%d/%y") if cd2 else "","start_time":ct2.strftime("%H:%M") if ct2 else "","end_date":ed.strftime("%m/%d/%y") if ed else "","end_time":et.strftime("%H:%M") if et else "","num_containers":str(nc),"res_cl_result":rcl,"res_cl_units":rcu,"analyses":sa,"comment":cmt})
        st.session_state["samples_data"]=samples;st.session_state["analysis_columns"]=ac

    with tab4:
        st.subheader("Additional Information")
        ca4,cb4=st.columns(2)
        with ca4:st.text_area("Additional Instructions from KELP",key="additional_instructions",height=80);st.text_input("# Coolers",key="num_coolers");st.text_input("Thermometer ID",key="thermometer_id");st.text_input("Temperature (°C)",key="temperature")
        with cb4:st.text_area("Customer Remarks / Hazards",key="customer_remarks",height=80);st.radio("Sample Received on Ice?",["Yes","No"],horizontal=True,key="received_on_ice");st.text_input("Tracking Number",key="tracking_number");st.selectbox("Delivery Method",["FedEX","UPS","In-Person","Courier","Other"],key="delivery_method")
        st.divider();st.subheader("KELP Internal Use")
        cc4,cd4=st.columns(2)
        with cc4:st.text_input("Project Manager",key="project_manager");st.text_input("AcctNum / Client ID",key="acct_num");st.text_input("Table #",key="table_number")
        with cd4:st.text_input("Profile / Template",key="profile_template");st.text_input("Prelog / Bottle Ord. ID",key="prelog_id");st.text_input("KELP Ordering ID",key="kelp_ordering_id",placeholder="KELP-MMDDYY-####")

    with tab5:
        st.subheader("Generate & Download CoC PDF")
        s5=st.session_state.get("selected_tests_dict",{});sm5=st.session_state.get("samples_data",[]);ac5=st.session_state.get("analysis_columns",[])
        warns=[]
        if not st.session_state.get("company_name"):warns.append("Company Name required")
        if not st.session_state.get("contact_name"):warns.append("Contact/Report To required")
        if not st.session_state.get("email"):warns.append("E-Mail required")
        if not s5:warns.append("No tests selected")
        if not sm5 or not any(s.get("sample_id") for s in sm5):warns.append("Sample ID required")
        for w in warns:st.warning(f"⚠️ {w}")
        with st.expander("📋 Preview",expanded=True):
            p1,p2=st.columns(2)
            with p1:st.markdown(f"**Company:** {st.session_state.get('company_name','—')}");st.markdown(f"**Contact:** {st.session_state.get('contact_name','—')}");st.markdown(f"**Samples:** {len(sm5)}")
            with p2:st.markdown(f"**Tests:** {len(s5)}");st.markdown(f"**Water Type:** {st.session_state.get('water_type','Potable')}");st.markdown(f"**TAT:** {st.session_state.get('rush','Standard')}")
        am=""
        st.divider()
        if st.button("🖨️ Generate CoC PDF",type="primary",use_container_width=True):
            d={k:st.session_state.get(k,"") for k in ["company_name","street_address","phone","email","cc_email","contact_name","project_number","project_name","invoice_to","invoice_email","site_info","purchase_order","quote_number","county_state","time_zone","data_deliverable","regulatory_program","reportable","rush","pwsid","field_filtered","additional_instructions","customer_remarks","num_coolers","thermometer_id","temperature","received_on_ice","tracking_number","delivery_method","project_manager","acct_num","table_number","profile_template","prelog_id","kelp_ordering_id"]}
            d["analysis_method"]=am;d["samples"]=sm5;d["selected_tests"]=s5;d["analysis_columns"]=ac5
            logo=os.path.join(os.path.dirname(os.path.abspath(__file__)),"kelp_logo.png")
            with st.spinner("Generating..."):pdf=generate_coc_pdf(d,logo_path=logo)
            st.success("✅ CoC PDF generated!")
            co=d["company_name"].replace(" ","_")[:20] or "KELP";ts=datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button("📥 Download CoC PDF",data=pdf,file_name=f"KELP_CoC_{co}_{ts}.pdf",mime="application/pdf",type="primary",use_container_width=True)
            st.info("💡 Print, sign, and include with your sample shipment to KELP.")

if __name__=="__main__":main()

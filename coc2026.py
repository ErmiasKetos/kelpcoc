"""KELP Chain-of-Custody Streamlit Application - v10
Complete rebuild matching original template layout exactly.
Key: Sidebar fields (Project Mgr, AcctNum, etc.) sit alongside
     tall analysis column headers, NOT in form header.
"""
import streamlit as st, io, os, datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader

# ═══ CONSTANTS ═══
PW,PH=landscape(letter)
ML,MR,MT,MB=15,15,12,12
CW=PW-ML-MR
HDR_BG=HexColor("#1F4E79")
KELP_TEAL=HexColor("#008080")
LIGHT_BLUE=HexColor("#DCE6F1")

# ═══ TEST CATALOGUE ═══
TEST_CATALOGUE={
    "PHYSICAL/GENERAL CHEMISTRY":["pH","Conductivity","Turbidity","TDS","TSS","Color","Odor","Temperature","Dissolved Oxygen","Hardness (SM2340C)","Alkalinity (SM2320B)","Specific Conductance"],
    "METALS":["Arsenic","Barium","Cadmium","Chromium","Copper","Iron","Lead","Manganese","Mercury","Nickel","Selenium","Silver","Zinc","Aluminum","Antimony","Beryllium","Boron","Calcium","Cobalt","Magnesium","Molybdenum","Potassium","Sodium","Strontium","Thallium","Tin","Titanium","Uranium","Vanadium"],
    "INORGANICS":["Chloride","Fluoride","Sulfate","Nitrate (as N)","Nitrite (as N)","Bromide","Phosphate"],
    "NUTRIENTS":["Total Nitrogen","TKN","Ammonia","Total Phosphorus","Orthophosphate","BOD","COD","TOC"],
    "ORGANICS":["TPH-GRO","TPH-DRO","BTEX","PAHs","PCBs","Pesticides","Herbicides","SVOCs","VOCs"],
    "DISINFECTION":["Total Coliform","E. coli","Heterotrophic Plate Count","Chlorine Residual","Bromate","Chlorite","Haloacetic Acids","Total Trihalomethanes"],
    "PFAS TESTING":["PFAS EPA 537.1","PFAS EPA 1633A","PFOS","PFOA","GenX"],
    "PACKAGES":["Complete Homeowner","Mortgage/Real Estate","Basic Potability","Title 22","Ag Irrigation"],
}
CAT_SHORT = {
    "PHYSICAL/GENERAL CHEMISTRY": "Physical/Gen Chem",
    "METALS": "Metals", "INORGANICS": "Inorganics", "NUTRIENTS": "Nutrients",
    "ORGANICS": "Organics", "DISINFECTION": "Disinfection",
    "PFAS TESTING": "PFAS", "PACKAGES": "Packages",
}
MATRIX_CODES = {"Drinking Water (DW)":"DW","Ground Water (GW)":"GW","Wastewater (WW)":"WW","Surface Water (SW)":"SW","Product (P)":"P","Other (OT)":"OT"}
DATA_DELIVERABLES = ["Level I (Std)","Level II","Level III","Level IV","Other"]
RUSH_OPTIONS = ["Standard (5-10 Day)","Same Day","1 Day","2 Day","3 Day","4 Day","5 Day"]
TIME_ZONES = ["PT","MT","CT","ET","AK"]
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kelp_logo.png")

# ═══ DRAWING HELPERS ═══
def _cb(c,x,y,checked=False,sz=7):
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,sz,sz,fill=0,stroke=1)
    if checked:
        c.setFont("Helvetica-Bold",sz);c.setFillColor(black);c.drawString(x+1,y+1,"X")

def _cell(c,x,y,w,h,text="",fs=6,al="left",bold=False,bg=None):
    if bg:
        c.setFillColor(bg);c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,w,h,fill=0,stroke=1)
    if text:
        fn="Helvetica-Bold" if bold else "Helvetica"
        c.setFont(fn,fs);c.setFillColor(black)
        if al=="center":c.drawCentredString(x+w/2,y+(h-fs)/2+1,str(text))
        elif al=="right":c.drawRightString(x+w-2,y+(h-fs)/2+1,str(text))
        else:c.drawString(x+2,y+(h-fs)/2+1,str(text))

def _lcell(c,x,y,w,h,label,value="",lfs=6,vfs=7):
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,w,h,fill=0,stroke=1)
    c.setFont("Helvetica",lfs);c.setFillColor(black);c.drawString(x+2,y+(h-lfs)/2+1,label)
    if value:
        lw=c.stringWidth(label,"Helvetica",lfs)+4
        c.setFont("Helvetica",vfs);c.drawString(x+lw,y+(h-vfs)/2+1,str(value))

def _hcell(c,x,y,w,h,text="",fs=6):
    """Header cell - NO fill, black text"""
    c.setStrokeColor(black);c.setLineWidth(0.4);c.rect(x,y,w,h,fill=0,stroke=1)
    if text:
        c.setFont("Helvetica-Bold",fs);c.setFillColor(black)
        lines = str(text).split('\n')
        if len(lines) == 1:
            c.drawCentredString(x+w/2,y+(h-fs)/2+1,text)
        else:
            for i, ln in enumerate(lines):
                ly = y + h - fs*1.2 - i*fs*1.2
                c.drawCentredString(x+w/2, ly, ln)

# ═══ PDF GENERATION ═══
def generate_coc_pdf(data, logo_path=None):
    buf=io.BytesIO()
    c=canvas.Canvas(buf,pagesize=landscape(letter))
    c.setTitle("KELP Chain-of-Custody")
    y=PH-MT

    # ═══ COLUMN LAYOUT (scaled from original 769px total) ═══
    # Original: x=17..786 = 769px
    # Left fixed columns: 459px (Customer Sample ID thru Res Cl Units)
    # 10 Analysis cols: 190px (19px each)
    # Sample Comment: 89px (but in our layout this becomes part of sidebar)
    # Pres NC: 23px
    # Sidebar (KELP Use Only): x=665..763 = ~98px + Pres NC 23px = 121px

    # In original, the sidebar right area = x=665..786 = 121px
    # Left table area = x=17..665 = 648px
    # Analysis columns within left area: x=476..665 = 189px (10 cols)
    # Left fixed columns: x=17..476 = 459px

    OW = 769.0
    TW = CW  # total width
    sf = TW / OW

    # Sidebar area (KELP Use Only + Sample Comment + Pres NC)
    sb_total = round(121 * sf)  # total sidebar width
    pnc_w = round(23 * sf)     # Preservation NC column
    kelp_sb_w = round(15 * sf) # "KELP Use Only" rotated label strip
    sb_fields_w = sb_total - pnc_w - kelp_sb_w  # fields area (Project Mgr etc.)

    # Main content area (left of sidebar)
    main_w = TW - sb_total

    # Left fixed columns
    cid  = round(161 * sf)
    cmx  = round(32 * sf)
    ccg  = round(27 * sf)
    csd  = round(55 * sf)
    cst  = round(35 * sf)
    ced  = round(50 * sf)
    cet  = round(35 * sf)
    cnc  = round(20 * sf)
    crr  = round(23 * sf)
    cru  = round(21 * sf)
    fixed_total = cid+cmx+ccg+csd+cst+ced+cet+cnc+crr+cru

    # Analysis columns
    NUM_ANALYSIS_COLS = 10
    acw_each = round(19 * sf)
    analysis_total = acw_each * NUM_ANALYSIS_COLS

    # Adjust: absorb rounding into cid
    left_used = fixed_total + analysis_total
    if left_used != main_w:
        cid += (main_w - left_used)
        fixed_total = cid+cmx+ccg+csd+cst+ced+cet+cnc+crr+cru

    # X position where analysis columns start
    analysis_start_x = ML + fixed_total

    # ═══ HEADER ═══
    hh=42
    c.setStrokeColor(KELP_TEAL);c.setLineWidth(2);c.line(ML,y,ML+TW,y)
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(ImageReader(logo_path),ML+3,y-hh+4,width=105,height=hh-8,preserveAspectRatio=True,mask='auto')
        except: pass
    c.setFont("Helvetica-Bold",13);c.setFillColor(black)
    c.drawCentredString(ML+main_w/2,y-16,"CHAIN-OF-CUSTODY")
    c.setFont("Helvetica",7);c.drawCentredString(ML+main_w/2,y-27,"Chain-of-Custody is a LEGAL DOCUMENT - Complete all relevant fields")

    # KELP USE ONLY header (top right) - original structure
    kx = ML + main_w
    c.setStrokeColor(black);c.setLineWidth(0.5);c.rect(kx,y-hh,sb_total,hh,fill=0,stroke=1)
    c.setFont("Helvetica-Bold",7);c.setFillColor(black)
    c.drawCentredString(kx+sb_total/2,y-10,"KELP USE ONLY- Affix Workorder Label Here")
    c.setFont("Helvetica",6);c.drawString(kx+4,y-22,"KELP Ordering ID:")
    c.setLineWidth(0.3);c.line(kx+68,y-24,kx+sb_total-4,y-24)
    kid=data.get("kelp_ordering_id","")
    if kid:
        c.setFont("Helvetica-Bold",7);c.drawString(kx+70,y-22,kid)
    c.setFont("Helvetica",5);c.drawString(kx+4,y-32,"Format: KELP-MMDDYY-####")
    c.setStrokeColor(KELP_TEAL);c.setLineWidth(1.5);c.line(ML,y-hh,ML+TW,y-hh)
    y-=hh

    # ═══ CLIENT INFO ROWS ═══
    rh=12
    lw=230
    cw2=main_w-lw  # center column width

    # Rows 1-5: client info (left + center)
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
    # Sidebar: KELP USE ONLY blank label area spanning all client rows
    c.setStrokeColor(black);c.setLineWidth(0.4)
    c.rect(ML+main_w, y, sb_total, rh*5, fill=0, stroke=1)

    # Row 6: Project Name | Invoice E-mail | Specify Container Size
    _lcell(c,ML,y-rh,lw,rh,"Project Name: ",data.get("project_name",""))
    inv_w = cw2*0.55
    scs_w = cw2 - inv_w  # remainder for "Specify Container Size"
    _lcell(c,ML+lw,y-rh,inv_w,rh,"Invoice E-mail: ",data.get("invoice_email",""))
    _cell(c,ML+lw+inv_w,y-rh,scs_w,rh)
    c.setFont("Helvetica-Bold",6);c.setFillColor(black)
    c.drawCentredString(ML+lw+inv_w+scs_w/2,y-rh+3,"Specify Container Size")
    # Container Size legend in sidebar
    _cell(c,ML+main_w,y-rh*2,sb_total,rh*2)
    c.setFont("Helvetica-Bold",5);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-4,"Container Size: (1) 1L, (2) 500mL, (3) 250mL,")
    c.drawString(ML+main_w+3,y-11,"(4) 125mL, (5) 100mL, (6) Other")
    y-=rh

    # Row 7: Site Collection | Purchase Order | Identify Container Preservative Type
    _lcell(c,ML,y-rh,lw,rh,"Site Collection Info/Facility ID (as applicable): ",data.get("site_info",""),lfs=5)
    _lcell(c,ML+lw,y-rh,inv_w,rh,"Purchase Order (if applicable): ",data.get("purchase_order",""),lfs=5.5)
    _cell(c,ML+lw+inv_w,y-rh,scs_w,rh)
    c.setFont("Helvetica-Bold",6);c.setFillColor(black)
    c.drawCentredString(ML+lw+inv_w+scs_w/2,y-rh+3,"Identify Container Preservative Type")
    # Preservative legend (spans 3 rows in sidebar)
    _cell(c,ML+main_w,y-rh*3,sb_total,rh*3)
    c.setFont("Helvetica-Bold",5);c.setFillColor(black)
    c.drawString(ML+main_w+3,y-4,"Preservative Types: (1) None, (2) HNO3, (3)")
    c.drawString(ML+main_w+3,y-11,"H2SO4, (4) HCl, (5) NaOH, (6) Zn Acetate, (7)")
    c.drawString(ML+main_w+3,y-18,"NaHSO4, (8) Sod.Thiosulfate, (9) Ascorbic Acid,")
    c.drawString(ML+main_w+3,y-25,"(10) MeOH, (11) Other")
    y-=rh

    # Row 8: (blank left) | Quote # | Analysis Requested
    _cell(c,ML,y-rh,lw,rh)
    _lcell(c,ML+lw,y-rh,inv_w,rh,"Quote #: ",data.get("quote_number",""))
    _cell(c,ML+lw+inv_w,y-rh,scs_w,rh)
    c.setFont("Helvetica-Bold",6);c.setFillColor(black)
    c.drawCentredString(ML+lw+inv_w+scs_w/2,y-rh+3,"Analysis Requested")
    y-=rh

    # ═══ TIME ZONE / COUNTY ROW ═══
    tzh=11
    _cell(c,ML,y-tzh,lw,tzh)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+2,y-tzh+3,"Sample Collection Time Zone :")
    st_tz=data.get("time_zone","PT")
    for i,tz in enumerate(["AK","PT","MT","CT","ET"]):
        cx2=ML+115+i*22
        _cb(c,cx2,y-tzh+2,checked=(tz==st_tz),sz=6)
        c.setFont("Helvetica",6);c.setFillColor(black);c.drawString(cx2+8,y-tzh+3,tz)
    _lcell(c,ML+lw,y-tzh,main_w-lw,tzh,"County / State origin of sample(s): ",data.get("county_state",""),lfs=5.5)
    # Sidebar row (preservative legend continues)
    y-=tzh

    # ═══ DATA DELIVERABLES / REGULATORY / RUSH ═══
    rr=11
    dd_w=130
    reg_w=main_w-dd_w
    sd=data.get("data_deliverable","Level I (Std)")
    sr=data.get("rush","Standard (5-10 Day)")

    # Data Deliverables spanning 4 rows
    _cell(c,ML,y-rr*4,dd_w,rr*4)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+2,y-5,"Data Deliverables:")
    dd_items=[("Level I (Std)",0),("Level II",0),("Level III",1),("Level IV",2),("Other",3)]
    for label,row_idx in dd_items:
        if row_idx<2:
            bx=ML+4 if label=="Level I (Std)" else ML+60; by=y-14
        else:
            bx=ML+4; by=y-14-row_idx*rr
        _cb(c,bx,by,checked=(label==sd),sz=6)
        c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(bx+8,by+1,label)

    # Row 1: Regulatory Program + Reportable
    _cell(c,ML+dd_w,y-rr,reg_w,rr)
    c.setFont("Helvetica",6);c.setFillColor(black)
    c.drawString(ML+dd_w+2,y-rr+3,f"Regulatory Program (DW, RCRA, etc.) as applicable: {data.get('regulatory_program','')}")
    rpx=ML+dd_w+reg_w*0.60
    c.drawString(rpx,y-rr+3,"Reportable")
    rv=data.get("reportable","No")=="Yes"
    _cb(c,rpx+40,y-rr+2,checked=rv,sz=7);c.drawString(rpx+49,y-rr+3,"Yes")
    _cb(c,rpx+68,y-rr+2,checked=not rv,sz=7);c.drawString(rpx+77,y-rr+3,"No")

    # Row 2: Rush
    _cell(c,ML+dd_w,y-rr*2,reg_w,rr)
    c.setFont("Helvetica-Bold",6);c.setFillColor(black)
    c.drawString(ML+dd_w+2,y-rr*2+3,"Rush (Pre-approval required):")
    for i,ro in enumerate(["Same Day","1 Day","2 Day","3 Day","4 Day"]):
        rx=ML+dd_w+130+i*42
        _cb(c,rx,y-rr*2+2,checked=(ro==sr),sz=7)
        c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(rx+9,y-rr*2+3,ro)

    # Row 3: 5 Day / Other | DW PWSID
    half_reg=reg_w*0.55
    _cell(c,ML+dd_w,y-rr*3,half_reg,rr)
    _cb(c,ML+dd_w+4,y-rr*3+2,checked=("5 Day"==sr),sz=7)
    c.setFont("Helvetica",5.5);c.setFillColor(black);c.drawString(ML+dd_w+14,y-rr*3+3,"5 Day")
    c.drawString(ML+dd_w+50,y-rr*3+3,"Other ___________")
    _cell(c,ML+dd_w+half_reg,y-rr*3,reg_w-half_reg,rr)
    c.drawString(ML+dd_w+half_reg+2,y-rr*3+3,f"DW PWSID # or WW Permit # as applicable: {data.get('pwsid','')}")

    # Row 4: Other | Field Filtered
    half_reg2=reg_w*0.40
    ff_w=reg_w*0.60
    _cell(c,ML+dd_w,y-rr*4,half_reg2,rr)
    c.setFont("Helvetica",5.5);c.setFillColor(black)
    c.drawString(ML+dd_w+2,y-rr*4+3,"Other ___________")
    _cell(c,ML+dd_w+half_reg2,y-rr*4,ff_w,rr)
    c.drawString(ML+dd_w+half_reg2+2,y-rr*4+3,"Field Filtered (if applicable):")
    ff=data.get("field_filtered","No")
    ffx=ML+dd_w+half_reg2+100
    _cb(c,ffx,y-rr*4+2,checked=(ff=="Yes"),sz=7);c.drawString(ffx+9,y-rr*4+3,"Yes")
    _cb(c,ffx+28,y-rr*4+2,checked=(ff=="No"),sz=7);c.drawString(ffx+37,y-rr*4+3,"No")

    # Sidebar for Deliverables rows (blank - preservative legend area continues / ends)
    c.setStrokeColor(black);c.setLineWidth(0.4)
    c.rect(ML+main_w,y-rr*4,sb_total,rr*4,fill=0,stroke=1)

    y-=rr*4

    # ═══ MATRIX LEGEND ═══
    lgh=9
    _cell(c,ML,y-lgh,main_w+sb_total,lgh)
    c.setFont("Helvetica",5);c.setFillColor(black)
    c.drawString(ML+2,y-lgh+2,"* Insert in Matrix box below: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water (SW), Other (OT)")
    y-=lgh

    # ═══════════════════════════════════════════════════════════════
    # ═══ SAMPLE TABLE WITH INTEGRATED SIDEBAR ═══
    # This is the key structure from the screenshot:
    # LEFT side: fixed columns (Customer Sample ID thru Res Cl Units)
    #   - Column headers are ~33pt tall (bottom portion)
    # CENTER: 10 analysis columns
    #   - Column headers are TALL (~130pt) with vertical rotated text
    #   - Above the headers: 3 rows (Container Size, Preservative, Analysis) - 
    #     BUT those are already drawn above in the form, so here we just have
    #     the tall vertical analysis headers + blank cells for container/preservative numbers
    # RIGHT sidebar: KELP Use Only strip + Project Mgr, AcctNum, Table#, Profile, Prelog, Sample Comment + Pres NC
    # ═══════════════════════════════════════════════════════════════

    samples=data.get("samples",[])
    acols=data.get("analysis_columns",[])

    # Heights
    srh = 13  # sample data row height
    max_rows = 10
    left_hdr_h = 33  # left column header height (Customer Sample ID etc.)
    
    # The tall analysis header spans: 3 label rows + container/preservative blank rows + left header
    # From screenshot: sidebar has ~6 fields (Project Mgr thru Sample Comment)
    # Each field row ~18-20pt, plus Sample Comment is taller
    # Total sidebar height = 5 fields * 18pt + Sample Comment ~33pt = ~123pt
    # This equals the tall analysis header height
    
    sb_field_h = 18  # height per sidebar field row
    sb_comment_h = 33  # Sample Comment row height
    tall_hdr_h = sb_field_h * 5 + sb_comment_h  # = 123pt total
    # The left headers only occupy the bottom left_hdr_h of this height
    # The top portion (tall_hdr_h - left_hdr_h = 90pt) has no left-side headers
    
    table_top = y  # top of the entire table+sidebar block
    
    # ─── Draw 10 Analysis Column Headers (tall vertical) ───
    ax = analysis_start_x
    for ai in range(NUM_ANALYSIS_COLS):
        label = acols[ai] if ai < len(acols) else ""
        # Full tall cell
        c.setStrokeColor(black);c.setLineWidth(0.4)
        c.rect(ax, y-tall_hdr_h, acw_each, tall_hdr_h, fill=0, stroke=1)
        if label:
            c.saveState()
            fs2 = 6 if len(label) < 18 else 5.5 if len(label) < 22 else 5
            c.setFont("Helvetica-Bold",fs2);c.setFillColor(black)
            c.translate(ax+acw_each/2+1.5, y-tall_hdr_h+3)
            c.rotate(90)
            c.drawString(0,0,label[:30])
            c.restoreState()
        ax += acw_each

    # ─── Draw Left Fixed Column Headers (bottom portion only) ───
    hdr_bottom_y = y - tall_hdr_h  # bottom of headers = top of data rows
    left_hdr_top = hdr_bottom_y + left_hdr_h  # left headers start here
    
    # Top blank area above left headers (tall_hdr_h - left_hdr_h)
    blank_top_h = tall_hdr_h - left_hdr_h
    c.setStrokeColor(black);c.setLineWidth(0.4)
    c.rect(ML, y-blank_top_h, fixed_total, blank_top_h, fill=0, stroke=1)
    
    # Left column headers in bottom portion
    cx = ML
    _hcell(c,cx,hdr_bottom_y,cid,left_hdr_h,"Customer Sample ID",fs=6)
    cx+=cid
    _hcell(c,cx,hdr_bottom_y,cmx,left_hdr_h,"Matrix\n*",fs=5)
    cx+=cmx
    _hcell(c,cx,hdr_bottom_y,ccg,left_hdr_h,"Comp /\nGrab",fs=5)
    cx+=ccg
    
    # Composite Start (grouped)
    csw=csd+cst
    h1=14
    _hcell(c,cx,hdr_bottom_y+left_hdr_h-h1,csw,h1,"Composite  Start",fs=5)
    _hcell(c,cx,hdr_bottom_y,csd,left_hdr_h-h1,"Date",fs=5)
    _hcell(c,cx+csd,hdr_bottom_y,cst,left_hdr_h-h1,"Time",fs=5)
    cx+=csw
    
    # Collected/End
    cew=ced+cet
    _hcell(c,cx,hdr_bottom_y+left_hdr_h-h1,cew,h1,"Collected or Composite\nEnd",fs=4.5)
    _hcell(c,cx,hdr_bottom_y,ced,left_hdr_h-h1,"Date",fs=5)
    _hcell(c,cx+ced,hdr_bottom_y,cet,left_hdr_h-h1,"Time",fs=4.5)
    cx+=cew
    
    _hcell(c,cx,hdr_bottom_y,cnc,left_hdr_h,"#\nCont.",fs=5)
    cx+=cnc
    
    # Residual Chlorine
    rcw=crr+cru
    _hcell(c,cx,hdr_bottom_y+left_hdr_h-h1,rcw,h1,"Residual\nChlorine",fs=5)
    _hcell(c,cx,hdr_bottom_y,crr,left_hdr_h-h1,"Result",fs=5)
    _hcell(c,cx+crr,hdr_bottom_y,cru,left_hdr_h-h1,"Units",fs=5)
    
    # ─── Draw Right Sidebar (alongside tall analysis headers) ───
    sx = ML + main_w  # sidebar start x
    
    # "KELP Use Only" rotated strip
    c.setStrokeColor(black);c.setLineWidth(0.4)
    c.rect(sx, y-tall_hdr_h, kelp_sb_w, tall_hdr_h, fill=0, stroke=1)
    c.saveState()
    c.setFont("Helvetica-Bold",6);c.setFillColor(black)
    mid_y = y - tall_hdr_h/2
    c.translate(sx+kelp_sb_w/2+2, mid_y-25)
    c.rotate(90)
    c.drawString(0,0,"KELP Use Only")
    c.restoreState()
    
    # Sidebar fields
    fx = sx + kelp_sb_w
    fy = y  # start from top
    
    sb_fields = [
        ("Project Mgr.: ", data.get("project_manager","")),
        ("AcctNum / Client ID: ", data.get("acct_num","")),
        ("Table #: ", data.get("table_number","")),
        ("Profile / Template: ", data.get("profile_template","")),
        ("Prelog / Bottle Ord. ID: ", data.get("prelog_id","")),
    ]
    for label, val in sb_fields:
        _lcell(c, fx, fy-sb_field_h, sb_fields_w, sb_field_h, label, val, lfs=5.5, vfs=6)
        fy -= sb_field_h
    
    # Sample Comment (taller row)
    _cell(c, fx, fy-sb_comment_h, sb_fields_w, sb_comment_h)
    c.setFont("Helvetica-Bold",7);c.setFillColor(black)
    c.drawCentredString(fx+sb_fields_w/2, fy-sb_comment_h/2-2, "Sample Comment")
    
    # Preservation NC column (far right, full height of tall header + all data rows)
    pnc_x = sx + kelp_sb_w + sb_fields_w
    total_table_h = tall_hdr_h + max_rows * srh
    c.setStrokeColor(black);c.setLineWidth(0.4)
    c.rect(pnc_x, y-total_table_h, pnc_w, total_table_h, fill=0, stroke=1)
    # Rotated text
    c.saveState()
    c.setFont("Helvetica",4);c.setFillColor(black)
    c.translate(pnc_x+pnc_w/2+2, y-total_table_h+3)
    c.rotate(90)
    c.drawString(0,0,"Preservation non-conformance identified for sample.")
    c.restoreState()
    
    y -= tall_hdr_h  # move y to top of data rows
    
    # ═══ SAMPLE DATA ROWS ═══
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
        
        # 10 Analysis columns
        s_analyses=s.get("analyses",[])
        for ai in range(NUM_ANALYSIS_COLS):
            ac_name=acols[ai] if ai<len(acols) else ""
            chk="X" if ac_name and ac_name in s_analyses else ""
            _cell(c,cx,y-srh,acw_each,srh,chk,fs=7,al="center",bold=True)
            cx+=acw_each
        
        # Sample Comment (in sidebar area, below the comment header)
        cmt=s.get("comment","")
        _cell(c, sx+kelp_sb_w, y-srh, sb_fields_w, srh, cmt, fs=5.5)
        
        # Pres NC data cell (already covered by the big rect, draw divider)
        c.setStrokeColor(black);c.setLineWidth(0.4)
        c.line(pnc_x, y-srh, pnc_x+pnc_w, y-srh)
        
        # KELP Use Only strip data row divider  
        c.line(sx, y-srh, sx+kelp_sb_w, y-srh)
        
        y-=srh

    # ═══ CUSTOMER REMARKS ═══
    rmh=30
    _cell(c,ML,y-rmh,TW,rmh)
    c.setFont("Helvetica-Bold",7);c.setFillColor(black)
    c.drawString(ML+2,y-8,"Customer Remarks / Special Conditions / Possible Hazards:")
    rmk=data.get("customer_remarks","")
    if rmk:
        c.setFont("Helvetica",6);c.drawString(ML+4,y-18,rmk[:150])
    c.drawString(ML+2,y-rmh+3,"Additional Instructions: "+data.get("additional_instructions","")[:100])
    y-=rmh

    # ═══ RELINQUISHED/RECEIVED BLOCK ═══
    rrh=14
    c.setFont("Helvetica-Bold",7);c.setFillColor(black)
    half=TW/2
    labels=[
        ("Relinquished by (signature / print):","Date:","Time:"),
        ("Received by (signature / print):","Date:","Time:"),
    ]
    for i in range(3):
        for j,(l1,l2,l3) in enumerate(labels):
            bx=ML+j*half
            _cell(c,bx,y-rrh,half*0.50,rrh)
            _cell(c,bx+half*0.50,y-rrh,half*0.25,rrh)
            _cell(c,bx+half*0.75,y-rrh,half*0.25,rrh)
            c.setFont("Helvetica",5.5);c.setFillColor(black)
            c.drawString(bx+2,y-rrh+4,l1)
            c.drawString(bx+half*0.50+2,y-rrh+4,l2)
            c.drawString(bx+half*0.75+2,y-rrh+4,l3)
        y-=rrh

    # ═══ LAB RECEIVING (bottom) ═══
    lrh=35
    _cell(c,ML,y-lrh,TW,lrh)
    c.setFont("Helvetica-Bold",8);c.setFillColor(black)
    c.drawString(ML+2,y-8,"Lab Receiving (Cooler / Shipping Conditions)")
    fields=[
        (f"No. of Coolers: {data.get('num_coolers','')}",0.18),
        (f"Thermometer ID: {data.get('thermometer_id','')}",0.22),
        (f"Temperature (\u00b0C): {data.get('temperature','')}",0.18),
    ]
    fx2=ML+4; fy2=y-18
    for txt,pct in fields:
        w2=TW*pct
        c.setFont("Helvetica",6);c.drawString(fx2,fy2,txt)
        fx2+=w2
    # Received on ice
    c.drawString(fx2,fy2,"Received on ice: ")
    ri_val=data.get("received_on_ice","Yes")
    _cb(c,fx2+55,fy2-1,checked=(ri_val=="Yes"),sz=7);c.drawString(fx2+64,fy2,"Yes")
    _cb(c,fx2+80,fy2-1,checked=(ri_val=="No"),sz=7);c.drawString(fx2+89,fy2,"No")
    _cb(c,fx2+105,fy2-1,checked=(ri_val=="N/A"),sz=7);c.drawString(fx2+114,fy2,"N/A")
    # Bottom row
    c.setFont("Helvetica",6)
    c.drawString(ML+4,y-lrh+4,f"Tracking #: {data.get('tracking_number','')}")
    c.drawString(ML+TW*0.35,y-lrh+4,f"Delivery Method: {data.get('delivery_method','')}")
    c.drawString(ML+TW*0.65,y-lrh+4,"Custody Seal Intact:  \u2610 Yes  \u2610 No  \u2610 N/A")

    c.showPage();c.save();buf.seek(0);return buf


# ═══ STREAMLIT UI ═══
def main():
    st.set_page_config(page_title="KELP Chain-of-Custody",layout="wide")
    st.title("KELP Chain-of-Custody Generator")

    c1,c2=st.columns(2)
    with c1:
        st.subheader("Client Information")
        cn=st.text_input("Company Name","Luna Owners Association")
        sa=st.text_input("Street Address","")
        pn=st.text_input("Customer Project #","12345")
        prn=st.text_input("Project Name","03-05-2025")
        si=st.text_input("Site Collection Info","")
    with c2:
        ct=st.text_input("Contact/Report To","ermias@ketos.co")
        ph=st.text_input("Phone #","")
        em=st.text_input("E-Mail","")
        cem=st.text_input("Cc E-Mail","")
        it=st.text_input("Invoice to","")
        ie=st.text_input("Invoice E-mail","")

    st.subheader("Test Delivery & Regulatory")
    c3,c4,c5=st.columns(3)
    with c3:
        dd=st.selectbox("Data Deliverables",DATA_DELIVERABLES)
        tz=st.selectbox("Time Zone",TIME_ZONES,index=1)
    with c4:
        rush=st.selectbox("Rush",RUSH_OPTIONS)
        rp=st.text_input("Regulatory Program","")
    with c5:
        rep=st.radio("Reportable",["Yes","No"],index=1,horizontal=True)
        ff=st.radio("Field Filtered",["Yes","No"],index=1,horizontal=True)
        pw=st.text_input("DW PWSID / WW Permit #","")

    st.subheader("Additional Info")
    c6,c7=st.columns(2)
    with c6:
        cs=st.text_input("County / State origin","")
        po=st.text_input("Purchase Order","")
        qn=st.text_input("Quote #","")
    with c7:
        pm=st.text_input("Project Manager","")
        an2=st.text_input("AcctNum / Client ID","")
        tn=st.text_input("Table #","")
        pt=st.text_input("Profile / Template","")
        pl=st.text_input("Prelog / Bottle Ord. ID","")
        koid=st.text_input("KELP Ordering ID","")

    # ═══ TESTS & SAMPLES ═══
    st.subheader("Select Tests")
    s2={}
    for cat,tests in TEST_CATALOGUE.items():
        with st.expander(cat):
            for t in tests:
                if st.checkbox(t,key=f"test_{t}"):
                    s2[t]={"method":cat}

    # Group selected tests by category
    cat_tests={}
    for tname in s2:
        found_cat=None
        for cat,tests in TEST_CATALOGUE.items():
            if tname in tests:
                found_cat=cat;break
        if found_cat is None:found_cat="OTHER"
        cat_tests.setdefault(found_cat,[]).append(tname)

    ac=[CAT_SHORT.get(cat,cat[:18]) for cat in cat_tests.keys()]
    test_to_cat={}
    for cat,tnames in cat_tests.items():
        for tn in tnames:
            test_to_cat[tn]=CAT_SHORT.get(cat,cat[:18])

    st.subheader("Sample Details")
    n_samples=st.number_input("Number of Samples",1,10,2)
    mx=st.selectbox("Default Matrix",list(MATRIX_CODES.keys()))
    mx_code=MATRIX_CODES[mx]
    cg=st.selectbox("Comp/Grab",["GRAB","COMP"])

    sample_list=[]
    for i in range(n_samples):
        with st.expander(f"Sample {i+1}",expanded=(i<2)):
            sid=st.text_input("Sample ID",f"Sample-{i+1}",key=f"sid_{i}")
            an=list(s2.keys())
            sa_cats=set()
            for j,a in enumerate(an):
                if st.checkbox(a[:30],value=True,key=f"a_{i}_{j}"):
                    cat_name=test_to_cat.get(a,a)
                    sa_cats.add(cat_name)
            sa_list=list(sa_cats)
            now=datetime.datetime.now()
            sd2=st.text_input("Collection Date",now.strftime("%m/%d/%y"),key=f"sd_{i}")
            st2=st.text_input("Collection Time",now.strftime("%H:%M"),key=f"st_{i}")
            cmt=st.text_input("Comment","",key=f"cmt_{i}")
            sample_list.append({"sample_id":sid,"matrix":mx_code,"comp_grab":cg,
                "start_date":sd2,"start_time":st2,"end_date":sd2,"end_time":st2,
                "num_containers":"1","res_cl_result":"","res_cl_units":"",
                "analyses":sa_list,"comment":cmt})

    # ═══ SHIPPING ═══
    st.subheader("Shipping / Receiving")
    c8,c9=st.columns(2)
    with c8:
        nc=st.text_input("# Coolers","1")
        tid=st.text_input("Thermometer ID","")
        tmp=st.text_input("Temperature (°C)","")
    with c9:
        roi=st.radio("Received on Ice",["Yes","No","N/A"],horizontal=True)
        trk=st.text_input("Tracking #","")
        dm=st.selectbox("Delivery Method",["FedEX","UPS","USPS","Hand Delivery","Courier","Other"])

    rmk=st.text_area("Customer Remarks / Special Conditions","")
    ai2=st.text_input("Additional Instructions","")

    if st.button("Generate CoC PDF",type="primary"):
        d={
            "company_name":cn,"street_address":sa,"phone":ph,"email":em,
            "cc_email":cem,"contact_name":ct,"project_number":pn,"project_name":prn,
            "invoice_to":it,"invoice_email":ie,"site_info":si,"purchase_order":po,
            "quote_number":qn,"county_state":cs,"time_zone":tz,
            "data_deliverable":dd,"regulatory_program":rp,"reportable":rep,
            "rush":rush,"pwsid":pw,"field_filtered":ff,
            "additional_instructions":ai2,"customer_remarks":rmk,
            "num_coolers":nc,"thermometer_id":tid,"temperature":tmp,
            "received_on_ice":roi,"tracking_number":trk,"delivery_method":dm,
            "project_manager":pm,"acct_num":an2,"table_number":tn,
            "profile_template":pt,"prelog_id":pl,"kelp_ordering_id":koid,
            "samples":sample_list,"analysis_columns":ac,
        }
        lp=LOGO_PATH if os.path.exists(LOGO_PATH) else None
        pdf=generate_coc_pdf(d,logo_path=lp)
        st.download_button("Download CoC PDF",pdf.read(),"KELP_CoC.pdf","application/pdf")
        st.success("CoC generated!")

if __name__=="__main__":
    main()

"""
coc_pdf_engine.py - KELP COC PDF Generator v20

- 0.15" margins (11pt) for maximum printable area
- Hybrid chemical symbols in analysis column headers
- 2-line vertical text per column (method + label)
"""
import io, os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

PW, PH = landscape(letter)  # 792 x 612
KELP_BLUE = HexColor("#1F4E79")
HDR_BLUE = HexColor("#1F4E79")
SECTION_BG = HexColor("#E8ECF0")
ROW_SHADE = HexColor("#F0F4F8")
LW_OUTER = 1.0; LW_SECTION = 0.5; LW_INNER = 0.3
DOC_ID = "KELP-QMS-FORM-001"; DOC_VERSION = "1.1"; DOC_EFF_DATE = "February 19, 2026"

# 0.15" margins = ~11pt
LM = 11; RM = 781
TM = 601; BM = 11

# Font sizes — slightly larger thanks to reclaimed space
FS_TITLE = 12; FS_LABEL = 7; FS_VALUE = 8.5; FS_HEADER = 7
FS_LEGEND = 6; FS_FOOTER = 5; FS_VERT = 5.5

COL2 = 230; ACOL = 470

from coc_catalog import KELP_ANALYTE_CATALOG, CAT_SHORT_MAP, SYMBOL_MAP, to_symbol, generate_coc_id, get_methods_for_category, POTABLE_MATRICES, NONPOTABLE_MATRICES

# === Drawing primitives ===

def R(c,x,y,w,h,fill=None,lw=LW_INNER):
    if fill: c.setFillColor(fill); c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(lw); c.rect(x,y,w,h,fill=0,stroke=1)

def T(c,x,y,txt,fs=FS_LABEL,bold=False,font="Helvetica",maxw=None,color=black):
    if not txt: return
    fn="Helvetica-Bold" if bold else font; s=float(fs); t=str(txt)
    if maxw:
        while s>4.0 and stringWidth(t,fn,s)>maxw: s-=0.3
    c.setFont(fn,s); c.setFillColor(color); c.drawString(x,y,t)

def TV(c,x,y,txt,fs=FS_VALUE,maxw=None):
    T(c,x,y,txt,fs=fs,bold=True,maxw=maxw)

def TC(c,x,y,w,txt,fs=FS_HEADER,bold=False,color=black):
    if not txt: return
    fn="Helvetica-Bold" if bold else "Helvetica"
    c.setFont(fn,fs); c.setFillColor(color); c.drawCentredString(x+w/2,y,str(txt))

def CB(c,x,y,checked=False,sz=7):
    c.setStrokeColor(black); c.setLineWidth(0.3); c.rect(x,y,sz,sz,fill=0,stroke=1)
    if checked: c.setFont("Helvetica-Bold",sz); c.setFillColor(black); c.drawCentredString(x+sz/2,y+0.5,"X")

def VTEXT(c,x,y,txt,fs=FS_VERT,bold=False):
    fn="Helvetica-Bold" if bold else "Helvetica"
    c.saveState(); c.setFont(fn,fs); c.setFillColor(black)
    c.translate(x,y); c.rotate(90); c.drawString(0,0,str(txt)); c.restoreState()

def HLINE(c,x1,x2,y,lw=0.4):
    c.setStrokeColor(black); c.setLineWidth(lw); c.line(x1,y,x2,y)

def SECTION_LABEL(c,x,y,w,h,text):
    c.setFillColor(SECTION_BG); c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(LW_SECTION); c.rect(x,y,w,h,fill=0,stroke=1)
    c.setFont("Helvetica-Bold",6.5); c.setFillColor(HDR_BLUE)
    c.drawCentredString(x+w/2,y+h/2-2.5,text)

def _footer(c,pn,tp,coc_id=""):
    y=BM+3
    c.setStrokeColor(black); c.setLineWidth(0.3); c.line(LM,y,RM,y)
    c.setFont("Helvetica",4.5); c.setFillColor(black)
    lt=DOC_ID+"  |  Version "+DOC_VERSION+"  |  Effective: "+DOC_EFF_DATE
    if coc_id: lt+="  |  COC ID: "+coc_id
    c.drawString(LM,y-8,lt); c.drawRightString(RM,y-8,"CONTROLLED DOCUMENT  |  Page "+str(pn)+" of "+str(tp))


def _build_analysis_columns(samples, avail_h):
    """Build columns using hybrid symbols. Each label must fit as single vertical line.
    Method strings are matrix-aware based on sample matrices present."""
    short_to_full = {v: k for k, v in CAT_SHORT_MAP.items()}
    cat_analytes = {}
    cat_order = list(KELP_ANALYTE_CATALOG.keys())
    
    # Collect all matrices present on the COC
    all_matrices = set()
    for s in samples:
        m = (s.get("matrix") or "").upper().strip()
        if m: all_matrices.add(m)
    
    for s in samples:
        analyses = s.get("analyses", {})
        if isinstance(analyses, list): analyses = {cat: [] for cat in analyses}
        for cn, al in analyses.items():
            if not al: continue
            resolved = cn
            if cn not in KELP_ANALYTE_CATALOG and cn in short_to_full:
                resolved = short_to_full[cn]
            if resolved not in cat_analytes: cat_analytes[resolved] = []
            for a in al:
                if a not in cat_analytes[resolved]: cat_analytes[resolved].append(a)

    columns = []
    fn_bold = "Helvetica-Bold"
    max_text_w = avail_h - 8

    for cn in cat_order:
        if cn not in cat_analytes: continue
        analytes = cat_analytes[cn]
        method = get_methods_for_category(cn, all_matrices)
        short = CAT_SHORT_MAP.get(cn, cn)
        msub = "(" + method + ")"

        # Convert to symbols
        sym_analytes = [to_symbol(a) for a in analytes]

        # Try all in one label
        full = short + " (" + ", ".join(sym_analytes) + ")"
        if stringWidth(full, fn_bold, FS_VERT) <= max_text_w:
            columns.append({"label": full, "method": msub, "cat_name": cn})
            continue

        # Split into chunks that fit
        chunks = []; cur = []
        for sa in sym_analytes:
            test_label = short + " (" + ", ".join(cur + [sa]) + ")"
            if cur and stringWidth(test_label, fn_bold, FS_VERT) > max_text_w:
                chunks.append(cur); cur = [sa]
            else:
                cur.append(sa)
        if cur: chunks.append(cur)

        for ci, chunk in enumerate(chunks):
            lbl = (short if ci == 0 else short + " cont'd") + " (" + ", ".join(chunk) + ")"
            columns.append({"label": lbl, "method": msub, "cat_name": cn})

    return columns


def generate_coc_pdf(data, logo_path=None):
    buf = io.BytesIO(); c = canvas.Canvas(buf, pagesize=(PW, PH))
    c.setTitle("KELP Chain-of-Custody")
    d = data; g = lambda k, dflt="": d.get(k, dflt) or dflt
    coc_id = d.get("coc_id") or generate_coc_id()
    total_pages = 2

    # Geometry calculations
    hdr_top = TM; hdr_bot = TM - 34; hdr_h = hdr_top - hdr_bot
    rh = 11; sec_h = 9; z3_rh = 20; lrh = 14
    lw_ = COL2 - LM; cw_ = ACOL - COL2

    ci_top = hdr_bot - sec_h
    z3_top = ci_top - 5*rh - sec_h
    z4_top = z3_top - sec_h - 4*z3_rh
    ya_b = z4_top - lrh
    dd_top = ya_b; dd_h = lrh*4
    ml_top = dd_top - dd_h; ml_h = 11
    th_top = ml_top - ml_h; th_h = 34
    tall_bot = th_top - th_h
    tall_h = z4_top - tall_bot

    dyn_cols = _build_analysis_columns(d.get("samples", []), tall_h)
    num_acols = max(len(dyn_cols), 1)

    KELP_STRIP_W=10; COMMENT_W=88; PNC_W=20
    right_fixed = KELP_STRIP_W + COMMENT_W + PNC_W
    atotal_w = RM - ACOL - right_fixed
    col_w = max(18, atotal_w / num_acols)
    if col_w * num_acols > atotal_w: col_w = atotal_w / num_acols
    AX = [ACOL + i * col_w for i in range(num_acols + 1)]
    KSX = AX[-1]; SBX = KSX + KELP_STRIP_W; PNCX = SBX + COMMENT_W

    # === PAGE 1 ===
    # Header
    c.setStrokeColor(KELP_BLUE); c.setLineWidth(1.5)
    c.line(LM, hdr_top, RM, hdr_top); c.line(LM, hdr_bot, RM, hdr_bot)
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER)
    c.rect(LM, hdr_bot, RM-LM, hdr_h, fill=0, stroke=1)

    if logo_path and os.path.exists(logo_path):
        try: c.drawImage(ImageReader(logo_path), LM+3, hdr_bot+3, width=80, height=hdr_h-6, preserveAspectRatio=True, mask='auto')
        except: pass

    c.setFont("Helvetica-Bold", FS_TITLE); c.setFillColor(KELP_BLUE)
    c.drawCentredString((LM+ACOL)/2, hdr_top-14, "CHAIN-OF-CUSTODY RECORD")
    c.setFont("Helvetica", 7.5); c.setFillColor(black)
    c.drawCentredString((LM+ACOL)/2, hdr_top-25, "Chain-of-Custody is a LEGAL DOCUMENT \u2014 Complete all relevant fields")

    kx = 570; R(c, kx, hdr_bot+1, RM-kx-1, hdr_h-2, lw=LW_SECTION)
    T(c, kx+3, hdr_top-11, "KELP USE ONLY", fs=7.5, bold=True, color=HDR_BLUE)
    T(c, kx+3, hdr_top-21, "KELP Ordering ID:", fs=FS_LABEL)
    HLINE(c, kx+75, RM-6, hdr_top-23, lw=0.3)
    kid = g("kelp_ordering_id")
    if kid: TV(c, kx+77, hdr_top-21, kid)
    T(c, kx+3, hdr_top-31, "COC ID: "+coc_id, fs=6.5, bold=True, color=HDR_BLUE)

    # CLIENT INFO
    SECTION_LABEL(c, LM, ci_top-sec_h, ACOL-LM, sec_h, "CLIENT INFORMATION")
    ci_y = ci_top - sec_h
    def ci_row(ri, ll, lv, cl, cv):
        yt=ci_y-ri*rh; yb=yt-rh; R(c,LM,yb,lw_,rh); R(c,COL2,yb,cw_,rh)
        if ll: T(c,LM+2,yb+2,ll); tw=stringWidth(ll,"Helvetica",FS_LABEL)+3; TV(c,LM+tw,yb+1,lv,maxw=lw_-tw-2)
        if cl: T(c,COL2+2,yb+2,cl); tw=stringWidth(cl,"Helvetica",FS_LABEL)+3; TV(c,COL2+tw,yb+1,cv,maxw=cw_-tw-2)

    ci_row(0, "Company Name:", g("company_name"), "Contact/Report To:", g("contact_name"))
    ci_row(1, "Street Address:", g("street_address"), "Phone #:", g("phone"))
    ci_row(2, "", "", "E-Mail:", g("email"))
    ci_row(3, "", "", "Cc E-Mail:", g("cc_email"))
    ci_row(4, "Customer Project #:", g("project_number"), "Invoice to:", g("invoice_to"))
    R(c, ACOL, ci_y-5*rh, RM-ACOL, 5*rh)

    # PROJECT DETAILS
    SECTION_LABEL(c, LM, z3_top, ACOL-LM, sec_h, "PROJECT DETAILS")
    y6b = z3_top - z3_rh
    R(c,LM,y6b,lw_,z3_rh); T(c,LM+2,y6b+11,"Project Name:"); TV(c,LM+58,y6b+11,g("project_name"),maxw=lw_-62)
    R(c,COL2,y6b,cw_,z3_rh); T(c,COL2+2,y6b+11,"Invoice E-mail:"); TV(c,COL2+62,y6b+11,g("invoice_email"),maxw=cw_-66)

    y7b = y6b - z3_rh
    R(c,LM,y7b,lw_,z3_rh); T(c,LM+2,y7b+11,"Site Collection Info/Facility ID (as applicable):"); TV(c,LM+2,y7b+1,g("site_info"),maxw=lw_-6)
    R(c,COL2,y7b,cw_,z3_rh); T(c,COL2+2,y7b+11,"Purchase Order (if applicable):"); TV(c,COL2+2,y7b+1,g("purchase_order"),maxw=cw_-6)

    y8b = y7b - z3_rh
    R(c,LM,y8b,lw_,z3_rh); R(c,COL2,y8b,cw_,z3_rh)
    T(c,COL2+2,y8b+11,"Quote #:"); TV(c,COL2+38,y8b+11,g("quote_number"),maxw=cw_-42)

    y9b = y8b - z3_rh
    R(c,LM,y9b,lw_,z3_rh); T(c,LM+2,y9b+11,"County / State origin of sample(s):"); TV(c,LM+2,y9b+1,g("county_state"),maxw=lw_-6)
    R(c,COL2,y9b,cw_,z3_rh)

    # Right legend area
    acol_w = AX[-1]-AX[0]; lr_x = SBX; lr_w = RM - SBX
    R(c,ACOL,y6b,acol_w+KELP_STRIP_W,z3_rh); R(c,lr_x,y6b,lr_w,z3_rh)
    R(c,ACOL,y7b,acol_w+KELP_STRIP_W,z3_rh)
    cs = g("container_size")
    if cs: T(c,ACOL+2,y7b+11,"Specify Container Size:"); TV(c,ACOL+100,y7b+11,cs,maxw=acol_w-105)
    else: TC(c,ACOL,y7b+7,acol_w,"Specify Container Size",fs=FS_HEADER,bold=True)
    R(c,lr_x,y7b,lr_w,z3_rh)
    T(c,lr_x+2,y7b+13,"Container Size: (1) 1L, (2) 500mL,",fs=5.5,bold=True)
    T(c,lr_x+2,y7b+7,"(3) 250mL, (4) 125mL, (5) 100mL,",fs=5.5,bold=True)
    T(c,lr_x+2,y7b+1,"(6) Other",fs=5.5,bold=True)

    R(c,ACOL,y8b,acol_w+KELP_STRIP_W,z3_rh)
    pv = g("preservative_type")
    if pv: T(c,ACOL+2,y8b+11,"Identify Container Preservative Type:"); TV(c,ACOL+150,y8b+11,pv,maxw=acol_w-155)
    else: TC(c,ACOL,y8b+7,acol_w,"Identify Container Preservative Type",fs=FS_HEADER,bold=True)
    R(c,lr_x,y9b,lr_w,y7b-y9b)
    T(c,lr_x+2,y8b+11,"Preservative: (1) None, (2) HNO3,",fs=FS_LEGEND)
    T(c,lr_x+2,y8b+5,"(3) H2SO4, (4) HCl, (5) NaOH,",fs=FS_LEGEND)
    T(c,lr_x+2,y9b+15,"(6) Zn Acetate, (7) NaHSO4,",fs=FS_LEGEND)
    T(c,lr_x+2,y9b+9,"(8) Sod.Thiosulfate, (9) Ascorbic Acid,",fs=FS_LEGEND)
    T(c,lr_x+2,y9b+3,"(10) MeOH, (11) Other",fs=FS_LEGEND)
    R(c,ACOL,y9b,acol_w+KELP_STRIP_W,z3_rh)
    TC(c,ACOL,y9b+7,acol_w,"Analysis Requested",fs=FS_HEADER,bold=True)

    # ZONE 4: Timezone + Deliverables
    R(c,LM,ya_b,ACOL-LM,lrh); T(c,LM+2,ya_b+4,"Sample Collection Time Zone:")
    tz = g("time_zone","PT")
    for tn,tx in [("AK",135),("PT",165),("MT",193),("CT",219),("ET",245)]:
        CB(c,tx,ya_b+3,checked=(tn==tz)); T(c,tx+9,ya_b+4,tn)

    dd_w=135; reg_w=ACOL-LM-dd_w; rx=LM+dd_w
    R(c,LM,dd_top-dd_h,dd_w,dd_h); T(c,LM+2,dd_top-11,"Data Deliverables:")
    sd = g("data_deliverable","Level I (Std)")
    for lbl,bx,by in [("Level I (Std)",LM+6,dd_top-25),("Level II",LM+72,dd_top-25),("Level III",LM+6,dd_top-40),("Other",LM+6,dd_top-55)]:
        CB(c,bx,by,checked=(lbl==sd)); T(c,bx+9,by+1,lbl)

    R(c,rx,dd_top-lrh,reg_w,lrh); T(c,rx+2,dd_top-lrh+4,"Regulatory Program (DW, RCRA, etc.):")
    rpx=rx+reg_w*0.62; T(c,rpx,dd_top-lrh+4,"Reportable"); rv=g("reportable")=="Yes"
    CB(c,rpx+44,dd_top-lrh+3,checked=rv); T(c,rpx+53,dd_top-lrh+4,"Yes")
    CB(c,rpx+72,dd_top-lrh+3,checked=not rv); T(c,rpx+81,dd_top-lrh+4,"No")

    R(c,rx,dd_top-lrh*2,reg_w,lrh); T(c,rx+2,dd_top-lrh*2+4,"Rush (Pre-approval required):",bold=True)
    sr = g("rush","Standard (5-10 Day)")
    for i,ro in enumerate(["Same Day","1 Day","2 Day","3 Day","4 Day"]):
        rrx=rx+130+i*42; CB(c,rrx,dd_top-lrh*2+3,checked=(ro==sr)); T(c,rrx+9,dd_top-lrh*2+4,ro,fs=FS_LEGEND)

    h3=reg_w*0.5; R(c,rx,dd_top-lrh*3,h3,lrh)
    CB(c,rx+4,dd_top-lrh*3+3,checked=("5 Day" in sr)); T(c,rx+14,dd_top-lrh*3+4,"5 Day"); T(c,rx+50,dd_top-lrh*3+4,"Other ____________")
    R(c,rx+h3,dd_top-lrh*3,reg_w-h3,lrh); T(c,rx+h3+2,dd_top-lrh*3+4,"DW PWSID # or WW Permit #:",fs=FS_LEGEND)

    R(c,rx,dd_top-lrh*4,reg_w,lrh); T(c,rx+2,dd_top-lrh*4+4,"Field Filtered (if applicable):")
    ff = g("field_filtered","No")
    CB(c,rx+130,dd_top-lrh*4+3,checked=(ff=="Yes")); T(c,rx+139,dd_top-lrh*4+4,"Yes")
    CB(c,rx+163,dd_top-lrh*4+3,checked=(ff=="No")); T(c,rx+172,dd_top-lrh*4+4,"No")

    R(c,LM,ml_top-ml_h,ACOL-LM,ml_h)
    T(c,LM+2,ml_top-ml_h+2,"* Matrix: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water(SW), Other(OT)",fs=FS_LEGEND)

    # TABLE HEADERS
    grp_h=13; sub_h=th_h-grp_h
    R(c,LM,th_top-th_h,163,th_h); TC(c,LM,th_top-th_h/2-3,163,"Customer Sample ID",fs=FS_HEADER,bold=True)
    R(c,174,th_top-th_h,32,th_h); TC(c,174,th_top-th_h/2+3,32,"Matrix",fs=FS_HEADER,bold=True); TC(c,174,th_top-th_h/2-7,32,"*",fs=FS_HEADER,bold=True)
    R(c,206,th_top-th_h,28,th_h); TC(c,206,th_top-th_h/2+3,28,"Comp /",fs=FS_HEADER,bold=True); TC(c,206,th_top-th_h/2-7,28,"Grab",fs=FS_HEADER,bold=True)
    R(c,234,th_top-grp_h,90,grp_h); TC(c,234,th_top-grp_h+2,90,"Composite Start",fs=FS_LABEL,bold=True)
    R(c,234,th_top-th_h,55,sub_h); TC(c,234,th_top-th_h+sub_h/2-3,55,"Date",fs=FS_LABEL,bold=True)
    R(c,289,th_top-th_h,35,sub_h); TC(c,289,th_top-th_h+sub_h/2-3,35,"Time",fs=FS_LABEL,bold=True)
    R(c,324,th_top-grp_h,86,grp_h); TC(c,324,th_top-grp_h+2,86,"Collected or Composite End",fs=FS_LEGEND,bold=True)
    R(c,324,th_top-th_h,51,sub_h); TC(c,324,th_top-th_h+sub_h/2-3,51,"Date",fs=FS_LABEL,bold=True)
    R(c,375,th_top-th_h,35,sub_h); TC(c,375,th_top-th_h+sub_h/2-3,35,"Time",fs=FS_LABEL,bold=True)
    R(c,410,th_top-th_h,20,th_h); TC(c,410,th_top-th_h/2+3,20,"#",fs=FS_HEADER,bold=True); TC(c,410,th_top-th_h/2-7,20,"Cont.",fs=FS_HEADER,bold=True)
    R(c,430,th_top-grp_h,40,grp_h); TC(c,430,th_top-grp_h+2,40,"Residual Chlorine",fs=5.5,bold=True)
    R(c,430,th_top-th_h,22,sub_h); TC(c,430,th_top-th_h+sub_h/2-3,22,"Result",fs=FS_LEGEND,bold=True)
    R(c,452,th_top-th_h,18,sub_h); TC(c,452,th_top-th_h+sub_h/2-3,18,"Units",fs=FS_LEGEND,bold=True)

    # VERTICAL ANALYSIS COLUMNS
    cat_col_indices = {}
    for ci, col_info in enumerate(dyn_cols):
        cat = col_info["cat_name"]
        if cat not in cat_col_indices: cat_col_indices[cat] = []
        cat_col_indices[cat].append(ci)

    for ci in range(num_acols):
        ax0 = AX[ci]; ax1 = AX[ci+1] if ci+1 < len(AX) else KSX; aw = ax1 - ax0
        R(c, ax0, tall_bot, aw, tall_h)
        if ci < len(dyn_cols):
            col_info = dyn_cols[ci]; label = col_info["label"]; method = col_info["method"]
            avail = tall_h - 8
            lfs = FS_VERT
            while lfs > 4.0 and stringWidth(label, "Helvetica-Bold", lfs) > avail:
                lfs -= 0.3
            mfs = min(lfs, FS_VERT - 0.5)
            label_x = ax0 + aw * 0.65
            method_x = ax0 + aw * 0.25
            VTEXT(c, label_x, tall_bot + 4, label, fs=lfs, bold=True)
            VTEXT(c, method_x, tall_bot + 4, method, fs=mfs, bold=False)

    # KELP strip
    R(c, KSX, tall_bot, KELP_STRIP_W, tall_h)
    VTEXT(c, KSX+KELP_STRIP_W/2+2, tall_bot+tall_h/2-20, "KELP Use Only", fs=FS_LEGEND, bold=True)

    # Right side fields
    sbf_h = 19; fy = z4_top
    for lbl,val in [("Project Mgr.:",g("project_manager")),("AcctNum / Client ID:",g("acct_num")),("Table #:",g("table_number")),("Profile / Template:",g("profile_template")),("Prelog / Bottle Ord. ID:",g("prelog_id"))]:
        R(c,SBX,fy-sbf_h,COMMENT_W,sbf_h)
        T(c,SBX+2,fy-sbf_h+9,lbl); lw2=stringWidth(lbl,"Helvetica",FS_LABEL)+3
        TV(c,SBX+lw2,fy-sbf_h+9,val,fs=7,maxw=COMMENT_W-lw2-2); fy-=sbf_h
    sc_h = fy - tall_bot; R(c, SBX, tall_bot, COMMENT_W, sc_h)
    TC(c, SBX, tall_bot+sc_h/2-3, COMMENT_W, "Sample Comment", fs=FS_HEADER, bold=True)

    R(c, PNCX, tall_bot, PNC_W, z4_top-tall_bot)
    VTEXT(c, PNCX+PNC_W/2+4, tall_bot+6, "Preservation non-conformance", fs=5, bold=False)
    VTEXT(c, PNCX+PNC_W/2-4, tall_bot+6, "identified for sample.", fs=5, bold=False)

    # SAMPLE DATA ROWS
    data_rh = 18; max_rows = 10; data_top = tall_bot
    SECTION_LABEL(c, LM, data_top-sec_h, ACOL-LM, sec_h, "SAMPLE INFORMATION"); data_top -= sec_h
    samples = d.get("samples", [])
    for ri in range(max_rows):
        s = samples[ri] if ri < len(samples) else {}
        ryb = data_top - (ri+1)*data_rh
        if ri%2==1: c.setFillColor(ROW_SHADE); c.rect(LM,ryb,RM-LM,data_rh,fill=1,stroke=0)
        c.setFont("Helvetica",5.5); c.setFillColor(black); c.drawCentredString(LM+5,ryb+8,str(ri+1))
        for x0,x1,val,align,fs in [(LM,174,s.get("sample_id",""),"left",FS_VALUE),(174,206,s.get("matrix",""),"center",FS_VALUE),(206,234,s.get("comp_grab",""),"center",7),(234,289,s.get("start_date",""),"center",7),(289,324,s.get("start_time",""),"center",7),(324,375,s.get("end_date","") or s.get("collected_date",""),"center",7),(375,410,s.get("end_time","") or s.get("collected_time",""),"center",7),(410,430,s.get("num_containers",""),"center",FS_VALUE),(430,452,s.get("res_cl_result",""),"center",7),(452,ACOL,s.get("res_cl_units",""),"center",FS_LEGEND)]:
            R(c,x0,ryb,x1-x0,data_rh)
            if val:
                if align=="center": TC(c,x0,ryb+6,x1-x0,str(val),fs=fs,bold=True)
                else: TV(c,x0+11,ryb+6,str(val),fs=fs,maxw=x1-x0-15)
        sa = s.get("analyses",{})
        if isinstance(sa,list): sa = {cat:[] for cat in sa}
        for ci2 in range(num_acols):
            ax0=AX[ci2]; ax1=AX[ci2+1] if ci2+1<len(AX) else KSX; R(c,ax0,ryb,ax1-ax0,data_rh)
        short_to_full = {v: k for k, v in CAT_SHORT_MAP.items()}
        for cat_name,al in sa.items():
            if not al: continue
            resolved = cat_name
            if cat_name not in cat_col_indices and cat_name in short_to_full:
                resolved = short_to_full[cat_name]
            if resolved not in cat_col_indices: continue
            for ci_idx in cat_col_indices[resolved]:
                if ci_idx < num_acols:
                    ax0=AX[ci_idx]; ax1=AX[ci_idx+1] if ci_idx+1<len(AX) else KSX
                    TC(c,ax0,ryb+6,ax1-ax0,"X",fs=FS_VALUE,bold=True)
        R(c,KSX,ryb,KELP_STRIP_W,data_rh); R(c,SBX,ryb,COMMENT_W,data_rh)
        cmt = s.get("comment","")
        if cmt: TV(c,SBX+2,ryb+6,cmt,fs=FS_LEGEND,maxw=COMMENT_W-4)
        R(c,PNCX,ryb,PNC_W,data_rh)

    # BOTTOM ZONE
    bot_top = data_top - max_rows*data_rh
    SECTION_LABEL(c, LM, bot_top-sec_h, RM-LM, sec_h, "CHAIN OF CUSTODY RECORD / LABORATORY RECEIVING"); bot_top -= sec_h
    inst_h = 14; half_w = (RM-LM)/2
    R(c,LM,bot_top-inst_h,half_w,inst_h)
    T(c,LM+2,bot_top-7,"Additional Instructions for KELP:",bold=True)
    TV(c,LM+4,bot_top-13,g("additional_instructions"),fs=6,maxw=half_w-8)
    R(c,LM+half_w,bot_top-inst_h,half_w,inst_h)
    T(c,LM+half_w+2,bot_top-7,"Customer Remarks / Special Conditions / Possible Hazards:",bold=True)
    TV(c,LM+half_w+4,bot_top-13,g("customer_remarks"),fs=6,maxw=half_w-8)

    lr_top = bot_top - inst_h; lr_h = 10
    R(c,LM,lr_top-lr_h,RM-LM,lr_h)
    T(c,LM+2,lr_top-9,"# Coolers:"); TV(c,LM+48,lr_top-9,g("num_coolers"),fs=7.5)
    T(c,LM+78,lr_top-9,"Thermometer ID:"); TV(c,LM+140,lr_top-9,g("thermometer_id"),fs=7.5)
    T(c,LM+200,lr_top-9,"Temp. (\u00b0C):"); TV(c,LM+240,lr_top-9,g("temperature"),fs=7.5)
    roi = g("received_on_ice","Yes")
    T(c,LM+320,lr_top-9,"Sample Received on ice:")
    CB(c,LM+414,lr_top-11,checked=(roi=="Yes")); T(c,LM+423,lr_top-9,"Yes")
    CB(c,LM+448,lr_top-11,checked=(roi=="No")); T(c,LM+457,lr_top-9,"No")

    rel_top = lr_top - lr_h; rel_h = 9
    for row_i in range(3):
        ryb=rel_top-(row_i+1)*rel_h; rw1=215; dtw=92
        R(c,LM,ryb,rw1,rel_h); T(c,LM+2,ryb+3,"Relinquished by/Company: (Signature)",fs=FS_LEGEND)
        R(c,LM+rw1,ryb,dtw,rel_h); T(c,LM+rw1+2,ryb+3,"Date/Time:",fs=FS_LEGEND)
        rcx=LM+rw1+dtw; R(c,rcx,ryb,rw1,rel_h); T(c,rcx+2,ryb+3,"Received by/Company: (Signature)",fs=FS_LEGEND)
        R(c,rcx+rw1,ryb,dtw,rel_h); T(c,rcx+rw1+2,ryb+3,"Date/Time:",fs=FS_LEGEND)
        rsx=rcx+rw1+dtw; rsw=RM-rsx; R(c,rsx,ryb,rsw,rel_h)
        if row_i==0: T(c,rsx+2,ryb+3,"Tracking #:",fs=FS_LEGEND); TV(c,rsx+46,ryb+3,g("tracking_number"),fs=7,maxw=rsw-50)
        elif row_i==1:
            dm=g("delivery_method"); T(c,rsx+2,ryb+3,"Delivered by:",fs=FS_LEGEND); avail=rsw-58
            for dn,dx in [("In-Person",rsx+55),("FedEx",rsx+55+avail*0.30),("UPS",rsx+55+avail*0.55),("Other",rsx+55+avail*0.75)]:
                CB(c,dx,ryb+1,checked=(dn==dm),sz=6); T(c,dx+8,ryb+3,dn,fs=FS_LEGEND)

    form_bot = rel_top - 3*rel_h
    if form_bot < BM+18: form_bot = BM+18
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER)
    c.rect(LM, form_bot, RM-LM, hdr_top-form_bot, fill=0, stroke=1)
    disc_y = form_bot - 5
    T(c, LM+5, disc_y, "Submitting a sample via this chain of custody constitutes acknowledgment and acceptance of the KELP\u2019s Terms and Conditions", fs=4.5)
    _footer(c, 1, total_pages, coc_id)

    # === PAGE 2: INSTRUCTIONS ===
    c.showPage()
    c.setFont("Helvetica-Bold",14); c.setFillColor(KELP_BLUE)
    c.drawCentredString(PW/2,PH-40,"Chain of Custody (COC) Instructions")
    c.setFont("Helvetica",10); c.setFillColor(black)
    c.drawCentredString(PW/2,PH-56,"Complete all relevant fields on the COC form. Incomplete information may cause delays.")
    col1_x=30; col2_x=420; cw=370; bfs=9.5; hfs=10; bul_=9
    def sec_heading(yp,text,cx=col1_x):
        c.setFont("Helvetica-Bold",hfs); c.setFillColor(HDR_BLUE); c.drawString(cx,yp,text); return yp-14
    def bullet_item(yp,label,desc,cx=col1_x,w=cw):
        c.setFont("Helvetica-Bold",bfs); c.setFillColor(black); c.drawString(cx+bul_,yp,"\u2022")
        lw2=stringWidth(label,"Helvetica-Bold",bfs); c.drawString(cx+bul_+8,yp,label)
        c.setFont("Helvetica",bfs); remaining=desc; xs=cx+bul_+8+lw2+3; aw=w-bul_-8-(lw2+3)
        while remaining:
            words=remaining.split(); line=""; rw=[]
            for i,word in enumerate(words):
                test=line+(" " if line else "")+word
                if stringWidth(test,"Helvetica",bfs)<=aw: line=test
                else: rw=words[i:]; break
            else: rw=[]
            c.drawString(xs,yp,line); remaining=" ".join(rw)
            if remaining: yp-=12; xs=cx+bul_+8; aw=w-bul_-8
        return yp-14
    y=PH-85; y=sec_heading(y,"1. Client & Project Information:")
    for l,d2 in [("Company Name:","Your company\u2019s name."),("Street Address:","Your mailing address."),("Contact/Report To:","Person designated to receive results."),("Customer Project # and Project Name:","Your project reference number and name."),("Site Collection Info/Facility ID:","Project location or facility ID."),("Time Zone:","Sample collection time zone (e.g., AK, PT, MT, CT, ET)."),("Purchase Order #:","Your PO number for invoicing."),("Invoice To:","Contact person for the invoice."),("Invoice Email:","Email address for the invoice."),("Phone #:","Your contact phone number."),("E-mail:","Your email for correspondence and the final report."),("Data Deliverable:","Required data deliverable level."),("Field Filtered:","Indicate if samples were filtered in the field (Yes/No)."),("Quote #:","Quote number, if applicable."),("DW PWSID # or WW Permit #:","Relevant permit numbers, if applicable.")]:
        y=bullet_item(y,l,d2)
    y-=4; y=sec_heading(y,"2. Sample Information:")
    for l,d2 in [("Customer Sample ID:","Unique sample identifier for the report."),("Collected Date:","Sample collection date."),("Collected Time:","Sample collection time."),("Comp/Grab:","GRAB for single-point; COMP for combined samples."),("Matrix:","Sample type (DW, GW, WW, P, SW, OT)."),("Container Size:","Specify size (e.g., 1L, 500mL).")]:
        y=bullet_item(y,l,d2)
    y2=PH-85
    for l,d2 in [("Container Preservative Type:","Specify preservative (None, HNO3, H2SO4, etc.)."),("Analysis Requested:","Mark categories for each sample. Analyte names shown in column headers."),("Sample Comment:","Notes about individual samples; identify MS/MSD samples here."),("Residual Chlorine:","Record results and units if measured.")]:
        y2=bullet_item(y2,l,d2,cx=col2_x,w=cw)
    y2-=4; y2=sec_heading(y2,"3. Additional Information & Instructions:",cx=col2_x)
    for l,d2 in [("Customer Remarks/Hazards:","Note special instructions or hazards."),("Rush Request:","For expedited results; pre-approval required."),("Relinquished By/Received By:","Sign and date at each transfer of custody.")]:
        y2=bullet_item(y2,l,d2,cx=col2_x,w=cw)
    y2-=4; y2=sec_heading(y2,"4. Sample Acceptance Policy Summary:",cx=col2_x)
    c.setFont("Helvetica",bfs); c.setFillColor(black); c.drawString(col2_x,y2,"For samples to be accepted, ensure:"); y2-=14
    for item in ["Complete COC documentation.","Readable, unique sample ID on containers (indelible ink).","Appropriate containers and sufficient volume.","Receipt within holding time and temperature requirements.","Containers are in good condition, seals intact (if used).","Proper preservation, no headspace in volatile water samples.","Adequate volume for MS/MSD if required."]:
        c.setFont("Helvetica",bfs); c.drawString(col2_x+bul_,y2,"\u2022  "+item); y2-=13
    y2-=6; c.setFont("Helvetica",8.5); c.setFillColor(black)
    closing="Failure to meet these may result in data qualifiers. A detailed policy is available from your Project Manager. Submitting samples implies acceptance of KELP Terms and Conditions."
    words=closing.split(); line=""; cy=y2
    for word in words:
        test=line+(" " if line else "")+word
        if stringWidth(test,"Helvetica",8.5)<=cw: line=test
        else: c.drawString(col2_x,cy,line); cy-=12; line=word
    if line: c.drawString(col2_x,cy,line)
    _footer(c, 2, total_pages, coc_id)
    c.showPage(); c.save(); buf.seek(0)
    return buf, coc_id

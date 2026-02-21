"""
coc_pdf_engine.py — KELP COC PDF Generator v17

Fixes from v16:
1. Restore Comp/Grab date columns always (Start + End headers always shown)
   — Comp/Grab logic is Streamlit-side only
2. Ensure all content fits printable area (landscape letter with safe margins)
3. Wrap vertical column text if too long (multi-line vertical)
4. Remove "Analysis:" row from deliverables section
5. Remove "Level IV" from data deliverables
"""
import io, os, textwrap
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

PW, PH = landscape(letter)

KELP_BLUE   = HexColor("#1F4E79")
HDR_BLUE    = HexColor("#1F4E79")
SECTION_BG  = HexColor("#E8ECF0")
ROW_SHADE   = HexColor("#F0F4F8")

LW_OUTER = 1.0; LW_SECTION = 0.5; LW_INNER = 0.3

DOC_ID = "KELP-QMS-FORM-001"
DOC_TITLE = "Chain-of-Custody"
DOC_VERSION = "1.1"
DOC_EFF_DATE = "February 19, 2026"

FS_TITLE=11; FS_LABEL=6; FS_VALUE=8; FS_HEADER=6.5
FS_LEGEND=5; FS_FOOTER=5; FS_SUBTITLE=7

# Page geometry — safe printable area
LM = 18; RM = 774   # ~0.25" margins
COL2 = 237; ACOL = 476

from coc_catalog import KELP_ANALYTE_CATALOG, CAT_SHORT_MAP, generate_coc_id


# ═══════════ Drawing primitives ═══════════
def R(c,x,y,w,h,fill=None,lw=LW_INNER):
    if fill: c.setFillColor(fill); c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(lw); c.rect(x,y,w,h,fill=0,stroke=1)

def T(c,x,y,txt,fs=FS_LABEL,bold=False,font="Helvetica",maxw=None,color=black):
    if not txt: return
    fn="Helvetica-Bold" if bold else font; s=float(fs); t=str(txt)
    if maxw:
        while s>3.5 and stringWidth(t,fn,s)>maxw: s-=0.3
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

def VTEXT(c,x,y,txt,fs=FS_HEADER,bold=False):
    fn="Helvetica-Bold" if bold else "Helvetica"
    c.saveState(); c.setFont(fn,fs); c.setFillColor(black)
    c.translate(x,y); c.rotate(90); c.drawString(0,0,str(txt)); c.restoreState()

def VTEXT_WRAP(c, x, y, txt, fs, bold, col_w, max_h):
    """Render vertical text with wrapping. Multiple lines spaced across column width."""
    fn = "Helvetica-Bold" if bold else "Helvetica"
    # Calculate how many chars fit in max_h height (rendered vertically)
    char_w = stringWidth("W", fn, fs)
    chars_fit = max(10, int(max_h / char_w * 1.6))  # approximate
    
    if stringWidth(txt, fn, fs) <= max_h:
        # Fits in one line
        VTEXT(c, x, y, txt, fs=fs, bold=bold)
        return
    
    # Wrap text
    lines = textwrap.wrap(txt, width=chars_fit)
    num_lines = len(lines)
    line_spacing = min(8, (col_w - 2) / max(num_lines, 1))
    start_x = x + (col_w / 2) + (num_lines - 1) * line_spacing / 2
    
    for i, line in enumerate(lines):
        lx = start_x - i * line_spacing
        VTEXT(c, lx, y, line, fs=fs, bold=bold)

def HLINE(c,x1,x2,y,lw=0.4):
    c.setStrokeColor(black); c.setLineWidth(lw); c.line(x1,y,x2,y)

def SECTION_LABEL(c,x,y,w,h,text):
    c.setFillColor(SECTION_BG); c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(LW_SECTION); c.rect(x,y,w,h,fill=0,stroke=1)
    c.setFont("Helvetica-Bold",5.5); c.setFillColor(HDR_BLUE)
    c.drawCentredString(x+w/2,y+h/2-2,text)

def _footer(c,pn,tp,coc_id=""):
    c.setFont("Helvetica",FS_FOOTER); c.setFillColor(black)
    c.setStrokeColor(black); c.setLineWidth(0.3); c.line(LM,14,RM,14)
    lt=f"{DOC_ID}  |  Version {DOC_VERSION}  |  Effective: {DOC_EFF_DATE}"
    if coc_id: lt+=f"  |  COC ID: {coc_id}"
    c.drawString(LM,6,lt)
    c.drawRightString(RM,6,f"CONTROLLED DOCUMENT  |  Page {pn} of {tp}")


# ═══════════ Build dynamic analysis columns ═══════════
def _build_analysis_columns(samples):
    cat_analytes = {}
    cat_order = list(KELP_ANALYTE_CATALOG.keys())
    for s in samples:
        analyses = s.get("analyses", {})
        if isinstance(analyses, list): analyses = {cat: [] for cat in analyses}
        for cat_name, al in analyses.items():
            if not al: continue
            if cat_name not in cat_analytes: cat_analytes[cat_name] = []
            for a in al:
                if a not in cat_analytes[cat_name]: cat_analytes[cat_name].append(a)

    columns = []
    MAX_CHARS = 100
    for cat_name in cat_order:
        if cat_name not in cat_analytes: continue
        analytes = cat_analytes[cat_name]
        method = ", ".join(KELP_ANALYTE_CATALOG[cat_name]["methods"])
        short = CAT_SHORT_MAP.get(cat_name, cat_name)
        a_str = ", ".join(analytes)
        full = f"{short} ({a_str})"
        msub = f"({method})"

        if len(full) <= MAX_CHARS:
            columns.append({"label": full, "method": msub, "cat_name": cat_name})
        else:
            chunks = []; cur = []; cur_len = len(short) + 3
            for a in analytes:
                add = len(a) + (2 if cur else 0)
                if cur_len + add > MAX_CHARS and cur:
                    chunks.append(cur); cur = [a]; cur_len = len(short) + 10 + len(a)
                else:
                    cur.append(a); cur_len += add
            if cur: chunks.append(cur)
            for ci, chunk in enumerate(chunks):
                lbl = f"{short} ({', '.join(chunk)})" if ci == 0 else f"{short} cont'd ({', '.join(chunk)})"
                columns.append({"label": lbl, "method": msub, "cat_name": cat_name})
    return columns


def generate_coc_pdf(data, logo_path=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PW, PH))
    c.setTitle("KELP Chain-of-Custody")
    d = data
    g = lambda k, dflt="": d.get(k, dflt) or dflt
    coc_id = d.get("coc_id") or generate_coc_id()
    dyn_cols = _build_analysis_columns(d.get("samples", []))
    num_acols = max(len(dyn_cols), 1)
    total_pages = 2

    # Layout: right side fixed widths
    KELP_STRIP_W = 9.5; COMMENT_W = 85; PNC_W = 20
    right_fixed = KELP_STRIP_W + COMMENT_W + PNC_W
    atotal_w = RM - ACOL - right_fixed
    col_w = max(18, atotal_w / num_acols)
    # Clamp columns to fit
    if col_w * num_acols > atotal_w:
        col_w = atotal_w / num_acols

    AX = [ACOL + i * col_w for i in range(num_acols + 1)]
    KSX = AX[-1]  # KELP strip x
    SBX = KSX + KELP_STRIP_W
    PNCX = SBX + COMMENT_W

    # ═══════════════════════════════════════
    # PAGE 1
    # ═══════════════════════════════════════
    hdr_top=590; hdr_bot=558; hdr_h=hdr_top-hdr_bot
    c.setStrokeColor(KELP_BLUE); c.setLineWidth(1.5)
    c.line(LM,hdr_top,RM,hdr_top); c.line(LM,hdr_bot,RM,hdr_bot)
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER)
    c.rect(LM,hdr_bot,RM-LM,hdr_h,fill=0,stroke=1)

    if logo_path and os.path.exists(logo_path):
        try: c.drawImage(ImageReader(logo_path),LM+3,hdr_bot+3,width=80,height=hdr_h-6,preserveAspectRatio=True,mask='auto')
        except: pass

    c.setFont("Helvetica-Bold",FS_TITLE); c.setFillColor(KELP_BLUE)
    c.drawCentredString((LM+ACOL)/2,hdr_top-13,"CHAIN-OF-CUSTODY RECORD")
    c.setFont("Helvetica",7); c.setFillColor(black)
    c.drawCentredString((LM+ACOL)/2,hdr_top-24,"Chain-of-Custody is a LEGAL DOCUMENT \u2014 Complete all relevant fields")

    # KELP USE ONLY
    kx=570
    R(c,kx,hdr_bot+1,RM-kx-1,hdr_h-2,lw=LW_SECTION)
    T(c,kx+3,hdr_top-10,"KELP USE ONLY",fs=6.5,bold=True,color=HDR_BLUE)
    T(c,kx+3,hdr_top-20,"KELP Ordering ID:",fs=FS_LABEL)
    HLINE(c,kx+68,RM-6,hdr_top-22,lw=0.3)
    kid=g("kelp_ordering_id")
    if kid: TV(c,kx+70,hdr_top-20,kid)
    T(c,kx+3,hdr_top-29,f"COC ID: {coc_id}",fs=5.5,bold=True,color=HDR_BLUE)

    # CLIENT INFO
    ci_top=hdr_bot; rh=10; lw_=COL2-LM; cw_=ACOL-COL2; sec_h=8
    SECTION_LABEL(c,LM,ci_top-sec_h,ACOL-LM,sec_h,"CLIENT INFORMATION"); ci_top-=sec_h

    def ci_row(ri,ll,lv,cl,cv):
        yt=ci_top-ri*rh; yb=yt-rh
        R(c,LM,yb,lw_,rh); R(c,COL2,yb,cw_,rh)
        if ll: T(c,LM+2,yb+2,ll); tw=stringWidth(ll,"Helvetica",FS_LABEL)+3; TV(c,LM+tw,yb+1,lv,maxw=lw_-tw-2)
        if cl: T(c,COL2+2,yb+2,cl); tw=stringWidth(cl,"Helvetica",FS_LABEL)+3; TV(c,COL2+tw,yb+1,cv,maxw=cw_-tw-2)

    ci_row(0,"Company Name:",g("company_name"),"Contact/Report To:",g("contact_name"))
    ci_row(1,"Street Address:",g("street_address"),"Phone #:",g("phone"))
    ci_row(2,"","","E-Mail:",g("email"))
    ci_row(3,"","","Cc E-Mail:",g("cc_email"))
    ci_row(4,"Customer Project #:",g("project_number"),"Invoice to:",g("invoice_to"))
    R(c,ACOL,ci_top-5*rh,RM-ACOL,5*rh)

    # PROJECT DETAILS
    z3_top=ci_top-5*rh; z3_rh=19
    SECTION_LABEL(c,LM,z3_top-sec_h,ACOL-LM,sec_h,"PROJECT DETAILS"); z3_top-=sec_h

    y6b=z3_top-z3_rh
    R(c,LM,y6b,lw_,z3_rh); T(c,LM+2,y6b+10,"Project Name:",fs=FS_LABEL); TV(c,LM+55,y6b+10,g("project_name"),maxw=lw_-59)
    R(c,COL2,y6b,cw_,z3_rh); T(c,COL2+2,y6b+10,"Invoice E-mail:",fs=FS_LABEL); TV(c,COL2+60,y6b+10,g("invoice_email"),maxw=cw_-64)

    y7b=y6b-z3_rh
    R(c,LM,y7b,lw_,z3_rh); T(c,LM+2,y7b+10,"Site Collection Info/Facility ID (as applicable):",fs=FS_LABEL); TV(c,LM+2,y7b+1,g("site_info"),maxw=lw_-6)
    R(c,COL2,y7b,cw_,z3_rh); T(c,COL2+2,y7b+10,"Purchase Order (if applicable):",fs=FS_LABEL); TV(c,COL2+2,y7b+1,g("purchase_order"),maxw=cw_-6)

    y8b=y7b-z3_rh
    R(c,LM,y8b,lw_,z3_rh); R(c,COL2,y8b,cw_,z3_rh)
    T(c,COL2+2,y8b+10,"Quote #:",fs=FS_LABEL); TV(c,COL2+35,y8b+10,g("quote_number"),maxw=cw_-39)

    y9b=y8b-z3_rh
    R(c,LM,y9b,lw_,z3_rh); T(c,LM+2,y9b+10,"County / State origin of sample(s):",fs=FS_LABEL); TV(c,LM+2,y9b+1,g("county_state"),maxw=lw_-6)
    R(c,COL2,y9b,cw_,z3_rh)

    # Right side legend
    acol_w = AX[-1] - AX[0]
    legend_right_x = SBX
    legend_right_w = RM - SBX

    R(c,ACOL,y6b,acol_w+KELP_STRIP_W,z3_rh)
    R(c,legend_right_x,y6b,legend_right_w,z3_rh)

    # Container size
    R(c,ACOL,y7b,acol_w+KELP_STRIP_W,z3_rh)
    cs_val=g("container_size")
    if cs_val: T(c,ACOL+2,y7b+10,"Specify Container Size:",fs=FS_LABEL); TV(c,ACOL+95,y7b+10,cs_val,maxw=acol_w-100)
    else: TC(c,ACOL,y7b+6,acol_w,"Specify Container Size",fs=FS_HEADER,bold=True)
    R(c,legend_right_x,y7b,legend_right_w,z3_rh)
    T(c,legend_right_x+2,y7b+10,"Container Size: (1) 1L, (2) 500mL,",fs=FS_LEGEND,bold=True)
    T(c,legend_right_x+2,y7b+3,"(3) 250mL, (4) 125mL, (5) 100mL, (6) Other",fs=FS_LEGEND,bold=True)

    # Preservative
    R(c,ACOL,y8b,acol_w+KELP_STRIP_W,z3_rh)
    pv=g("preservative_type")
    if pv: T(c,ACOL+2,y8b+10,"Identify Container Preservative Type:",fs=FS_LABEL); TV(c,ACOL+145,y8b+10,pv,maxw=acol_w-150)
    else: TC(c,ACOL,y8b+6,acol_w,"Identify Container Preservative Type",fs=FS_HEADER,bold=True)
    R(c,legend_right_x,y9b,legend_right_w,y7b-y9b)
    T(c,legend_right_x+2,y8b+10,"Preservative: (1) None, (2) HNO3,",fs=4.5)
    T(c,legend_right_x+2,y8b+4,"(3) H2SO4, (4) HCl, (5) NaOH,",fs=4.5)
    T(c,legend_right_x+2,y9b+14,"(6) Zn Acetate, (7) NaHSO4,",fs=4.5)
    T(c,legend_right_x+2,y9b+8,"(8) Sod.Thiosulfate, (9) Ascorbic Acid,",fs=4.5)
    T(c,legend_right_x+2,y9b+2,"(10) MeOH, (11) Other",fs=4.5)

    R(c,ACOL,y9b,acol_w+KELP_STRIP_W,z3_rh)
    TC(c,ACOL,y9b+6,acol_w,"Analysis Requested",fs=FS_HEADER,bold=True)

    # ZONE 4: TimeZone + Deliverables
    z4_top=z3_top-sec_h-4*z3_rh; lrh=14.5
    ya_b=z4_top-lrh
    R(c,LM,ya_b,ACOL-LM,lrh); T(c,LM+2,ya_b+3,"Sample Collection Time Zone:",fs=FS_LABEL)
    tz=g("time_zone","PT")
    for tn,tx in [("AK",130),("PT",160),("MT",186),("CT",210),("ET",236)]:
        CB(c,tx,ya_b+2,checked=(tn==tz)); T(c,tx+9,ya_b+3,tn,fs=FS_LABEL)

    dd_top=ya_b; dd_h=lrh*4; dd_w=130  # 4 rows (removed Level IV + Analysis row)
    reg_w=ACOL-LM-dd_w; rx=LM+dd_w

    R(c,LM,dd_top-dd_h,dd_w,dd_h); T(c,LM+2,dd_top-10,"Data Deliverables:",fs=FS_LABEL)
    sd=g("data_deliverable","Level I (Std)")
    # No Level IV per correction #5
    for lbl,bx,by in [
        ("Level I (Std)",LM+6,dd_top-24),("Level II",LM+68,dd_top-24),
        ("Level III",LM+6,dd_top-38),("Other",LM+6,dd_top-52)]:
        CB(c,bx,by,checked=(lbl==sd)); T(c,bx+9,by+1,lbl,fs=FS_LABEL)

    # Row 1: Regulatory + Reportable
    R(c,rx,dd_top-lrh,reg_w,lrh)
    T(c,rx+2,dd_top-lrh+3,"Regulatory Program (DW, RCRA, etc.):",fs=FS_LABEL)
    TV(c,rx+145,dd_top-lrh+3,g("regulatory_program"),maxw=50)
    rpx=rx+reg_w*0.65; T(c,rpx,dd_top-lrh+3,"Reportable",fs=FS_LABEL)
    rv=g("reportable")=="Yes"
    CB(c,rpx+38,dd_top-lrh+2,checked=rv); T(c,rpx+47,dd_top-lrh+3,"Yes",fs=FS_LABEL)
    CB(c,rpx+65,dd_top-lrh+2,checked=not rv); T(c,rpx+74,dd_top-lrh+3,"No",fs=FS_LABEL)

    # Row 2: Rush
    R(c,rx,dd_top-lrh*2,reg_w,lrh)
    T(c,rx+2,dd_top-lrh*2+3,"Rush (Pre-approval required):",fs=FS_LABEL,bold=True)
    sr=g("rush","Standard (5-10 Day)")
    for i,ro in enumerate(["Same Day","1 Day","2 Day","3 Day","4 Day"]):
        rrx=rx+120+i*40; CB(c,rrx,dd_top-lrh*2+2,checked=(ro==sr)); T(c,rrx+9,dd_top-lrh*2+3,ro,fs=FS_LEGEND)

    # Row 3: 5Day + PWSID
    h3=reg_w*0.5
    R(c,rx,dd_top-lrh*3,h3,lrh)
    CB(c,rx+4,dd_top-lrh*3+2,checked=("5 Day" in sr)); T(c,rx+14,dd_top-lrh*3+3,"5 Day",fs=FS_LABEL)
    T(c,rx+48,dd_top-lrh*3+3,"Other ____________",fs=FS_LABEL)
    R(c,rx+h3,dd_top-lrh*3,reg_w-h3,lrh)
    T(c,rx+h3+2,dd_top-lrh*3+3,"DW PWSID # or WW Permit #:",fs=FS_LEGEND)
    TV(c,rx+h3+108,dd_top-lrh*3+3,g("pwsid"),fs=7,maxw=reg_w-h3-113)

    # Row 4: Field Filtered (NO Analysis row)
    R(c,rx,dd_top-lrh*4,reg_w,lrh)
    T(c,rx+2,dd_top-lrh*4+3,"Field Filtered (if applicable):",fs=FS_LABEL)
    ff=g("field_filtered","No")
    CB(c,rx+120,dd_top-lrh*4+2,checked=(ff=="Yes")); T(c,rx+129,dd_top-lrh*4+3,"Yes",fs=FS_LABEL)
    CB(c,rx+150,dd_top-lrh*4+2,checked=(ff=="No")); T(c,rx+159,dd_top-lrh*4+3,"No",fs=FS_LABEL)

    ml_top=dd_top-dd_h; ml_h=10
    R(c,LM,ml_top-ml_h,ACOL-LM,ml_h)
    T(c,LM+2,ml_top-ml_h+2,"* Matrix: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water(SW), Other(OT)",fs=FS_LEGEND)

    # TABLE COLUMN HEADERS — always show Composite Start + End (same as original template)
    th_top=ml_top-ml_h; th_h=32; grp_h=12; sub_h=th_h-grp_h

    R(c,17.3,th_top-th_h,160.8,th_h)
    TC(c,17.3,th_top-th_h/2-3,160.8,"Customer Sample ID",fs=FS_HEADER,bold=True)
    R(c,178.1,th_top-th_h,31.5,th_h)
    TC(c,178.1,th_top-th_h/2+3,31.5,"Matrix",fs=FS_HEADER,bold=True)
    TC(c,178.1,th_top-th_h/2-7,31.5,"*",fs=FS_HEADER,bold=True)
    R(c,209.6,th_top-th_h,27.8,th_h)
    TC(c,209.6,th_top-th_h/2+3,27.8,"Comp /",fs=FS_HEADER,bold=True)
    TC(c,209.6,th_top-th_h/2-7,27.8,"Grab",fs=FS_HEADER,bold=True)

    # Composite Start
    R(c,237.4,th_top-grp_h,89.6,grp_h)
    TC(c,237.4,th_top-grp_h+2,89.6,"Composite Start",fs=FS_LABEL,bold=True)
    R(c,237.4,th_top-th_h,54.3,sub_h)
    TC(c,237.4,th_top-th_h+sub_h/2-3,54.3,"Date",fs=FS_LABEL,bold=True)
    R(c,291.7,th_top-th_h,35.3,sub_h)
    TC(c,291.7,th_top-th_h+sub_h/2-3,35.3,"Time",fs=FS_LABEL,bold=True)

    # Collected or Composite End
    R(c,327.0,th_top-grp_h,85.2,grp_h)
    TC(c,327.0,th_top-grp_h+2,85.2,"Collected or Composite End",fs=FS_LEGEND,bold=True)
    R(c,327.0,th_top-th_h,50.2,sub_h)
    TC(c,327.0,th_top-th_h+sub_h/2-3,50.2,"Date",fs=FS_LABEL,bold=True)
    R(c,377.2,th_top-th_h,35.0,sub_h)
    TC(c,377.2,th_top-th_h+sub_h/2-3,35.0,"Time",fs=FS_LEGEND,bold=True)

    R(c,412.2,th_top-th_h,19.9,th_h)
    TC(c,412.2,th_top-th_h/2+3,19.9,"#",fs=FS_HEADER,bold=True)
    TC(c,412.2,th_top-th_h/2-7,19.9,"Cont.",fs=FS_HEADER,bold=True)

    R(c,432.1,th_top-grp_h,44.7,grp_h)
    TC(c,432.1,th_top-grp_h+2,44.7,"Residual Chlorine",fs=FS_LEGEND,bold=True)
    R(c,432.1,th_top-th_h,23.1,sub_h)
    TC(c,432.1,th_top-th_h+sub_h/2-3,23.1,"Result",fs=FS_LEGEND,bold=True)
    R(c,455.2,th_top-th_h,21.6,sub_h)
    TC(c,455.2,th_top-th_h+sub_h/2-3,21.6,"Units",fs=FS_LEGEND,bold=True)

    # Tall vertical analysis headers (dynamic, with wrapping)
    tall_bot=th_top-th_h; tall_h=z4_top-tall_bot

    cat_col_indices={}
    for ci,col_info in enumerate(dyn_cols):
        cat=col_info["cat_name"]
        if cat not in cat_col_indices: cat_col_indices[cat]=[]
        cat_col_indices[cat].append(ci)

    for ci in range(num_acols):
        ax0=AX[ci]; ax1=AX[ci+1] if ci+1<len(AX) else KSX; aw=ax1-ax0
        R(c,ax0,tall_bot,aw,tall_h)
        if ci<len(dyn_cols):
            col_info=dyn_cols[ci]
            label=col_info["label"]; method=col_info["method"]
            VTEXT_WRAP(c,ax0+aw/2+4,tall_bot+3,label,fs=min(FS_HEADER,5.5 if len(label)>60 else FS_HEADER),bold=True,col_w=aw,max_h=tall_h-6)
            VTEXT(c,ax0+aw/2-4,tall_bot+3,method,fs=FS_LEGEND,bold=False)

    # KELP Use Only strip
    R(c,KSX,tall_bot,KELP_STRIP_W,tall_h)
    VTEXT(c,KSX+KELP_STRIP_W/2+2,tall_bot+tall_h/2-22,"KELP Use Only",fs=FS_LEGEND,bold=True)

    # Right side fields
    sbf_h=18; fy=z4_top
    for lbl,val in [("Project Mgr.:",g("project_manager")),("AcctNum / Client ID:",g("acct_num")),
                    ("Table #:",g("table_number")),("Profile / Template:",g("profile_template")),
                    ("Prelog / Bottle Ord. ID:",g("prelog_id"))]:
        R(c,SBX,fy-sbf_h,COMMENT_W,sbf_h)
        T(c,SBX+2,fy-sbf_h+8,lbl,fs=FS_LABEL)
        lw2=stringWidth(lbl,"Helvetica",FS_LABEL)+3
        TV(c,SBX+lw2,fy-sbf_h+8,val,fs=7,maxw=COMMENT_W-lw2-2)
        fy-=sbf_h
    sc_h=fy-tall_bot; R(c,SBX,tall_bot,COMMENT_W,sc_h)
    TC(c,SBX,tall_bot+sc_h/2-3,COMMENT_W,"Sample Comment",fs=FS_HEADER,bold=True)

    # PNC column
    R(c,PNCX,tall_bot,PNC_W,z4_top-tall_bot)
    VTEXT(c,PNCX+PNC_W/2+3,tall_bot+5,"Preservation non-conformance",fs=4.5,bold=False)
    VTEXT(c,PNCX+PNC_W/2-5,tall_bot+5,"identified for sample.",fs=4.5,bold=False)

    # SAMPLE DATA ROWS
    data_rh=19; max_rows=10
    data_top=tall_bot
    SECTION_LABEL(c,LM,data_top-sec_h,ACOL-LM,sec_h,"SAMPLE INFORMATION"); data_top-=sec_h
    samples=d.get("samples",[])

    for ri in range(max_rows):
        s=samples[ri] if ri<len(samples) else {}
        ryb=data_top-(ri+1)*data_rh
        if ri%2==1: c.setFillColor(ROW_SHADE); c.rect(LM,ryb,RM-LM,data_rh,fill=1,stroke=0)
        c.setFont("Helvetica",4.5); c.setFillColor(black); c.drawCentredString(LM+4,ryb+7,str(ri+1))

        for x0,x1,val,align,fs in [
            (17.3,178.1,s.get("sample_id",""),"left",FS_VALUE),
            (178.1,209.6,s.get("matrix",""),"center",FS_VALUE),
            (209.6,237.4,s.get("comp_grab",""),"center",6),
            (237.4,291.7,s.get("start_date",""),"center",6),
            (291.7,327.0,s.get("start_time",""),"center",6),
            (327.0,377.2,s.get("end_date","") or s.get("collected_date",""),"center",6),
            (377.2,412.2,s.get("end_time","") or s.get("collected_time",""),"center",6),
            (412.2,432.1,s.get("num_containers",""),"center",FS_VALUE),
            (432.1,455.2,s.get("res_cl_result",""),"center",6),
            (455.2,476.8,s.get("res_cl_units",""),"center",FS_LEGEND)]:
            R(c,x0,ryb,x1-x0,data_rh)
            if val:
                if align=="center": TC(c,x0,ryb+5,x1-x0,str(val),fs=fs,bold=True)
                else: TV(c,x0+10,ryb+5,str(val),fs=fs,maxw=x1-x0-14)

        # Analysis X marks
        sa=s.get("analyses",{})
        if isinstance(sa,list): sa={cat:[] for cat in sa}
        for ci in range(num_acols):
            ax0=AX[ci]; ax1=AX[ci+1] if ci+1<len(AX) else KSX
            R(c,ax0,ryb,ax1-ax0,data_rh)
        for cat_name,al in sa.items():
            if not al or cat_name not in cat_col_indices: continue
            ci_idx=cat_col_indices[cat_name][0]
            if ci_idx<num_acols:
                ax0=AX[ci_idx]; ax1=AX[ci_idx+1] if ci_idx+1<len(AX) else KSX
                TC(c,ax0,ryb+5,ax1-ax0,"X",fs=FS_VALUE,bold=True)

        R(c,KSX,ryb,KELP_STRIP_W,data_rh)
        R(c,SBX,ryb,COMMENT_W,data_rh)
        cmt=s.get("comment","")
        if cmt: TV(c,SBX+2,ryb+5,cmt,fs=FS_LEGEND,maxw=COMMENT_W-4)
        R(c,PNCX,ryb,PNC_W,data_rh)

    # BOTTOM ZONE
    bot_top=data_top-max_rows*data_rh
    SECTION_LABEL(c,LM,bot_top-sec_h,RM-LM,sec_h,"CHAIN OF CUSTODY RECORD / LABORATORY RECEIVING"); bot_top-=sec_h

    inst_h=26; half_w=(RM-LM)/2
    R(c,LM,bot_top-inst_h,half_w,inst_h)
    T(c,LM+2,bot_top-8,"Additional Instructions from KELP:",fs=FS_LABEL,bold=True)
    TV(c,LM+4,bot_top-18,g("additional_instructions"),fs=6,maxw=half_w-8)
    R(c,LM+half_w,bot_top-inst_h,half_w,inst_h)
    T(c,LM+half_w+2,bot_top-8,"Customer Remarks / Special Conditions / Possible Hazards:",fs=FS_LABEL,bold=True)
    TV(c,LM+half_w+4,bot_top-18,g("customer_remarks"),fs=6,maxw=half_w-8)

    lr_top=bot_top-inst_h; lr_h=13
    R(c,LM,lr_top-lr_h,RM-LM,lr_h)
    T(c,LM+2,lr_top-9,"# Coolers:",fs=FS_LABEL); TV(c,LM+42,lr_top-9,g("num_coolers"),fs=7)
    T(c,LM+70,lr_top-9,"Thermometer ID:",fs=FS_LABEL); TV(c,LM+130,lr_top-9,g("thermometer_id"),fs=7)
    T(c,LM+185,lr_top-9,"Temp. (\u00b0C):",fs=FS_LABEL); TV(c,LM+225,lr_top-9,g("temperature"),fs=7)
    roi=g("received_on_ice","Yes")
    T(c,LM+310,lr_top-9,"Sample Received on ice:",fs=FS_LABEL)
    CB(c,LM+400,lr_top-11,checked=(roi=="Yes")); T(c,LM+409,lr_top-9,"Yes",fs=FS_LABEL)
    CB(c,LM+432,lr_top-11,checked=(roi=="No")); T(c,LM+441,lr_top-9,"No",fs=FS_LABEL)

    rel_top=lr_top-lr_h; rel_h=12
    for row_i in range(3):
        ryb=rel_top-(row_i+1)*rel_h; rw1=210; dtw=90
        R(c,LM,ryb,rw1,rel_h); T(c,LM+2,ryb+3,"Relinquished by/Company: (Signature)",fs=FS_LEGEND)
        R(c,LM+rw1,ryb,dtw,rel_h); T(c,LM+rw1+2,ryb+3,"Date/Time:",fs=FS_LEGEND)
        rcx=LM+rw1+dtw; R(c,rcx,ryb,rw1,rel_h); T(c,rcx+2,ryb+3,"Received by/Company: (Signature)",fs=FS_LEGEND)
        R(c,rcx+rw1,ryb,dtw,rel_h); T(c,rcx+rw1+2,ryb+3,"Date/Time:",fs=FS_LEGEND)
        rsx=rcx+rw1+dtw; rsw=RM-rsx; R(c,rsx,ryb,rsw,rel_h)
        if row_i==0: T(c,rsx+2,ryb+3,"Tracking #:",fs=FS_LEGEND); TV(c,rsx+42,ryb+3,g("tracking_number"),fs=6,maxw=rsw-46)
        elif row_i==1:
            dm=g("delivery_method"); T(c,rsx+2,ryb+3,"Delivered by:",fs=FS_LEGEND); avail=rsw-55
            for dn,dx in [("In-Person",rsx+52),("FedEx",rsx+52+avail*0.30),("UPS",rsx+52+avail*0.55),("Other",rsx+52+avail*0.75)]:
                CB(c,dx,ryb+1,checked=(dn==dm),sz=6); T(c,dx+8,ryb+3,dn,fs=FS_LEGEND)

    disc_y=rel_top-3*rel_h-8
    T(c,LM+5,disc_y,"Submitting a sample via this chain of custody constitutes acknowledgment and acceptance of the KELP\u2019s Terms and Conditions",fs=FS_LABEL)
    form_bot=disc_y-3
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER); c.rect(LM,form_bot,RM-LM,hdr_top-form_bot,fill=0,stroke=1)
    _footer(c,1,total_pages,coc_id)

    # ═══════════════════════════════════════
    # PAGE 2: INSTRUCTIONS
    # ═══════════════════════════════════════
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
    for l,d2 in [("Company Name:","Your company\u2019s name."),("Street Address:","Your mailing address."),
        ("Contact/Report To:","Person designated to receive results."),
        ("Customer Project # and Project Name:","Your project reference number and name."),
        ("Site Collection Info/Facility ID:","Project location or facility ID."),
        ("Time Zone:","Sample collection time zone (e.g., AK, PT, MT, CT, ET)."),
        ("Purchase Order #:","Your PO number for invoicing."),("Invoice To:","Contact person for the invoice."),
        ("Invoice Email:","Email address for the invoice."),("Phone #:","Your contact phone number."),
        ("E-mail:","Your email for correspondence and the final report."),
        ("Data Deliverable:","Required data deliverable level."),
        ("Field Filtered:","Indicate if samples were filtered in the field (Yes/No)."),
        ("Quote #:","Quote number, if applicable."),
        ("DW PWSID # or WW Permit #:","Relevant permit numbers, if applicable.")]:
        y=bullet_item(y,l,d2)

    y-=4; y=sec_heading(y,"2. Sample Information:")
    for l,d2 in [("Customer Sample ID:","Unique sample identifier for the report."),
        ("Collected Date:","Sample collection date."),("Collected Time:","Sample collection time."),
        ("Comp/Grab:","GRAB for single-point; COMP for combined samples."),
        ("Matrix:","Sample type (DW, GW, WW, P, SW, OT)."),("Container Size:","Specify size (e.g., 1L, 500mL).")]:
        y=bullet_item(y,l,d2)

    y2=PH-85
    for l,d2 in [("Container Preservative Type:","Specify preservative (None, HNO3, H2SO4, etc.)."),
        ("Analysis Requested:","Mark categories for each sample. Analyte names shown in column headers."),
        ("Sample Comment:","Notes about individual samples; identify MS/MSD samples here."),
        ("Residual Chlorine:","Record results and units if measured.")]:
        y2=bullet_item(y2,l,d2,cx=col2_x,w=cw)

    y2-=4; y2=sec_heading(y2,"3. Additional Information & Instructions:",cx=col2_x)
    for l,d2 in [("Customer Remarks/Hazards:","Note special instructions or hazards."),
        ("Rush Request:","For expedited results; pre-approval required."),
        ("Relinquished By/Received By:","Sign and date at each transfer of custody.")]:
        y2=bullet_item(y2,l,d2,cx=col2_x,w=cw)

    y2-=4; y2=sec_heading(y2,"4. Sample Acceptance Policy Summary:",cx=col2_x)
    c.setFont("Helvetica",bfs); c.setFillColor(black); c.drawString(col2_x,y2,"For samples to be accepted, ensure:"); y2-=14
    for item in ["Complete COC documentation.","Readable, unique sample ID on containers (indelible ink).",
        "Appropriate containers and sufficient volume.","Receipt within holding time and temperature requirements.",
        "Containers are in good condition, seals intact (if used).",
        "Proper preservation, no headspace in volatile water samples.","Adequate volume for MS/MSD if required."]:
        c.setFont("Helvetica",bfs); c.drawString(col2_x+bul_,y2,"\u2022  "+item); y2-=13

    y2-=6; c.setFont("Helvetica",8.5); c.setFillColor(black)
    closing="Failure to meet these may result in data qualifiers. A detailed policy is available from your Project Manager. Submitting samples implies acceptance of KELP Terms and Conditions."
    words=closing.split(); line=""; cy=y2
    for word in words:
        test=line+(" " if line else "")+word
        if stringWidth(test,"Helvetica",8.5)<=cw: line=test
        else: c.drawString(col2_x,cy,line); cy-=12; line=word
    if line: c.drawString(col2_x,cy,line)

    _footer(c,2,total_pages,coc_id)
    c.showPage(); c.save(); buf.seek(0)
    return buf, coc_id

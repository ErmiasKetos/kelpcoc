"""
coc_pdf_engine.py — KELP COC PDF Generator v15
Generates the 3-page Chain-of-Custody PDF with individual analyte detail.
"""
import io, os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

PW, PH = landscape(letter)

KELP_BLUE = HexColor("#1F4E79")
HDR_BLUE  = HexColor("#1F4E79")
ROW_SHADE = HexColor("#F2F2F2")
SECTION_BG = HexColor("#E8ECF0")

LW_OUTER = 1.0; LW_SECTION = 0.7; LW_INNER = 0.3

DOC_ID = "KELP-QMS-FORM-001"
DOC_TITLE = "Chain-of-Custody"
DOC_VERSION = "1.1"
DOC_EFF_DATE = "February 19, 2026"

FS_TITLE=11; FS_LABEL=6; FS_VALUE=8; FS_HEADER=6.5
FS_LEGEND=5; FS_FOOTER=5; FS_SUBTITLE=7

LM=17.3; RM=785.5; COL2=237.4; ACOL=476.8
KELP_STRIP=665.0; SB=674.6; PNC=763.4

AX=[476.8,494.8,513.5,532.2,551.2,569.9,588.7,607.4,626.1,645.3,664.3]

from coc_catalog import KELP_ANALYTE_CATALOG, DEFAULT_ANALYSIS_COLUMNS, CAT_SHORT_MAP

# ── Drawing helpers ──
def R(c,x,y,w,h,fill=None,lw=LW_INNER):
    if fill: c.setFillColor(fill); c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(lw); c.rect(x,y,w,h,fill=0,stroke=1)

def SECTION_LABEL(c,x,y,w,h,text):
    c.setFillColor(SECTION_BG); c.rect(x,y,w,h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(LW_SECTION); c.rect(x,y,w,h,fill=0,stroke=1)
    c.setFont("Helvetica-Bold",5.5); c.setFillColor(HDR_BLUE)
    c.drawCentredString(x+w/2, y+h/2-2, text)

def T(c,x,y,txt,fs=FS_LABEL,bold=False,font="Helvetica",maxw=None,color=black):
    if not txt: return
    fn="Helvetica-Bold" if bold else font; s=float(fs); t=str(txt)
    if maxw:
        while s>4 and stringWidth(t,fn,s)>maxw: s-=0.3
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

def HLINE(c,x1,x2,y,lw=0.4):
    c.setStrokeColor(black); c.setLineWidth(lw); c.line(x1,y,x2,y)

def _footer(c, page_num, total_pages):
    c.setFont("Helvetica",FS_FOOTER); c.setFillColor(black)
    c.setStrokeColor(black); c.setLineWidth(0.3); c.line(LM,14,RM,14)
    c.drawString(LM,6, f"{DOC_ID}-{DOC_TITLE.replace(' ','')}  |  Version {DOC_VERSION}  |  Effective: {DOC_EFF_DATE}")
    c.drawRightString(RM,6, f"CONTROLLED DOCUMENT - Do not copy without authorization  |  Page {page_num} of {total_pages}")


def generate_coc_pdf(data, logo_path=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PW,PH))
    c.setTitle("KELP Chain-of-Custody")
    d = data
    g = lambda k,dflt="": d.get(k,dflt) or dflt
    acols = list(DEFAULT_ANALYSIS_COLUMNS)

    # Determine total pages
    total_pages = 3

    # ═══════════════════════════════════════════════════════════
    # PAGE 1: COC FORM
    # ═══════════════════════════════════════════════════════════
    hdr_top=598; hdr_bot=560; hdr_h=hdr_top-hdr_bot
    c.setFillColor(KELP_BLUE); c.rect(LM,hdr_bot,RM-LM,hdr_h,fill=1,stroke=0)
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER); c.rect(LM,hdr_bot,RM-LM,hdr_h,fill=0,stroke=1)

    if logo_path and os.path.exists(logo_path):
        try: c.drawImage(ImageReader(logo_path),LM+4,hdr_bot+4,width=90,height=hdr_h-8,preserveAspectRatio=True,mask='auto')
        except: pass

    c.setFont("Helvetica-Bold",FS_TITLE); c.setFillColor(white)
    c.drawCentredString((LM+ACOL)/2, hdr_top-16, "CHAIN-OF-CUSTODY RECORD")
    c.setFont("Helvetica",FS_SUBTITLE); c.setFillColor(white)
    c.drawCentredString((LM+ACOL)/2, hdr_top-28, "Chain-of-Custody is a LEGAL DOCUMENT \u2014 Complete all relevant fields")

    kelp_x=570
    c.setFillColor(white); c.rect(kelp_x,hdr_bot+2,RM-kelp_x-2,hdr_h-4,fill=1,stroke=0)
    R(c,kelp_x,hdr_bot+2,RM-kelp_x-2,hdr_h-4,lw=LW_SECTION)
    T(c,kelp_x+3,hdr_top-12,"KELP USE ONLY",fs=6.5,bold=True,color=HDR_BLUE)
    T(c,kelp_x+3,hdr_top-24,"KELP Ordering ID:",fs=FS_LABEL)
    HLINE(c,kelp_x+68,RM-6,hdr_top-26,lw=0.3)
    kid=g("kelp_ordering_id")
    if kid: TV(c,kelp_x+70,hdr_top-24,kid)

    # CLIENT INFO
    ci_top=hdr_bot; rh=10; lw_=COL2-LM; cw_=ACOL-COL2; sec_h=8
    SECTION_LABEL(c,LM,ci_top-sec_h,ACOL-LM,sec_h,"CLIENT INFORMATION"); ci_top-=sec_h

    def ci_row(ri,ll,lv,cl,cv):
        yt=ci_top-ri*rh; yb=yt-rh
        R(c,LM,yb,lw_,rh)
        if ll: T(c,LM+2,yb+2,ll,fs=FS_LABEL); tw=stringWidth(ll,"Helvetica",FS_LABEL)+3; TV(c,LM+tw,yb+1,lv,maxw=lw_-tw-2)
        R(c,COL2,yb,cw_,rh)
        if cl: T(c,COL2+2,yb+2,cl,fs=FS_LABEL); tw=stringWidth(cl,"Helvetica",FS_LABEL)+3; TV(c,COL2+tw,yb+1,cv,maxw=cw_-tw-2)

    ci_row(0,"Company Name:",g("company_name"),"Contact/Report To:",g("contact_name"))
    ci_row(1,"Street Address:",g("street_address"),"Phone #:",g("phone"))
    ci_row(2,"","","E-Mail:",g("email"))
    ci_row(3,"","","Cc E-Mail:",g("cc_email"))
    ci_row(4,"Customer Project #:",g("project_number"),"Invoice to:",g("invoice_to"))
    rw_=RM-ACOL; R(c,ACOL,ci_top-5*rh,rw_,5*rh)

    # PROJECT DETAILS
    z3_top=ci_top-5*rh; z3_rh=19
    SECTION_LABEL(c,LM,z3_top-sec_h,ACOL-LM,sec_h,"PROJECT DETAILS"); z3_top-=sec_h
    z3_bot=z3_top-4*z3_rh

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

    acol_w=AX[10]-AX[0]; legend_x=KELP_STRIP; legend_w=RM-legend_x
    R(c,ACOL,y6b,KELP_STRIP-ACOL,z3_rh); R(c,legend_x,y6b,legend_w,z3_rh)

    # Container size
    R(c,AX[0],y7b,acol_w,z3_rh)
    cs_val=g("container_size")
    if cs_val: T(c,AX[0]+2,y7b+10,"Specify Container Size:",fs=FS_LABEL); TV(c,AX[0]+95,y7b+10,cs_val,maxw=acol_w-100)
    else: TC(c,AX[0],y7b+6,acol_w,"Specify Container Size",fs=FS_HEADER,bold=True)
    R(c,legend_x,y7b,legend_w,z3_rh)
    T(c,legend_x+2,y7b+10,"Container Size: (1) 1L, (2) 500mL,",fs=FS_LEGEND,bold=True)
    T(c,legend_x+2,y7b+3,"(3) 250mL, (4) 125mL, (5) 100mL, (6) Other",fs=FS_LEGEND,bold=True)

    # Preservative
    R(c,AX[0],y8b,acol_w,z3_rh)
    pres_val=g("preservative_type")
    if pres_val: T(c,AX[0]+2,y8b+10,"Identify Container Preservative Type:",fs=FS_LABEL); TV(c,AX[0]+145,y8b+10,pres_val,maxw=acol_w-150)
    else: TC(c,AX[0],y8b+6,acol_w,"Identify Container Preservative Type",fs=FS_HEADER,bold=True)
    R(c,legend_x,y9b,legend_w,y7b-y9b)
    T(c,legend_x+2,y8b+10,"Preservative: (1) None, (2) HNO3,",fs=4.5)
    T(c,legend_x+2,y8b+4,"(3) H2SO4, (4) HCl, (5) NaOH,",fs=4.5)
    T(c,legend_x+2,y9b+14,"(6) Zn Acetate, (7) NaHSO4,",fs=4.5)
    T(c,legend_x+2,y9b+8,"(8) Sod.Thiosulfate, (9) Ascorbic Acid,",fs=4.5)
    T(c,legend_x+2,y9b+2,"(10) MeOH, (11) Other",fs=4.5)

    R(c,AX[0],y9b,acol_w,z3_rh)
    TC(c,AX[0],y9b+6,acol_w,"Analysis Requested  (see Page 2 for detail)",fs=FS_HEADER,bold=True)

    # ZONE 4: TimeZone + Deliverables
    z4_top=z3_bot; lrh=14.5
    ya_b=z4_top-lrh
    R(c,LM,ya_b,ACOL-LM,lrh); T(c,LM+2,ya_b+3,"Sample Collection Time Zone:",fs=FS_LABEL)
    tz=g("time_zone","PT")
    for tn,tx in [("AK",130),("PT",160),("MT",186),("CT",210),("ET",236)]:
        CB(c,tx,ya_b+2,checked=(tn==tz)); T(c,tx+9,ya_b+3,tn,fs=FS_LABEL)

    dd_top=ya_b; dd_h=lrh*5; dd_w=130; reg_w=ACOL-LM-dd_w; rx=LM+dd_w
    R(c,LM,dd_top-dd_h,dd_w,dd_h); T(c,LM+2,dd_top-10,"Data Deliverables:",fs=FS_LABEL)
    sd=g("data_deliverable","Level I (Std)")
    for lbl,bx,by in [("Level I (Std)",LM+6,dd_top-24),("Level II",LM+68,dd_top-24),("Level III",LM+6,dd_top-38),("Level IV",LM+6,dd_top-52),("Other",LM+6,dd_top-66)]:
        CB(c,bx,by,checked=(lbl==sd)); T(c,bx+9,by+1,lbl,fs=FS_LABEL)

    R(c,rx,dd_top-lrh,reg_w,lrh); T(c,rx+2,dd_top-lrh+3,"Regulatory Program (DW, RCRA, etc.):",fs=FS_LABEL)
    TV(c,rx+145,dd_top-lrh+3,g("regulatory_program"),maxw=50)
    rpx=rx+reg_w*0.65; T(c,rpx,dd_top-lrh+3,"Reportable",fs=FS_LABEL)
    rv=g("reportable")=="Yes"
    CB(c,rpx+38,dd_top-lrh+2,checked=rv); T(c,rpx+47,dd_top-lrh+3,"Yes",fs=FS_LABEL)
    CB(c,rpx+65,dd_top-lrh+2,checked=not rv); T(c,rpx+74,dd_top-lrh+3,"No",fs=FS_LABEL)

    R(c,rx,dd_top-lrh*2,reg_w,lrh); T(c,rx+2,dd_top-lrh*2+3,"Rush (Pre-approval required):",fs=FS_LABEL,bold=True)
    sr=g("rush","Standard (5-10 Day)")
    for i,ro in enumerate(["Same Day","1 Day","2 Day","3 Day","4 Day"]):
        rrx=rx+120+i*40; CB(c,rrx,dd_top-lrh*2+2,checked=(ro==sr)); T(c,rrx+9,dd_top-lrh*2+3,ro,fs=FS_LEGEND)

    h3=reg_w*0.5
    R(c,rx,dd_top-lrh*3,h3,lrh); CB(c,rx+4,dd_top-lrh*3+2,checked=("5 Day" in sr)); T(c,rx+14,dd_top-lrh*3+3,"5 Day",fs=FS_LABEL)
    T(c,rx+48,dd_top-lrh*3+3,"Other ____________",fs=FS_LABEL)
    R(c,rx+h3,dd_top-lrh*3,reg_w-h3,lrh); T(c,rx+h3+2,dd_top-lrh*3+3,"DW PWSID # or WW Permit #:",fs=FS_LEGEND)
    TV(c,rx+h3+108,dd_top-lrh*3+3,g("pwsid"),fs=7,maxw=reg_w-h3-113)

    R(c,rx,dd_top-lrh*4,reg_w,lrh); T(c,rx+2,dd_top-lrh*4+3,"Field Filtered (if applicable):",fs=FS_LABEL)
    ff=g("field_filtered","No")
    CB(c,rx+120,dd_top-lrh*4+2,checked=(ff=="Yes")); T(c,rx+129,dd_top-lrh*4+3,"Yes",fs=FS_LABEL)
    CB(c,rx+150,dd_top-lrh*4+2,checked=(ff=="No")); T(c,rx+159,dd_top-lrh*4+3,"No",fs=FS_LABEL)

    R(c,rx,dd_top-lrh*5,reg_w,lrh); T(c,rx+2,dd_top-lrh*5+3,"Analysis:",fs=FS_LABEL)
    TV(c,rx+36,dd_top-lrh*5+3,g("analysis_profile_template"),maxw=reg_w-40)

    ml_top=dd_top-dd_h; ml_h=10
    R(c,LM,ml_top-ml_h,ACOL-LM,ml_h)
    T(c,LM+2,ml_top-ml_h+2,"* Matrix: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water(SW), Other(OT)",fs=FS_LEGEND)

    # Table column headers
    th_top=ml_top-ml_h; th_h=32; grp_h=12
    for x0,x1,lbl in [(17.3,178.1,"Customer Sample ID"),(178.1,209.6,"Matrix\n*"),(209.6,237.4,"Comp /\nGrab"),(412.2,432.1,"#\nCont.")]:
        R(c,x0,th_top-th_h,x1-x0,th_h); lines=lbl.split("\n")
        if len(lines)==1: TC(c,x0,th_top-th_h/2-3,x1-x0,lines[0],fs=FS_HEADER,bold=True)
        else: TC(c,x0,th_top-th_h/2+3,x1-x0,lines[0],fs=FS_HEADER,bold=True); TC(c,x0,th_top-th_h/2-7,x1-x0,lines[1],fs=FS_HEADER,bold=True)

    sub_h=th_h-grp_h
    R(c,237.4,th_top-grp_h,327.0-237.4,grp_h); TC(c,237.4,th_top-grp_h+2,327.0-237.4,"Composite Start",fs=FS_LABEL,bold=True)
    R(c,237.4,th_top-th_h,291.7-237.4,sub_h); TC(c,237.4,th_top-th_h+sub_h/2-3,291.7-237.4,"Date",fs=FS_LABEL,bold=True)
    R(c,291.7,th_top-th_h,327.0-291.7,sub_h); TC(c,291.7,th_top-th_h+sub_h/2-3,327.0-291.7,"Time",fs=FS_LABEL,bold=True)
    R(c,327.0,th_top-grp_h,412.2-327.0,grp_h); TC(c,327.0,th_top-grp_h+2,412.2-327.0,"Collected or Composite End",fs=FS_LEGEND,bold=True)
    R(c,327.0,th_top-th_h,377.2-327.0,sub_h); TC(c,327.0,th_top-th_h+sub_h/2-3,377.2-327.0,"Date",fs=FS_LABEL,bold=True)
    R(c,377.2,th_top-th_h,412.2-377.2,sub_h); TC(c,377.2,th_top-th_h+sub_h/2-3,412.2-377.2,"Time",fs=FS_LEGEND,bold=True)
    R(c,432.1,th_top-grp_h,476.8-432.1,grp_h); TC(c,432.1,th_top-grp_h+2,476.8-432.1,"Residual Chlorine",fs=FS_LEGEND,bold=True)
    R(c,432.1,th_top-th_h,455.2-432.1,sub_h); TC(c,432.1,th_top-th_h+sub_h/2-3,455.2-432.1,"Result",fs=FS_LEGEND,bold=True)
    R(c,455.2,th_top-th_h,476.8-455.2,sub_h); TC(c,455.2,th_top-th_h+sub_h/2-3,476.8-455.2,"Units",fs=FS_LEGEND,bold=True)

    # Tall vertical analysis headers
    tall_bot=th_top-th_h; tall_h=z4_top-tall_bot
    for ai in range(10):
        ax0=AX[ai]; ax1=AX[ai+1]; aw=ax1-ax0; R(c,ax0,tall_bot,aw,tall_h)
        label=acols[ai] if ai<len(acols) else ""
        if label:
            lines=label.split("\n")
            if len(lines)>=2: VTEXT(c,ax0+aw/2+4,tall_bot+3,lines[0],fs=FS_HEADER,bold=True); VTEXT(c,ax0+aw/2-4,tall_bot+3,lines[1],fs=FS_LEGEND,bold=False)
            else: VTEXT(c,ax0+aw/2+2,tall_bot+3,label,fs=FS_HEADER if len(label)<20 else FS_LEGEND+0.5,bold=True)

    strip_w=SB-KELP_STRIP; R(c,KELP_STRIP,tall_bot,strip_w,tall_h)
    VTEXT(c,KELP_STRIP+strip_w/2+2,tall_bot+tall_h/2-25,"KELP Use Only",fs=FS_LEGEND,bold=True)

    sbf_w=PNC-SB; sbf_h=18; fy=z4_top
    for lbl,val in [("Project Mgr.:",g("project_manager")),("AcctNum / Client ID:",g("acct_num")),("Table #:",g("table_number")),("Profile / Template:",g("profile_template")),("Prelog / Bottle Ord. ID:",g("prelog_id"))]:
        R(c,SB,fy-sbf_h,sbf_w,sbf_h); T(c,SB+2,fy-sbf_h+8,lbl,fs=FS_LABEL)
        lw2=stringWidth(lbl,"Helvetica",FS_LABEL)+3; TV(c,SB+lw2,fy-sbf_h+8,val,fs=7,maxw=sbf_w-lw2-2); fy-=sbf_h
    sc_h=fy-tall_bot; R(c,SB,tall_bot,sbf_w,sc_h); TC(c,SB,tall_bot+sc_h/2-3,sbf_w,"Sample Comment",fs=FS_HEADER,bold=True)

    data_rh=19; max_rows=10
    pnc_bot=tall_bot-max_rows*data_rh; R(c,PNC,pnc_bot,RM-PNC,z4_top-pnc_bot)
    VTEXT(c,PNC+(RM-PNC)/2+2,pnc_bot+5,"Preservation non-conformance identified for sample.",fs=4)

    # SAMPLE DATA ROWS
    data_top=tall_bot; samples=d.get("samples",[])
    SECTION_LABEL(c,LM,data_top-sec_h,ACOL-LM,sec_h,"SAMPLE INFORMATION"); data_top-=sec_h

    cat_col_map={}
    for ci_idx,col_label in enumerate(acols):
        short=col_label.split("\n")[0] if col_label else ""
        if short: cat_col_map[short]=ci_idx

    for ri in range(max_rows):
        s=samples[ri] if ri<len(samples) else {}
        ry=data_top-ri*data_rh; ryb=ry-data_rh
        if ri%2==1: c.setFillColor(ROW_SHADE); c.rect(LM,ryb,RM-LM,data_rh,fill=1,stroke=0)
        c.setFont("Helvetica",4.5); c.setFillColor(black); c.drawCentredString(LM+4,ryb+7,str(ri+1))

        for x0,x1,val,align,fs in [(17.3,178.1,s.get("sample_id",""),"left",FS_VALUE),(178.1,209.6,s.get("matrix",""),"center",FS_VALUE),(209.6,237.4,s.get("comp_grab",""),"center",6),(237.4,291.7,s.get("start_date",""),"center",6),(291.7,327.0,s.get("start_time",""),"center",6),(327.0,377.2,s.get("end_date",""),"center",6),(377.2,412.2,s.get("end_time",""),"center",6),(412.2,432.1,s.get("num_containers",""),"center",FS_VALUE),(432.1,455.2,s.get("res_cl_result",""),"center",6),(455.2,476.8,s.get("res_cl_units",""),"center",FS_LEGEND)]:
            R(c,x0,ryb,x1-x0,data_rh)
            if val:
                if align=="center": TC(c,x0,ryb+5,x1-x0,str(val),fs=fs,bold=True)
                else: TV(c,x0+10,ryb+5,str(val),fs=fs,maxw=x1-x0-14)

        s_analyses=s.get("analyses",{})
        if isinstance(s_analyses,list): s_analyses={cat:[] for cat in s_analyses}
        for ai in range(10): R(c,AX[ai],ryb,AX[ai+1]-AX[ai],data_rh)

        for cat_name,analyte_list in s_analyses.items():
            if not analyte_list: continue
            short=CAT_SHORT_MAP.get(cat_name,cat_name)
            if short in cat_col_map:
                ci_idx=cat_col_map[short]
                if ci_idx<10: TC(c,AX[ci_idx],ryb+5,AX[ci_idx+1]-AX[ci_idx],"X",fs=FS_VALUE,bold=True)

        R(c,KELP_STRIP,ryb,strip_w,data_rh); R(c,SB,ryb,sbf_w,data_rh)
        cmt=s.get("comment","")
        if cmt: TV(c,SB+2,ryb+5,cmt,fs=FS_LEGEND,maxw=sbf_w-4)

    # BOTTOM ZONE
    bot_top=data_top-max_rows*data_rh
    SECTION_LABEL(c,LM,bot_top-sec_h,RM-LM,sec_h,"CHAIN OF CUSTODY RECORD / LABORATORY RECEIVING"); bot_top-=sec_h

    inst_h=28; half_w=(RM-LM)/2
    R(c,LM,bot_top-inst_h,half_w,inst_h); T(c,LM+2,bot_top-8,"Additional Instructions from KELP:",fs=FS_LABEL,bold=True)
    TV(c,LM+4,bot_top-20,g("additional_instructions"),fs=6,maxw=half_w-8)
    R(c,LM+half_w,bot_top-inst_h,half_w,inst_h); T(c,LM+half_w+2,bot_top-8,"Customer Remarks / Special Conditions / Possible Hazards:",fs=FS_LABEL,bold=True)
    TV(c,LM+half_w+4,bot_top-20,g("customer_remarks"),fs=6,maxw=half_w-8)

    lr_top=bot_top-inst_h; lr_h=14
    R(c,LM,lr_top-lr_h,RM-LM,lr_h)
    T(c,LM+2,lr_top-10,"# Coolers:",fs=FS_LABEL); TV(c,LM+42,lr_top-10,g("num_coolers"),fs=7)
    T(c,LM+70,lr_top-10,"Thermometer ID:",fs=FS_LABEL); TV(c,LM+130,lr_top-10,g("thermometer_id"),fs=7)
    T(c,LM+185,lr_top-10,"Temp. (\u00b0C):",fs=FS_LABEL); TV(c,LM+225,lr_top-10,g("temperature"),fs=7)
    roi=g("received_on_ice","Yes")
    T(c,LM+310,lr_top-10,"Sample Received on ice:",fs=FS_LABEL)
    CB(c,LM+400,lr_top-12,checked=(roi=="Yes")); T(c,LM+409,lr_top-10,"Yes",fs=FS_LABEL)
    CB(c,LM+432,lr_top-12,checked=(roi=="No")); T(c,LM+441,lr_top-10,"No",fs=FS_LABEL)

    rel_top=lr_top-lr_h; rel_h=13
    for row_i in range(3):
        ry=rel_top-row_i*rel_h; ryb=ry-rel_h; rw1=210; dtw=90
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

    disc_y=rel_top-3*rel_h-10
    T(c,LM+5,disc_y,"Submitting a sample via this chain of custody constitutes acknowledgment and acceptance of the KELP\u2019s Terms and Conditions",fs=FS_LABEL)
    form_bot=disc_y-4; c.setStrokeColor(black); c.setLineWidth(LW_OUTER); c.rect(LM,form_bot,RM-LM,hdr_top-form_bot,fill=0,stroke=1)
    _footer(c, 1, total_pages)

    # ═══════════════════════════════════════════════════════════
    # PAGE 2: ANALYSIS DETAIL
    # ═══════════════════════════════════════════════════════════
    c.showPage()
    c.setFont("Helvetica-Bold",14); c.setFillColor(KELP_BLUE)
    c.drawCentredString(PW/2, PH-40, "Analysis Detail \u2014 Selected Analytes per Sample")
    c.setFont("Helvetica",9); c.setFillColor(black)
    c.drawCentredString(PW/2, PH-56, "Detailed analyte breakdown for each sample on the COC. (Category, Method, Analytes)")

    y_pos=PH-80; line_h=13
    samples_with_analyses=[s for s in samples if s.get("analyses")]

    if not samples_with_analyses:
        c.setFont("Helvetica-Oblique",10); c.setFillColor(black)
        c.drawString(LM+10,y_pos,"No analyses selected for any samples.")
    else:
        for si,s in enumerate(samples_with_analyses):
            sid=s.get("sample_id","") or f"Sample {si+1}"
            s_analyses=s.get("analyses",{})
            if isinstance(s_analyses,list): s_analyses={cat:[] for cat in s_analyses}

            bar_h=16
            if y_pos-bar_h<40: c.showPage(); y_pos=PH-40; _footer(c,2,total_pages)
            c.setFillColor(KELP_BLUE); c.rect(LM,y_pos-bar_h,RM-LM,bar_h,fill=1,stroke=0)
            c.setFont("Helvetica-Bold",9); c.setFillColor(white)
            c.drawString(LM+6,y_pos-bar_h+4,f"Sample: {sid}")
            matrix=s.get("matrix","")
            if matrix: c.setFont("Helvetica",8); c.setFillColor(white); c.drawString(LM+250,y_pos-bar_h+4,f"Matrix: {matrix}")
            y_pos-=bar_h+4

            for cat_name in KELP_ANALYTE_CATALOG.keys():
                if cat_name not in s_analyses or not s_analyses[cat_name]: continue
                analytes=s_analyses[cat_name]
                methods_str=", ".join(KELP_ANALYTE_CATALOG[cat_name]["methods"])

                if y_pos-line_h<40: c.showPage(); y_pos=PH-40
                c.setFillColor(SECTION_BG); c.rect(LM+10,y_pos-line_h,RM-LM-20,line_h,fill=1,stroke=0)
                c.setFont("Helvetica-Bold",8); c.setFillColor(HDR_BLUE)
                c.drawString(LM+14,y_pos-line_h+3,f"{cat_name}  \u2014  {methods_str}")
                y_pos-=line_h+2

                analyte_text=", ".join(analytes)
                c.setFont("Helvetica",8); c.setFillColor(black)
                max_line_w=RM-LM-50; words=analyte_text.split(", "); line=""
                for w in words:
                    test=line+(", " if line else "")+w
                    if stringWidth(test,"Helvetica",8)<=max_line_w: line=test
                    else:
                        if y_pos-line_h<40: c.showPage(); y_pos=PH-40
                        c.drawString(LM+24,y_pos-10,line); y_pos-=line_h; line=w
                if line:
                    if y_pos-line_h<40: c.showPage(); y_pos=PH-40
                    c.drawString(LM+24,y_pos-10,line); y_pos-=line_h
                y_pos-=2
            y_pos-=8

    _footer(c, 2, total_pages)

    # ═══════════════════════════════════════════════════════════
    # PAGE 3: INSTRUCTIONS
    # ═══════════════════════════════════════════════════════════
    c.showPage()
    c.setFont("Helvetica-Bold",14); c.setFillColor(black)
    c.drawCentredString(PW/2,PH-40,"Chain of Custody (COC) Instructions")
    c.setFont("Helvetica",10); c.setFillColor(black)
    c.drawCentredString(PW/2,PH-56,"Complete all relevant fields on the COC form. Incomplete information may cause delays.")

    col1_x=30; col2_x=420; col_w=370; bfs=9.5; hfs=10; bul_=9

    def sec_heading(yp,text,col_x=col1_x):
        c.setFont("Helvetica-Bold",hfs); c.setFillColor(HDR_BLUE); c.drawString(col_x,yp,text); return yp-14

    def bullet_item(yp,label,desc,col_x=col1_x,w=col_w):
        c.setFont("Helvetica-Bold",bfs); c.setFillColor(black); c.drawString(col_x+bul_,yp,"\u2022")
        lbl_w=stringWidth(label,"Helvetica-Bold",bfs); c.drawString(col_x+bul_+8,yp,label)
        c.setFont("Helvetica",bfs); remaining=desc; first=True
        x_start=col_x+bul_+8+lbl_w+3; avail_w=w-bul_-8-(lbl_w+3)
        while remaining:
            words=remaining.split(); line=""; rest_words=[]
            for i,word in enumerate(words):
                test=line+(" " if line else "")+word
                if stringWidth(test,"Helvetica",bfs)<=avail_w: line=test
                else: rest_words=words[i:]; break
            else: rest_words=[]
            c.drawString(x_start,yp,line); remaining=" ".join(rest_words)
            if remaining: yp-=12; x_start=col_x+bul_+8; avail_w=w-bul_-8; first=False
        return yp-14

    y=PH-85; y=sec_heading(y,"1. Client & Project Information:")
    for lbl,desc in [("Company Name:","Your company\u2019s name."),("Street Address:","Your mailing address."),("Contact/Report To:","Person designated to receive results."),("Customer Project # and Project Name:","Your project reference number and name."),("Site Collection Info/Facility ID:","Project location or facility ID."),("Time Zone:","Sample collection time zone (e.g., AK, PT, MT, CT, ET)."),("Purchase Order #:","Your PO number for invoicing."),("Invoice To:","Contact person for the invoice."),("Invoice Email:","Email address for the invoice."),("Phone #:","Your contact phone number."),("E-mail:","Your email for correspondence and the final report."),("Data Deliverable:","Required data deliverable level."),("Field Filtered:","Indicate if samples were filtered in the field (Yes/No)."),("Quote #:","Quote number, if applicable."),("DW PWSID # or WW Permit #:","Relevant permit numbers, if applicable.")]:
        y=bullet_item(y,lbl,desc)

    y-=4; y=sec_heading(y,"2. Sample Information:")
    for lbl,desc in [("Customer Sample ID:","Unique sample identifier for the report."),("Collected Date:","Sample collection date."),("Collected Time:","Sample collection time."),("Comp/Grab:","GRAB for single-point; COMP for combined samples."),("Matrix:","Sample type (DW, GW, WW, P, SW, OT)."),("Container Size:","Specify size (e.g., 1L, 500mL).")]:
        y=bullet_item(y,lbl,desc)

    y2=PH-85
    for lbl,desc in [("Container Preservative Type:","Specify preservative (None, HNO3, H2SO4, etc.)."),("Analysis Requested:","Select individual analytes by category. Detailed breakdown on Page 2."),("Sample Comment:","Notes about individual samples; identify MS/MSD samples here."),("Residual Chlorine:","Record results and units if measured.")]:
        y2=bullet_item(y2,lbl,desc,col_x=col2_x,w=col_w)

    y2-=4; y2=sec_heading(y2,"3. Additional Information & Instructions:",col_x=col2_x)
    for lbl,desc in [("Customer Remarks/Hazards:","Note special instructions or hazards."),("Rush Request:","For expedited results; pre-approval required."),("Relinquished By/Received By:","Sign and date at each transfer of custody.")]:
        y2=bullet_item(y2,lbl,desc,col_x=col2_x,w=col_w)

    y2-=4; y2=sec_heading(y2,"4. Sample Acceptance Policy Summary:",col_x=col2_x)
    c.setFont("Helvetica",bfs); c.setFillColor(black); c.drawString(col2_x,y2,"For samples to be accepted, ensure:"); y2-=14
    for item in ["Complete COC documentation.","Readable, unique sample ID on containers (indelible ink).","Appropriate containers and sufficient volume.","Receipt within holding time and temperature requirements.","Containers are in good condition, seals intact (if used).","Proper preservation, no headspace in volatile water samples.","Adequate volume for MS/MSD if required."]:
        c.setFont("Helvetica",bfs); c.drawString(col2_x+bul_,y2,"\u2022  "+item); y2-=13

    y2-=6; c.setFont("Helvetica",8.5); c.setFillColor(black)
    closing="Failure to meet these may result in data qualifiers. A detailed policy is available from your Project Manager. Submitting samples implies acceptance of KELP Terms and Conditions."
    words=closing.split(); line=""; cy=y2
    for word in words:
        test=line+(" " if line else "")+word
        if stringWidth(test,"Helvetica",8.5)<=col_w: line=test
        else: c.drawString(col2_x,cy,line); cy-=12; line=word
    if line: c.drawString(col2_x,cy,line)

    _footer(c, 3, total_pages)
    c.showPage(); c.save(); buf.seek(0)
    return buf

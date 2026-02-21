"""
coc_pdf_engine.py - KELP COC PDF Generator v18

Fixes: vertical text spacing, footer overlap, print-friendly font sizes
"""
import io, os, textwrap
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

PW, PH = landscape(letter)
KELP_BLUE = HexColor("#1F4E79")
HDR_BLUE = HexColor("#1F4E79")
SECTION_BG = HexColor("#E8ECF0")
ROW_SHADE = HexColor("#F0F4F8")
LW_OUTER = 1.0; LW_SECTION = 0.5; LW_INNER = 0.3
DOC_ID = "KELP-QMS-FORM-001"; DOC_VERSION = "1.1"; DOC_EFF_DATE = "February 19, 2026"
FS_TITLE = 11; FS_LABEL = 6.5; FS_VALUE = 8; FS_HEADER = 6.5
FS_LEGEND = 5.5; FS_FOOTER = 5.5; FS_VERT = 5.5
LM = 18; RM = 774; COL2 = 237; ACOL = 476

from coc_catalog import KELP_ANALYTE_CATALOG, CAT_SHORT_MAP, generate_coc_id

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
    c.setFont("Helvetica-Bold",6); c.setFillColor(HDR_BLUE)
    c.drawCentredString(x+w/2,y+h/2-2.5,text)

def _footer(c,pn,tp,coc_id=""):
    y=18
    c.setStrokeColor(black); c.setLineWidth(0.3); c.line(LM,y,RM,y)
    c.setFont("Helvetica",FS_FOOTER); c.setFillColor(black)
    lt=DOC_ID+"  |  Version "+DOC_VERSION+"  |  Effective: "+DOC_EFF_DATE
    if coc_id: lt+="  |  COC ID: "+coc_id
    c.drawString(LM,y-9,lt); c.drawRightString(RM,y-9,"CONTROLLED DOCUMENT  |  Page "+str(pn)+" of "+str(tp))

def _wrap_vtext(label, max_h, fs, bold):
    fn="Helvetica-Bold" if bold else "Helvetica"
    if stringWidth(label,fn,fs)<=max_h-6: return [label]
    words=label.split(); lines=[]; cur=""
    for word in words:
        test=cur+(" " if cur else "")+word
        if stringWidth(test,fn,fs)<=max_h-8: cur=test
        else:
            if cur: lines.append(cur)
            cur=word
    if cur: lines.append(cur)
    return lines

def _build_analysis_columns(samples):
    cat_analytes={}; cat_order=list(KELP_ANALYTE_CATALOG.keys())
    for s in samples:
        analyses=s.get("analyses",{})
        if isinstance(analyses,list): analyses={cat:[] for cat in analyses}
        for cn,al in analyses.items():
            if not al: continue
            if cn not in cat_analytes: cat_analytes[cn]=[]
            for a in al:
                if a not in cat_analytes[cn]: cat_analytes[cn].append(a)
    columns=[]
    MAX_CHARS=100
    for cn in cat_order:
        if cn not in cat_analytes: continue
        analytes=cat_analytes[cn]
        method=", ".join(KELP_ANALYTE_CATALOG[cn]["methods"])
        short=CAT_SHORT_MAP.get(cn,cn)
        a_str=", ".join(analytes)
        full=short+" ("+a_str+")"
        msub="("+method+")"
        if len(full)<=MAX_CHARS:
            columns.append({"label":full,"method":msub,"cat_name":cn})
        else:
            chunks=[]; cur=[]; cur_len=len(short)+3
            for a in analytes:
                add=len(a)+(2 if cur else 0)
                if cur_len+add>MAX_CHARS and cur:
                    chunks.append(cur); cur=[a]; cur_len=len(short)+10+len(a)
                else: cur.append(a); cur_len+=add
            if cur: chunks.append(cur)
            for ci,chunk in enumerate(chunks):
                lbl=short+" ("+", ".join(chunk)+")" if ci==0 else short+" cont'd ("+", ".join(chunk)+")"
                columns.append({"label":lbl,"method":msub,"cat_name":cn})
    return columns

def generate_coc_pdf(data, logo_path=None):
    buf=io.BytesIO(); c=canvas.Canvas(buf,pagesize=(PW,PH))
    c.setTitle("KELP Chain-of-Custody")
    d=data; g=lambda k,dflt="": d.get(k,dflt) or dflt
    coc_id=d.get("coc_id") or generate_coc_id()
    dyn_cols=_build_analysis_columns(d.get("samples",[])); num_acols=max(len(dyn_cols),1)
    total_pages=2

    KELP_STRIP_W=10; COMMENT_W=84; PNC_W=20
    right_fixed=KELP_STRIP_W+COMMENT_W+PNC_W
    atotal_w=RM-ACOL-right_fixed; col_w=max(18,atotal_w/num_acols)
    if col_w*num_acols>atotal_w: col_w=atotal_w/num_acols
    AX=[ACOL+i*col_w for i in range(num_acols+1)]
    KSX=AX[-1]; SBX=KSX+KELP_STRIP_W; PNCX=SBX+COMMENT_W

    # === PAGE 1 ===
    hdr_top=588; hdr_bot=556; hdr_h=hdr_top-hdr_bot
    c.setStrokeColor(KELP_BLUE); c.setLineWidth(1.5)
    c.line(LM,hdr_top,RM,hdr_top); c.line(LM,hdr_bot,RM,hdr_bot)
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER); c.rect(LM,hdr_bot,RM-LM,hdr_h,fill=0,stroke=1)

    if logo_path and os.path.exists(logo_path):
        try: c.drawImage(ImageReader(logo_path),LM+3,hdr_bot+3,width=80,height=hdr_h-6,preserveAspectRatio=True,mask='auto')
        except: pass

    c.setFont("Helvetica-Bold",FS_TITLE); c.setFillColor(KELP_BLUE)
    c.drawCentredString((LM+ACOL)/2,hdr_top-13,"CHAIN-OF-CUSTODY RECORD")
    c.setFont("Helvetica",7); c.setFillColor(black)
    c.drawCentredString((LM+ACOL)/2,hdr_top-24,"Chain-of-Custody is a LEGAL DOCUMENT \u2014 Complete all relevant fields")

    kx=570; R(c,kx,hdr_bot+1,RM-kx-1,hdr_h-2,lw=LW_SECTION)
    T(c,kx+3,hdr_top-10,"KELP USE ONLY",fs=7,bold=True,color=HDR_BLUE)
    T(c,kx+3,hdr_top-20,"KELP Ordering ID:",fs=FS_LABEL)
    HLINE(c,kx+72,RM-6,hdr_top-22,lw=0.3)
    kid=g("kelp_ordering_id");
    if kid: TV(c,kx+74,hdr_top-20,kid)
    T(c,kx+3,hdr_top-29,"COC ID: "+coc_id,fs=6,bold=True,color=HDR_BLUE)

    ci_top=hdr_bot; rh=10; lw_=COL2-LM; cw_=ACOL-COL2; sec_h=8
    SECTION_LABEL(c,LM,ci_top-sec_h,ACOL-LM,sec_h,"CLIENT INFORMATION"); ci_top-=sec_h

    def ci_row(ri,ll,lv,cl,cv):
        yt=ci_top-ri*rh; yb=yt-rh; R(c,LM,yb,lw_,rh); R(c,COL2,yb,cw_,rh)
        if ll: T(c,LM+2,yb+2,ll); tw=stringWidth(ll,"Helvetica",FS_LABEL)+3; TV(c,LM+tw,yb+1,lv,maxw=lw_-tw-2)
        if cl: T(c,COL2+2,yb+2,cl); tw=stringWidth(cl,"Helvetica",FS_LABEL)+3; TV(c,COL2+tw,yb+1,cv,maxw=cw_-tw-2)

    ci_row(0,"Company Name:",g("company_name"),"Contact/Report To:",g("contact_name"))
    ci_row(1,"Street Address:",g("street_address"),"Phone #:",g("phone"))
    ci_row(2,"","","E-Mail:",g("email"))
    ci_row(3,"","","Cc E-Mail:",g("cc_email"))
    ci_row(4,"Customer Project #:",g("project_number"),"Invoice to:",g("invoice_to"))
    R(c,ACOL,ci_top-5*rh,RM-ACOL,5*rh)

    z3_top=ci_top-5*rh; z3_rh=19
    SECTION_LABEL(c,LM,z3_top-sec_h,ACOL-LM,sec_h,"PROJECT DETAILS"); z3_top-=sec_h

    y6b=z3_top-z3_rh
    R(c,LM,y6b,lw_,z3_rh); T(c,LM+2,y6b+10,"Project Name:"); TV(c,LM+55,y6b+10,g("project_name"),maxw=lw_-59)
    R(c,COL2,y6b,cw_,z3_rh); T(c,COL2+2,y6b+10,"Invoice E-mail:"); TV(c,COL2+60,y6b+10,g("invoice_email"),maxw=cw_-64)

    y7b=y6b-z3_rh
    R(c,LM,y7b,lw_,z3_rh); T(c,LM+2,y7b+10,"Site Collection Info/Facility ID (as applicable):"); TV(c,LM+2,y7b+1,g("site_info"),maxw=lw_-6)
    R(c,COL2,y7b,cw_,z3_rh); T(c,COL2+2,y7b+10,"Purchase Order (if applicable):"); TV(c,COL2+2,y7b+1,g("purchase_order"),maxw=cw_-6)

    y8b=y7b-z3_rh; R(c,LM,y8b,lw_,z3_rh); R(c,COL2,y8b,cw_,z3_rh)
    T(c,COL2+2,y8b+10,"Quote #:"); TV(c,COL2+35,y8b+10,g("quote_number"),maxw=cw_-39)

    y9b=y8b-z3_rh
    R(c,LM,y9b,lw_,z3_rh); T(c,LM+2,y9b+10,"County / State origin of sample(s):"); TV(c,LM+2,y9b+1,g("county_state"),maxw=lw_-6)
    R(c,COL2,y9b,cw_,z3_rh)

    acol_w=AX[-1]-AX[0]; lr_x=SBX; lr_w=RM-SBX
    R(c,ACOL,y6b,acol_w+KELP_STRIP_W,z3_rh); R(c,lr_x,y6b,lr_w,z3_rh)
    R(c,ACOL,y7b,acol_w+KELP_STRIP_W,z3_rh)
    cs=g("container_size")
    if cs: T(c,ACOL+2,y7b+10,"Specify Container Size:"); TV(c,ACOL+95,y7b+10,cs,maxw=acol_w-100)
    else: TC(c,ACOL,y7b+6,acol_w,"Specify Container Size",fs=FS_HEADER,bold=True)
    R(c,lr_x,y7b,lr_w,z3_rh)
    T(c,lr_x+2,y7b+12,"Container Size: (1) 1L, (2) 500mL,",fs=5,bold=True)
    T(c,lr_x+2,y7b+6,"(3) 250mL, (4) 125mL, (5) 100mL,",fs=5,bold=True)
    T(c,lr_x+2,y7b+1,"(6) Other",fs=5,bold=True)

    R(c,ACOL,y8b,acol_w+KELP_STRIP_W,z3_rh)
    pv=g("preservative_type")
    if pv: T(c,ACOL+2,y8b+10,"Identify Container Preservative Type:"); TV(c,ACOL+145,y8b+10,pv,maxw=acol_w-150)
    else: TC(c,ACOL,y8b+6,acol_w,"Identify Container Preservative Type",fs=FS_HEADER,bold=True)
    R(c,lr_x,y9b,lr_w,y7b-y9b)
    T(c,lr_x+2,y8b+10,"Preservative: (1) None, (2) HNO3,",fs=FS_LEGEND)
    T(c,lr_x+2,y8b+4,"(3) H2SO4, (4) HCl, (5) NaOH,",fs=FS_LEGEND)
    T(c,lr_x+2,y9b+14,"(6) Zn Acetate, (7) NaHSO4,",fs=FS_LEGEND)
    T(c,lr_x+2,y9b+8,"(8) Sod.Thiosulfate, (9) Ascorbic Acid,",fs=FS_LEGEND)
    T(c,lr_x+2,y9b+2,"(10) MeOH, (11) Other",fs=FS_LEGEND)
    R(c,ACOL,y9b,acol_w+KELP_STRIP_W,z3_rh)
    TC(c,ACOL,y9b+6,acol_w,"Analysis Requested",fs=FS_HEADER,bold=True)

    z4_top=z3_top-sec_h-4*z3_rh; lrh=14.5
    ya_b=z4_top-lrh; R(c,LM,ya_b,ACOL-LM,lrh); T(c,LM+2,ya_b+3,"Sample Collection Time Zone:")
    tz=g("time_zone","PT")
    for tn,tx in [("AK",130),("PT",160),("MT",186),("CT",210),("ET",236)]:
        CB(c,tx,ya_b+2,checked=(tn==tz)); T(c,tx+9,ya_b+3,tn)

    dd_top=ya_b; dd_h=lrh*4; dd_w=130; reg_w=ACOL-LM-dd_w; rx=LM+dd_w
    R(c,LM,dd_top-dd_h,dd_w,dd_h); T(c,LM+2,dd_top-10,"Data Deliverables:")
    sd=g("data_deliverable","Level I (Std)")
    for lbl,bx,by in [("Level I (Std)",LM+6,dd_top-24),("Level II",LM+68,dd_top-24),("Level III",LM+6,dd_top-38),("Other",LM+6,dd_top-52)]:
        CB(c,bx,by,checked=(lbl==sd)); T(c,bx+9,by+1,lbl)

    R(c,rx,dd_top-lrh,reg_w,lrh); T(c,rx+2,dd_top-lrh+3,"Regulatory Program (DW, RCRA, etc.):")
    rpx=rx+reg_w*0.62; T(c,rpx,dd_top-lrh+3,"Reportable"); rv=g("reportable")=="Yes"
    CB(c,rpx+42,dd_top-lrh+2,checked=rv); T(c,rpx+51,dd_top-lrh+3,"Yes")
    CB(c,rpx+68,dd_top-lrh+2,checked=not rv); T(c,rpx+77,dd_top-lrh+3,"No")

    R(c,rx,dd_top-lrh*2,reg_w,lrh); T(c,rx+2,dd_top-lrh*2+3,"Rush (Pre-approval required):",bold=True)
    sr=g("rush","Standard (5-10 Day)")
    for i,ro in enumerate(["Same Day","1 Day","2 Day","3 Day","4 Day"]):
        rrx=rx+125+i*40; CB(c,rrx,dd_top-lrh*2+2,checked=(ro==sr)); T(c,rrx+9,dd_top-lrh*2+3,ro,fs=FS_LEGEND)

    h3=reg_w*0.5; R(c,rx,dd_top-lrh*3,h3,lrh)
    CB(c,rx+4,dd_top-lrh*3+2,checked=("5 Day" in sr)); T(c,rx+14,dd_top-lrh*3+3,"5 Day"); T(c,rx+48,dd_top-lrh*3+3,"Other ____________")
    R(c,rx+h3,dd_top-lrh*3,reg_w-h3,lrh); T(c,rx+h3+2,dd_top-lrh*3+3,"DW PWSID # or WW Permit #:",fs=FS_LEGEND)

    R(c,rx,dd_top-lrh*4,reg_w,lrh); T(c,rx+2,dd_top-lrh*4+3,"Field Filtered (if applicable):")
    ff=g("field_filtered","No")
    CB(c,rx+125,dd_top-lrh*4+2,checked=(ff=="Yes")); T(c,rx+134,dd_top-lrh*4+3,"Yes")
    CB(c,rx+158,dd_top-lrh*4+2,checked=(ff=="No")); T(c,rx+167,dd_top-lrh*4+3,"No")

    ml_top=dd_top-dd_h; ml_h=10; R(c,LM,ml_top-ml_h,ACOL-LM,ml_h)
    T(c,LM+2,ml_top-ml_h+2,"* Matrix: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water(SW), Other(OT)",fs=FS_LEGEND)

    th_top=ml_top-ml_h; th_h=32; grp_h=12; sub_h=th_h-grp_h
    R(c,LM,th_top-th_h,160.8,th_h); TC(c,LM,th_top-th_h/2-3,160.8,"Customer Sample ID",fs=FS_HEADER,bold=True)
    R(c,178.1,th_top-th_h,31.5,th_h); TC(c,178.1,th_top-th_h/2+3,31.5,"Matrix",fs=FS_HEADER,bold=True); TC(c,178.1,th_top-th_h/2-7,31.5,"*",fs=FS_HEADER,bold=True)
    R(c,209.6,th_top-th_h,27.8,th_h); TC(c,209.6,th_top-th_h/2+3,27.8,"Comp /",fs=FS_HEADER,bold=True); TC(c,209.6,th_top-th_h/2-7,27.8,"Grab",fs=FS_HEADER,bold=True)
    R(c,237.4,th_top-grp_h,89.6,grp_h); TC(c,237.4,th_top-grp_h+2,89.6,"Composite Start",fs=FS_LABEL,bold=True)
    R(c,237.4,th_top-th_h,54.3,sub_h); TC(c,237.4,th_top-th_h+sub_h/2-3,54.3,"Date",fs=FS_LABEL,bold=True)
    R(c,291.7,th_top-th_h,35.3,sub_h); TC(c,291.7,th_top-th_h+sub_h/2-3,35.3,"Time",fs=FS_LABEL,bold=True)
    R(c,327.0,th_top-grp_h,85.2,grp_h); TC(c,327.0,th_top-grp_h+2,85.2,"Collected or Composite End",fs=FS_LEGEND,bold=True)
    R(c,327.0,th_top-th_h,50.2,sub_h); TC(c,327.0,th_top-th_h+sub_h/2-3,50.2,"Date",fs=FS_LABEL,bold=True)
    R(c,377.2,th_top-th_h,35.0,sub_h); TC(c,377.2,th_top-th_h+sub_h/2-3,35.0,"Time",fs=FS_LABEL,bold=True)
    R(c,412.2,th_top-th_h,19.9,th_h); TC(c,412.2,th_top-th_h/2+3,19.9,"#",fs=FS_HEADER,bold=True); TC(c,412.2,th_top-th_h/2-7,19.9,"Cont.",fs=FS_HEADER,bold=True)
    R(c,432.1,th_top-grp_h,44.7,grp_h); TC(c,432.1,th_top-grp_h+2,44.7,"Residual Chlorine",fs=FS_LEGEND,bold=True)
    R(c,432.1,th_top-th_h,23.1,sub_h); TC(c,432.1,th_top-th_h+sub_h/2-3,23.1,"Result",fs=FS_LEGEND,bold=True)
    R(c,455.2,th_top-th_h,21.6,sub_h); TC(c,455.2,th_top-th_h+sub_h/2-3,21.6,"Units",fs=FS_LEGEND,bold=True)

    # TALL VERTICAL ANALYSIS COLUMNS
    tall_bot=th_top-th_h; tall_h=z4_top-tall_bot
    cat_col_indices={}
    for ci,col_info in enumerate(dyn_cols):
        cat=col_info["cat_name"]
        if cat not in cat_col_indices: cat_col_indices[cat]=[]
        cat_col_indices[cat].append(ci)

    vert_line_gap=FS_VERT+3  # spacing between wrapped vertical lines

    for ci in range(num_acols):
        ax0=AX[ci]; ax1=AX[ci+1] if ci+1<len(AX) else KSX; aw=ax1-ax0
        R(c,ax0,tall_bot,aw,tall_h)
        if ci<len(dyn_cols):
            col_info=dyn_cols[ci]; label=col_info["label"]; method=col_info["method"]
            avail_h=tall_h-8
            label_lines=_wrap_vtext(label,avail_h,FS_VERT,bold=True)
            nl=len(label_lines)
            # Total width = label lines + gap + method line
            total_needed=(nl+1.5)*vert_line_gap
            vlg=vert_line_gap
            if total_needed>aw-4: vlg=max(FS_VERT+1,(aw-4)/(nl+1.5))
            # Label occupies right portion, method on left
            block_w=nl*vlg
            label_right=ax0+aw/2+block_w/2
            if label_right>ax1-2: label_right=ax1-2
            for li,line in enumerate(label_lines):
                lx=label_right-li*vlg
                VTEXT(c,lx,tall_bot+4,line,fs=FS_VERT,bold=True)
            method_x=label_right-nl*vlg-vlg*0.8
            if method_x<ax0+1: method_x=ax0+1
            VTEXT(c,method_x,tall_bot+4,method,fs=FS_VERT-0.5,bold=False)

    R(c,KSX,tall_bot,KELP_STRIP_W,tall_h)
    VTEXT(c,KSX+KELP_STRIP_W/2+2,tall_bot+tall_h/2-20,"KELP Use Only",fs=FS_LEGEND,bold=True)

    sbf_h=18; fy=z4_top
    for lbl,val in [("Project Mgr.:",g("project_manager")),("AcctNum / Client ID:",g("acct_num")),("Table #:",g("table_number")),("Profile / Template:",g("profile_template")),("Prelog / Bottle Ord. ID:",g("prelog_id"))]:
        R(c,SBX,fy-sbf_h,COMMENT_W,sbf_h)
        T(c,SBX+2,fy-sbf_h+8,lbl); lw2=stringWidth(lbl,"Helvetica",FS_LABEL)+3
        TV(c,SBX+lw2,fy-sbf_h+8,val,fs=7,maxw=COMMENT_W-lw2-2); fy-=sbf_h
    sc_h=fy-tall_bot; R(c,SBX,tall_bot,COMMENT_W,sc_h)
    TC(c,SBX,tall_bot+sc_h/2-3,COMMENT_W,"Sample Comment",fs=FS_HEADER,bold=True)

    R(c,PNCX,tall_bot,PNC_W,z4_top-tall_bot)
    VTEXT(c,PNCX+PNC_W/2+4,tall_bot+6,"Preservation non-conformance",fs=5,bold=False)
    VTEXT(c,PNCX+PNC_W/2-4,tall_bot+6,"identified for sample.",fs=5,bold=False)

    # SAMPLE DATA ROWS
    data_rh=19; max_rows=10; data_top=tall_bot
    SECTION_LABEL(c,LM,data_top-sec_h,ACOL-LM,sec_h,"SAMPLE INFORMATION"); data_top-=sec_h
    samples=d.get("samples",[])
    for ri in range(max_rows):
        s=samples[ri] if ri<len(samples) else {}; ryb=data_top-(ri+1)*data_rh
        if ri%2==1: c.setFillColor(ROW_SHADE); c.rect(LM,ryb,RM-LM,data_rh,fill=1,stroke=0)
        c.setFont("Helvetica",5); c.setFillColor(black); c.drawCentredString(LM+4,ryb+7,str(ri+1))
        for x0,x1,val,align,fs in [(LM,178.1,s.get("sample_id",""),"left",FS_VALUE),(178.1,209.6,s.get("matrix",""),"center",FS_VALUE),(209.6,237.4,s.get("comp_grab",""),"center",6.5),(237.4,291.7,s.get("start_date",""),"center",6.5),(291.7,327.0,s.get("start_time",""),"center",6.5),(327.0,377.2,s.get("end_date","") or s.get("collected_date",""),"center",6.5),(377.2,412.2,s.get("end_time","") or s.get("collected_time",""),"center",6.5),(412.2,432.1,s.get("num_containers",""),"center",FS_VALUE),(432.1,455.2,s.get("res_cl_result",""),"center",6.5),(455.2,ACOL,s.get("res_cl_units",""),"center",FS_LEGEND)]:
            R(c,x0,ryb,x1-x0,data_rh)
            if val:
                if align=="center": TC(c,x0,ryb+5,x1-x0,str(val),fs=fs,bold=True)
                else: TV(c,x0+10,ryb+5,str(val),fs=fs,maxw=x1-x0-14)
        sa=s.get("analyses",{})
        if isinstance(sa,list): sa={cat:[] for cat in sa}
        for ci2 in range(num_acols):
            ax0=AX[ci2]; ax1=AX[ci2+1] if ci2+1<len(AX) else KSX; R(c,ax0,ryb,ax1-ax0,data_rh)
        for cat_name,al in sa.items():
            if not al or cat_name not in cat_col_indices: continue
            ci_idx=cat_col_indices[cat_name][0]
            if ci_idx<num_acols:
                ax0=AX[ci_idx]; ax1=AX[ci_idx+1] if ci_idx+1<len(AX) else KSX
                TC(c,ax0,ryb+5,ax1-ax0,"X",fs=FS_VALUE,bold=True)
        R(c,KSX,ryb,KELP_STRIP_W,data_rh); R(c,SBX,ryb,COMMENT_W,data_rh)
        cmt=s.get("comment","")
        if cmt: TV(c,SBX+2,ryb+5,cmt,fs=FS_LEGEND,maxw=COMMENT_W-4)
        R(c,PNCX,ryb,PNC_W,data_rh)

    # BOTTOM ZONE
    bot_top=data_top-max_rows*data_rh
    SECTION_LABEL(c,LM,bot_top-sec_h,RM-LM,sec_h,"CHAIN OF CUSTODY RECORD / LABORATORY RECEIVING"); bot_top-=sec_h
    inst_h=20; half_w=(RM-LM)/2
    R(c,LM,bot_top-inst_h,half_w,inst_h)
    T(c,LM+2,bot_top-7,"Additional Instructions from KELP:",bold=True)
    TV(c,LM+4,bot_top-16,g("additional_instructions"),fs=6,maxw=half_w-8)
    R(c,LM+half_w,bot_top-inst_h,half_w,inst_h)
    T(c,LM+half_w+2,bot_top-7,"Customer Remarks / Special Conditions / Possible Hazards:",bold=True)
    TV(c,LM+half_w+4,bot_top-16,g("customer_remarks"),fs=6,maxw=half_w-8)

    lr_top=bot_top-inst_h; lr_h=11
    R(c,LM,lr_top-lr_h,RM-LM,lr_h)
    T(c,LM+2,lr_top-8,"# Coolers:"); TV(c,LM+44,lr_top-8,g("num_coolers"),fs=7)
    T(c,LM+72,lr_top-8,"Thermometer ID:"); TV(c,LM+132,lr_top-8,g("thermometer_id"),fs=7)
    T(c,LM+190,lr_top-8,"Temp. (\u00b0C):"); TV(c,LM+228,lr_top-8,g("temperature"),fs=7)
    roi=g("received_on_ice","Yes")
    T(c,LM+310,lr_top-8,"Sample Received on ice:")
    CB(c,LM+400,lr_top-10,checked=(roi=="Yes")); T(c,LM+409,lr_top-8,"Yes")
    CB(c,LM+432,lr_top-10,checked=(roi=="No")); T(c,LM+441,lr_top-8,"No")

    rel_top=lr_top-lr_h; rel_h=10
    for row_i in range(3):
        ryb=rel_top-(row_i+1)*rel_h; rw1=210; dtw=90
        R(c,LM,ryb,rw1,rel_h); T(c,LM+2,ryb+3,"Relinquished by/Company: (Signature)",fs=FS_LEGEND)
        R(c,LM+rw1,ryb,dtw,rel_h); T(c,LM+rw1+2,ryb+3,"Date/Time:",fs=FS_LEGEND)
        rcx=LM+rw1+dtw; R(c,rcx,ryb,rw1,rel_h); T(c,rcx+2,ryb+3,"Received by/Company: (Signature)",fs=FS_LEGEND)
        R(c,rcx+rw1,ryb,dtw,rel_h); T(c,rcx+rw1+2,ryb+3,"Date/Time:",fs=FS_LEGEND)
        rsx=rcx+rw1+dtw; rsw=RM-rsx; R(c,rsx,ryb,rsw,rel_h)
        if row_i==0: T(c,rsx+2,ryb+3,"Tracking #:",fs=FS_LEGEND); TV(c,rsx+44,ryb+3,g("tracking_number"),fs=6.5,maxw=rsw-48)
        elif row_i==1:
            dm=g("delivery_method"); T(c,rsx+2,ryb+3,"Delivered by:",fs=FS_LEGEND); avail=rsw-55
            for dn,dx in [("In-Person",rsx+52),("FedEx",rsx+52+avail*0.30),("UPS",rsx+52+avail*0.55),("Other",rsx+52+avail*0.75)]:
                CB(c,dx,ryb+1,checked=(dn==dm),sz=6); T(c,dx+8,ryb+3,dn,fs=FS_LEGEND)

    # Outer border closes at bottom of last relinquished row
    form_bot=rel_top-3*rel_h
    if form_bot<22: form_bot=22
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER); c.rect(LM,form_bot,RM-LM,hdr_top-form_bot,fill=0,stroke=1)

    # Disclaimer BELOW outer border, above footer
    disc_y=form_bot-2
    T(c,LM+5,disc_y,"Submitting a sample via this chain of custody constitutes acknowledgment and acceptance of the KELP\u2019s Terms and Conditions",fs=5)
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER); c.rect(LM,form_bot,RM-LM,hdr_top-form_bot,fill=0,stroke=1)
    _footer(c,1,total_pages,coc_id)

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

    _footer(c,2,total_pages,coc_id)
    c.showPage(); c.save(); buf.seek(0)
    return buf, coc_id

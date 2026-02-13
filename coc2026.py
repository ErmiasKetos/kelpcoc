"""
KELP Chain-of-Custody v14
Document ID: KELP-QMS-FORM-001-ChainOfCustody
Version: 1.0

v14 improvements over v13:
  - Teal header bar with white title text
  - Section divider labels (CLIENT INFORMATION, SAMPLE INFORMATION, etc.)
  - Alternating row shading on data rows (very light gray)
  - Row numbering (1-10) in left margin
  - Bold outer borders, thinner inner grid lines
  - Analysis columns show method in parentheses
  - Default analysis column labels with EPA method references

FONT HIERARCHY (fillable form best practices):
  - Title:           Helvetica-Bold  11pt  white on teal
  - Section labels:  Helvetica       6pt   (recedes visually)
  - Table headers:   Helvetica-Bold  6-7pt
  - User values:     Helvetica-Bold  8pt   (stands out as filled data)
  - Legend/ref text:  Helvetica      5pt
  - Footer:          Helvetica      5pt
"""
import io, os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.colors import black, white, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

PW, PH = landscape(letter)  # 792 x 612

KELP_BLUE = HexColor("#1F4E79")   # KELP brand dark blue — headers, accents
LABEL_GRAY = HexColor("#D6DCE4")  # KELP brand light gray — label cells
HDR_BLUE = HexColor("#1F4E79")
ROW_SHADE = HexColor("#F2F2F2")   # very light gray for alternating rows
SECTION_BG = HexColor("#E8ECF0")  # subtle blue-gray tint for section labels

# Line weight hierarchy
LW_OUTER = 1.0    # bold outer borders
LW_SECTION = 0.7  # section dividers
LW_INNER = 0.3    # thin inner grid

# Default analysis columns with EPA method references
DEFAULT_ANALYSIS_COLUMNS = [
    "General Chem\n(SM 2310/4500)",
    "Metals\n(EPA 200.8)",
    "Inorganics/Anions\n(EPA 300.1)",
    "PFAS\n(EPA 537/1633A)",
    "Hardness\n(SM 2340C)",
    "pH/Cond/Turb\n(EPA 150/120)",
    "Nitrate/Nitrite\n(EPA 300.1)",
    "Lead & Copper\n(EPA 200.8)",
    "Coliform\n(SM 9223B)",
    "Other\n(Specify)",
]

# Document ID
DOC_ID = "KELP-QMS-FORM-001"
DOC_TITLE = "Chain-of-Custody"
DOC_VERSION = "1.0"
DOC_EFF_DATE = "February 13, 2026"

# ── Font sizes ──
FS_TITLE = 11       # page title
FS_LABEL = 6        # field labels (static, light)
FS_VALUE = 8        # user-entered values (bold, prominent)
FS_HEADER = 6.5     # table column headers
FS_LEGEND = 5       # legend/reference text
FS_FOOTER = 5       # footer compliance text
FS_SUBTITLE = 7     # subtitle text

# ── Global X boundaries ──
LM = 17.3;  RM = 785.5
COL2 = 237.4
ACOL = 476.8
KELP_STRIP = 665.0
SB = 674.6
PNC = 763.4

AX = [476.8, 494.8, 513.5, 532.2, 551.2, 569.9, 588.7, 607.4, 626.1, 645.3, 664.3]

# ── Helpers ──
def R(c, x, y, w, h, fill=None, lw=LW_INNER):
    if fill:
        c.setFillColor(fill); c.rect(x, y, w, h, fill=1, stroke=0)
    c.setStrokeColor(black); c.setLineWidth(lw)
    c.rect(x, y, w, h, fill=0, stroke=1)

def RBOX(c, x, y, w, h, fill=None):
    """Draw a box with bold outer border."""
    R(c, x, y, w, h, fill=fill, lw=LW_OUTER)

def SECTION_LABEL(c, x, y, w, h, text):
    """Draw a section divider label with teal-tinted background."""
    c.setFillColor(SECTION_BG); c.rect(x, y, w, h, fill=1, stroke=0)
    c.setStrokeColor(black); c.setLineWidth(LW_SECTION)
    c.rect(x, y, w, h, fill=0, stroke=1)
    c.setFont("Helvetica-Bold", 5.5); c.setFillColor(HDR_BLUE)
    c.drawCentredString(x + w/2, y + h/2 - 2, text)

def T(c, x, y, txt, fs=FS_LABEL, bold=False, font="Helvetica", maxw=None, color=black):
    """Draw text. Default = label style (small, regular)."""
    if not txt: return
    fn = "Helvetica-Bold" if bold else font
    s = float(fs)
    t = str(txt)
    if maxw:
        while s > 4 and stringWidth(t, fn, s) > maxw:
            s -= 0.3
    c.setFont(fn, s); c.setFillColor(color)
    c.drawString(x, y, t)

def TV(c, x, y, txt, fs=FS_VALUE, maxw=None):
    """Draw user VALUE text — bold, larger, stands out."""
    T(c, x, y, txt, fs=fs, bold=True, maxw=maxw)

def TC(c, x, y, w, txt, fs=FS_HEADER, bold=False, color=black):
    if not txt: return
    fn = "Helvetica-Bold" if bold else "Helvetica"
    c.setFont(fn, fs); c.setFillColor(color)
    c.drawCentredString(x + w/2, y, str(txt))

def CB(c, x, y, checked=False, sz=7):
    c.setStrokeColor(black); c.setLineWidth(0.3)
    c.rect(x, y, sz, sz, fill=0, stroke=1)
    if checked:
        c.setFont("Helvetica-Bold", sz); c.setFillColor(black)
        c.drawCentredString(x + sz/2, y + 0.5, "X")

def VTEXT(c, x, y, txt, fs=FS_HEADER, bold=False):
    fn = "Helvetica-Bold" if bold else "Helvetica"
    c.saveState()
    c.setFont(fn, fs); c.setFillColor(black)
    c.translate(x, y); c.rotate(90)
    c.drawString(0, 0, str(txt))
    c.restoreState()

def HLINE(c, x1, x2, y, lw=0.4):
    c.setStrokeColor(black); c.setLineWidth(lw); c.line(x1, y, x2, y)


def generate_coc_pdf(data, logo_path=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PW, PH))
    c.setTitle("KELP Chain-of-Custody")
    d = data
    g = lambda k, dflt="": d.get(k, dflt) or dflt

    # ═══════════════════════════════════════════════════════════
    # ZONE 1: HEADER  (RL y = 560..600)
    # ═══════════════════════════════════════════════════════════
    hdr_top = 598; hdr_bot = 560; hdr_h = hdr_top - hdr_bot

    # Dark blue header bar — KELP brand
    c.setFillColor(KELP_BLUE)
    c.rect(LM, hdr_bot, RM - LM, hdr_h, fill=1, stroke=0)

    # Bold outer border
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER)
    c.rect(LM, hdr_bot, RM - LM, hdr_h, fill=0, stroke=1)

    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(ImageReader(logo_path), LM + 4, hdr_bot + 4,
                        width=90, height=hdr_h - 8,
                        preserveAspectRatio=True, mask='auto')
        except: pass

    # White title text on teal
    c.setFont("Helvetica-Bold", FS_TITLE); c.setFillColor(white)
    c.drawCentredString((LM + ACOL)/2, hdr_top - 16, "CHAIN-OF-CUSTODY RECORD")
    c.setFont("Helvetica", FS_SUBTITLE); c.setFillColor(white)
    c.drawCentredString((LM + ACOL)/2, hdr_top - 28,
       "Chain-of-Custody is a LEGAL DOCUMENT \u2014 Complete all relevant fields")

    # KELP USE ONLY box — white background inset on right
    kelp_x = 570
    c.setFillColor(white)
    c.rect(kelp_x, hdr_bot + 2, RM - kelp_x - 2, hdr_h - 4, fill=1, stroke=0)
    R(c, kelp_x, hdr_bot + 2, RM - kelp_x - 2, hdr_h - 4, lw=LW_SECTION)
    T(c, kelp_x + 3, hdr_top - 12, "KELP USE ONLY", fs=6.5, bold=True, color=HDR_BLUE)
    T(c, kelp_x + 3, hdr_top - 24, "KELP Ordering ID:", fs=FS_LABEL)
    HLINE(c, kelp_x + 68, RM - 6, hdr_top - 26, lw=0.3)
    kid = g("kelp_ordering_id")
    if kid: TV(c, kelp_x + 70, hdr_top - 24, kid)

    # ═══════════════════════════════════════════════════════════
    # ZONE 2: CLIENT INFO ROWS  (5 rows × 10pt)
    # ═══════════════════════════════════════════════════════════
    ci_top = hdr_bot; rh = 10
    lw_ = COL2 - LM; cw_ = ACOL - COL2

    # Section label bar
    sec_h = 8
    SECTION_LABEL(c, LM, ci_top - sec_h, ACOL - LM, sec_h, "CLIENT INFORMATION")
    ci_top -= sec_h

    def ci_row(row_idx, l_label, l_val, c_label, c_val):
        yt = ci_top - row_idx * rh; yb = yt - rh
        R(c, LM, yb, lw_, rh)
        if l_label:
            T(c, LM + 2, yb + 2, l_label, fs=FS_LABEL)
            tw = stringWidth(l_label, "Helvetica", FS_LABEL) + 3
            TV(c, LM + tw, yb + 1, l_val, maxw=lw_ - tw - 2)
        R(c, COL2, yb, cw_, rh)
        if c_label:
            T(c, COL2 + 2, yb + 2, c_label, fs=FS_LABEL)
            tw = stringWidth(c_label, "Helvetica", FS_LABEL) + 3
            TV(c, COL2 + tw, yb + 1, c_val, maxw=cw_ - tw - 2)

    ci_row(0, "Company Name:", g("company_name"), "Contact/Report To:", g("contact_name"))
    ci_row(1, "Street Address:", g("street_address"), "Phone #:", g("phone"))
    ci_row(2, "", "", "E-Mail:", g("email"))
    ci_row(3, "", "", "Cc E-Mail:", g("cc_email"))
    ci_row(4, "Customer Project #:", g("project_number"), "Invoice to:", g("invoice_to"))

    rw_ = RM - ACOL
    R(c, ACOL, ci_top - 5 * rh, rw_, 5 * rh)

    # ═══════════════════════════════════════════════════════════
    # ZONE 3: SPLIT ZONE — Client LEFT + 3 Label Rows RIGHT
    # ═══════════════════════════════════════════════════════════
    z3_top = ci_top - 5 * rh; z3_rh = 19

    # Section label
    SECTION_LABEL(c, LM, z3_top - sec_h, ACOL - LM, sec_h, "PROJECT DETAILS")
    z3_top -= sec_h
    z3_bot = z3_top - 4 * z3_rh

    # --- LEFT: 4 rows of client info ---
    y6t = z3_top; y6b = y6t - z3_rh

    # Row 6: Project Name | Invoice E-mail
    R(c, LM, y6b, lw_, z3_rh)
    T(c, LM + 2, y6b + 10, "Project Name:", fs=FS_LABEL)
    TV(c, LM + 55, y6b + 10, g("project_name"), maxw=lw_ - 59)
    R(c, COL2, y6b, cw_, z3_rh)
    T(c, COL2 + 2, y6b + 10, "Invoice E-mail:", fs=FS_LABEL)
    TV(c, COL2 + 60, y6b + 10, g("invoice_email"), maxw=cw_ - 64)

    # Row 7: Site Collection | Purchase Order
    y7t = y6b; y7b = y7t - z3_rh
    R(c, LM, y7b, lw_, z3_rh)
    T(c, LM + 2, y7b + 10, "Site Collection Info/Facility ID (as applicable):", fs=FS_LABEL)
    TV(c, LM + 2, y7b + 1, g("site_info"), maxw=lw_ - 6)
    R(c, COL2, y7b, cw_, z3_rh)
    T(c, COL2 + 2, y7b + 10, "Purchase Order (if applicable):", fs=FS_LABEL)
    TV(c, COL2 + 2, y7b + 1, g("purchase_order"), maxw=cw_ - 6)

    # Row 8: (blank) | Quote #
    y8t = y7b; y8b = y8t - z3_rh
    R(c, LM, y8b, lw_, z3_rh)
    R(c, COL2, y8b, cw_, z3_rh)
    T(c, COL2 + 2, y8b + 10, "Quote #:", fs=FS_LABEL)
    TV(c, COL2 + 35, y8b + 10, g("quote_number"), maxw=cw_ - 39)

    # Row 9: County/State
    y9t = y8b; y9b = y9t - z3_rh
    R(c, LM, y9b, lw_, z3_rh)
    T(c, LM + 2, y9b + 10, "County / State origin of sample(s):", fs=FS_LABEL)
    TV(c, LM + 2, y9b + 1, g("county_state"), maxw=lw_ - 6)
    R(c, COL2, y9b, cw_, z3_rh)

    # --- RIGHT: Analysis columns area ---
    acol_w = AX[10] - AX[0]
    legend_x = KELP_STRIP; legend_w = RM - legend_x

    # Transition row (alongside Row 6)
    R(c, ACOL, y6b, KELP_STRIP - ACOL, z3_rh)
    R(c, legend_x, y6b, legend_w, z3_rh)

    # Row 1: Specify Container Size — merged write-in
    R(c, AX[0], y7b, acol_w, z3_rh)
    cs_val = g("container_size")
    if cs_val:
        T(c, AX[0] + 2, y7b + 10, "Specify Container Size:", fs=FS_LABEL)
        TV(c, AX[0] + 95, y7b + 10, cs_val, maxw=acol_w - 100)
    else:
        TC(c, AX[0], y7b + 6, acol_w, "Specify Container Size", fs=FS_HEADER, bold=True)

    # Container Size legend
    R(c, legend_x, y7b, legend_w, z3_rh)
    T(c, legend_x + 2, y7b + 10, "Container Size: (1) 1L, (2) 500mL,", fs=FS_LEGEND, bold=True)
    T(c, legend_x + 2, y7b + 3, "(3) 250mL, (4) 125mL, (5) 100mL, (6) Other", fs=FS_LEGEND, bold=True)

    # Row 2: Identify Container Preservative Type — merged write-in
    R(c, AX[0], y8b, acol_w, z3_rh)
    pres_val = g("preservative_type")
    if pres_val:
        T(c, AX[0] + 2, y8b + 10, "Identify Container Preservative Type:", fs=FS_LABEL)
        TV(c, AX[0] + 145, y8b + 10, pres_val, maxw=acol_w - 150)
    else:
        TC(c, AX[0], y8b + 6, acol_w, "Identify Container Preservative Type", fs=FS_HEADER, bold=True)

    # Preservative legend — ONE merged cell spanning rows 2 and 3 (y8b to y9b)
    R(c, legend_x, y9b, legend_w, y7b - y9b)  # spans from y9b to y7b (2 rows)
    T(c, legend_x + 2, y8b + 10, "Preservative: (1) None, (2) HNO3,", fs=4.5)
    T(c, legend_x + 2, y8b + 4, "(3) H2SO4, (4) HCl, (5) NaOH,", fs=4.5)
    T(c, legend_x + 2, y9b + 14, "(6) Zn Acetate, (7) NaHSO4,", fs=4.5)
    T(c, legend_x + 2, y9b + 8, "(8) Sod.Thiosulfate, (9) Ascorbic Acid,", fs=4.5)
    T(c, legend_x + 2, y9b + 2, "(10) MeOH, (11) Other", fs=4.5)

    # Row 3: Analysis Requested — static header label
    R(c, AX[0], y9b, acol_w, z3_rh)
    TC(c, AX[0], y9b + 6, acol_w, "Analysis Requested", fs=FS_HEADER, bold=True)

    # ═══════════════════════════════════════════════════════════
    # ZONE 4: TimeZone+Deliverables LEFT | Tall Headers+Sidebar RIGHT
    # ═══════════════════════════════════════════════════════════
    z4_top = z3_bot
    lrh = 14.5  # left-side row height

    # Time Zone
    ya_t = z4_top; ya_b = ya_t - lrh
    R(c, LM, ya_b, ACOL - LM, lrh)
    T(c, LM + 2, ya_b + 3, "Sample Collection Time Zone:", fs=FS_LABEL)
    tz = g("time_zone", "PT")
    for tz_name, tx in [("AK", 130), ("PT", 160), ("MT", 186), ("CT", 210), ("ET", 236)]:
        CB(c, tx, ya_b + 2, checked=(tz_name == tz))
        T(c, tx + 9, ya_b + 3, tz_name, fs=FS_LABEL)

    # Data Deliverables block (5 rows)
    dd_top = ya_b; dd_h = lrh * 5; dd_w = 130
    reg_w = ACOL - LM - dd_w; rx = LM + dd_w

    R(c, LM, dd_top - dd_h, dd_w, dd_h)
    T(c, LM + 2, dd_top - 10, "Data Deliverables:", fs=FS_LABEL)
    sd = g("data_deliverable", "Level I (Std)")
    for lbl, bx, by in [
        ("Level I (Std)", LM + 6, dd_top - 24),
        ("Level II", LM + 68, dd_top - 24),
        ("Level III", LM + 6, dd_top - 38),
        ("Level IV", LM + 6, dd_top - 52),
        ("Other", LM + 6, dd_top - 66),
    ]:
        CB(c, bx, by, checked=(lbl == sd))
        T(c, bx + 9, by + 1, lbl, fs=FS_LABEL)

    # DD Row 1: Regulatory + Reportable
    R(c, rx, dd_top - lrh, reg_w, lrh)
    T(c, rx + 2, dd_top - lrh + 3, "Regulatory Program (DW, RCRA, etc.):", fs=FS_LABEL)
    TV(c, rx + 145, dd_top - lrh + 3, g("regulatory_program"), maxw=50)
    rpx = rx + reg_w * 0.65
    T(c, rpx, dd_top - lrh + 3, "Reportable", fs=FS_LABEL)
    rv = g("reportable") == "Yes"
    CB(c, rpx + 38, dd_top - lrh + 2, checked=rv); T(c, rpx + 47, dd_top - lrh + 3, "Yes", fs=FS_LABEL)
    CB(c, rpx + 65, dd_top - lrh + 2, checked=not rv); T(c, rpx + 74, dd_top - lrh + 3, "No", fs=FS_LABEL)

    # DD Row 2: Rush
    R(c, rx, dd_top - lrh*2, reg_w, lrh)
    T(c, rx + 2, dd_top - lrh*2 + 3, "Rush (Pre-approval required):", fs=FS_LABEL, bold=True)
    sr = g("rush", "Standard (5-10 Day)")
    for i, ro in enumerate(["Same Day", "1 Day", "2 Day", "3 Day", "4 Day"]):
        rrx = rx + 120 + i * 40
        CB(c, rrx, dd_top - lrh*2 + 2, checked=(ro == sr))
        T(c, rrx + 9, dd_top - lrh*2 + 3, ro, fs=FS_LEGEND)

    # DD Row 3: 5-Day + PWSID
    h3 = reg_w * 0.5
    R(c, rx, dd_top - lrh*3, h3, lrh)
    CB(c, rx + 4, dd_top - lrh*3 + 2, checked=("5 Day" in sr))
    T(c, rx + 14, dd_top - lrh*3 + 3, "5 Day", fs=FS_LABEL)
    T(c, rx + 48, dd_top - lrh*3 + 3, "Other ____________", fs=FS_LABEL)
    R(c, rx + h3, dd_top - lrh*3, reg_w - h3, lrh)
    T(c, rx + h3 + 2, dd_top - lrh*3 + 3, "DW PWSID # or WW Permit #:", fs=FS_LEGEND)
    TV(c, rx + h3 + 108, dd_top - lrh*3 + 3, g("pwsid"), fs=7, maxw=reg_w - h3 - 113)

    # DD Row 4: Field Filtered
    R(c, rx, dd_top - lrh*4, reg_w, lrh)
    T(c, rx + 2, dd_top - lrh*4 + 3, "Field Filtered (if applicable):", fs=FS_LABEL)
    ff = g("field_filtered", "No")
    CB(c, rx + 120, dd_top - lrh*4 + 2, checked=(ff == "Yes")); T(c, rx + 129, dd_top - lrh*4 + 3, "Yes", fs=FS_LABEL)
    CB(c, rx + 150, dd_top - lrh*4 + 2, checked=(ff == "No")); T(c, rx + 159, dd_top - lrh*4 + 3, "No", fs=FS_LABEL)

    # DD Row 5: Analysis profile
    R(c, rx, dd_top - lrh*5, reg_w, lrh)
    T(c, rx + 2, dd_top - lrh*5 + 3, "Analysis:", fs=FS_LABEL)
    TV(c, rx + 36, dd_top - lrh*5 + 3, g("analysis_profile_template"), maxw=reg_w - 40)

    # Matrix Legend
    ml_top = dd_top - dd_h; ml_h = 10
    R(c, LM, ml_top - ml_h, ACOL - LM, ml_h)
    T(c, LM + 2, ml_top - ml_h + 2,
      "* Matrix: Drinking Water(DW), Ground Water(GW), Wastewater(WW), Product(P), Surface Water(SW), Other(OT)",
      fs=FS_LEGEND)

    # Table Column Headers
    th_top = ml_top - ml_h; th_h = 32; grp_h = 12

    for x0, x1, lbl in [
        (17.3, 178.1, "Customer Sample ID"),
        (178.1, 209.6, "Matrix\n*"),
        (209.6, 237.4, "Comp /\nGrab"),
        (412.2, 432.1, "#\nCont."),
    ]:
        R(c, x0, th_top - th_h, x1 - x0, th_h)
        lines = lbl.split("\n")
        if len(lines) == 1:
            TC(c, x0, th_top - th_h/2 - 3, x1 - x0, lines[0], fs=FS_HEADER, bold=True)
        else:
            TC(c, x0, th_top - th_h/2 + 3, x1 - x0, lines[0], fs=FS_HEADER, bold=True)
            TC(c, x0, th_top - th_h/2 - 7, x1 - x0, lines[1], fs=FS_HEADER, bold=True)

    sub_h = th_h - grp_h
    # Composite Start
    R(c, 237.4, th_top - grp_h, 327.0 - 237.4, grp_h)
    TC(c, 237.4, th_top - grp_h + 2, 327.0 - 237.4, "Composite Start", fs=FS_LABEL, bold=True)
    R(c, 237.4, th_top - th_h, 291.7 - 237.4, sub_h)
    TC(c, 237.4, th_top - th_h + sub_h/2 - 3, 291.7 - 237.4, "Date", fs=FS_LABEL, bold=True)
    R(c, 291.7, th_top - th_h, 327.0 - 291.7, sub_h)
    TC(c, 291.7, th_top - th_h + sub_h/2 - 3, 327.0 - 291.7, "Time", fs=FS_LABEL, bold=True)

    # Collected or Composite End
    R(c, 327.0, th_top - grp_h, 412.2 - 327.0, grp_h)
    TC(c, 327.0, th_top - grp_h + 2, 412.2 - 327.0, "Collected or Composite End", fs=FS_LEGEND, bold=True)
    R(c, 327.0, th_top - th_h, 377.2 - 327.0, sub_h)
    TC(c, 327.0, th_top - th_h + sub_h/2 - 3, 377.2 - 327.0, "Date", fs=FS_LABEL, bold=True)
    R(c, 377.2, th_top - th_h, 412.2 - 377.2, sub_h)
    TC(c, 377.2, th_top - th_h + sub_h/2 - 3, 412.2 - 377.2, "Time", fs=FS_LEGEND, bold=True)

    # Residual Chlorine
    R(c, 432.1, th_top - grp_h, 476.8 - 432.1, grp_h)
    TC(c, 432.1, th_top - grp_h + 2, 476.8 - 432.1, "Residual Chlorine", fs=FS_LEGEND, bold=True)
    R(c, 432.1, th_top - th_h, 455.2 - 432.1, sub_h)
    TC(c, 432.1, th_top - th_h + sub_h/2 - 3, 455.2 - 432.1, "Result", fs=FS_LEGEND, bold=True)
    R(c, 455.2, th_top - th_h, 476.8 - 455.2, sub_h)
    TC(c, 455.2, th_top - th_h + sub_h/2 - 3, 476.8 - 455.2, "Units", fs=FS_LEGEND, bold=True)

    # --- RIGHT: Tall Vertical Analysis Headers ---
    tall_bot = th_top - th_h; tall_h = z4_top - tall_bot
    acols = d.get("analysis_columns", DEFAULT_ANALYSIS_COLUMNS)

    for ai in range(10):
        ax0 = AX[ai]; ax1 = AX[ai + 1]; aw = ax1 - ax0
        R(c, ax0, tall_bot, aw, tall_h)
        label = acols[ai] if ai < len(acols) else ""
        if label:
            # Split multi-line labels (e.g., "Metals\n(EPA 200.8)")
            lines = label.split("\n")
            if len(lines) >= 2:
                # Line 1: test name (bold)
                VTEXT(c, ax0 + aw/2 + 4, tall_bot + 3, lines[0], fs=FS_HEADER, bold=True)
                # Line 2: method ref (regular, smaller)
                VTEXT(c, ax0 + aw/2 - 4, tall_bot + 3, lines[1], fs=FS_LEGEND, bold=False)
            else:
                fs = FS_HEADER if len(label) < 20 else FS_LEGEND + 0.5
                VTEXT(c, ax0 + aw/2 + 2, tall_bot + 3, label, fs=fs, bold=True)

    # KELP Use Only strip
    strip_w = SB - KELP_STRIP
    R(c, KELP_STRIP, tall_bot, strip_w, tall_h)
    VTEXT(c, KELP_STRIP + strip_w/2 + 2, tall_bot + tall_h/2 - 25, "KELP Use Only", fs=FS_LEGEND, bold=True)

    # Sidebar fields
    sbf_w = PNC - SB; sbf_h = 18; fy = z4_top
    for lbl, val in [
        ("Project Mgr.:", g("project_manager")),
        ("AcctNum / Client ID:", g("acct_num")),
        ("Table #:", g("table_number")),
        ("Profile / Template:", g("profile_template")),
        ("Prelog / Bottle Ord. ID:", g("prelog_id")),
    ]:
        R(c, SB, fy - sbf_h, sbf_w, sbf_h)
        T(c, SB + 2, fy - sbf_h + 8, lbl, fs=FS_LABEL)
        lw2 = stringWidth(lbl, "Helvetica", FS_LABEL) + 3
        TV(c, SB + lw2, fy - sbf_h + 8, val, fs=7, maxw=sbf_w - lw2 - 2)
        fy -= sbf_h

    sc_h = fy - tall_bot
    R(c, SB, tall_bot, sbf_w, sc_h)
    TC(c, SB, tall_bot + sc_h/2 - 3, sbf_w, "Sample Comment", fs=FS_HEADER, bold=True)

    # Preservation NC column
    data_rh = 19; max_rows = 10
    pnc_bot = tall_bot - max_rows * data_rh
    R(c, PNC, pnc_bot, RM - PNC, z4_top - pnc_bot)
    VTEXT(c, PNC + (RM - PNC)/2 + 2, pnc_bot + 5,
          "Preservation non-conformance identified for sample.", fs=4)

    # ═══════════════════════════════════════════════════════════
    # ZONE 5: SAMPLE DATA ROWS (with alternating shading + row numbers)
    # ═══════════════════════════════════════════════════════════
    data_top = tall_bot
    samples = d.get("samples", [])

    # Add a SAMPLE INFORMATION section label
    SECTION_LABEL(c, LM, data_top - sec_h, ACOL - LM, sec_h, "SAMPLE INFORMATION")
    data_top -= sec_h

    for ri in range(max_rows):
        s = samples[ri] if ri < len(samples) else {}
        ry = data_top - ri * data_rh; ryb = ry - data_rh

        # Alternating row shading
        if ri % 2 == 1:
            c.setFillColor(ROW_SHADE)
            c.rect(LM, ryb, RM - LM, data_rh, fill=1, stroke=0)

        # Row number in small font at left edge
        c.setFont("Helvetica", 4.5); c.setFillColor(black)
        c.drawCentredString(LM + 4, ryb + 7, str(ri + 1))

        for x0, x1, val, align, fs in [
            (17.3, 178.1, s.get("sample_id", ""), "left", FS_VALUE),
            (178.1, 209.6, s.get("matrix", ""), "center", FS_VALUE),
            (209.6, 237.4, s.get("comp_grab", ""), "center", 6),
            (237.4, 291.7, s.get("start_date", ""), "center", 6),
            (291.7, 327.0, s.get("start_time", ""), "center", 6),
            (327.0, 377.2, s.get("end_date", ""), "center", 6),
            (377.2, 412.2, s.get("end_time", ""), "center", 6),
            (412.2, 432.1, s.get("num_containers", ""), "center", FS_VALUE),
            (432.1, 455.2, s.get("res_cl_result", ""), "center", 6),
            (455.2, 476.8, s.get("res_cl_units", ""), "center", FS_LEGEND),
        ]:
            R(c, x0, ryb, x1 - x0, data_rh)
            if val:
                if align == "center":
                    TC(c, x0, ryb + 5, x1 - x0, str(val), fs=fs, bold=True)
                else:
                    TV(c, x0 + 10, ryb + 5, str(val), fs=fs, maxw=x1 - x0 - 14)

        # Analysis X marks — match by first line of column label
        s_analyses = s.get("analyses", [])
        for ai in range(10):
            ax0 = AX[ai]; ax1 = AX[ai + 1]
            R(c, ax0, ryb, ax1 - ax0, data_rh)
            ac_name = acols[ai] if ai < len(acols) else ""
            # Match against first line (before \n) for backward compat
            ac_short = ac_name.split("\n")[0] if ac_name else ""
            if ac_short and ac_short in s_analyses:
                TC(c, ax0, ryb + 5, ax1 - ax0, "X", fs=FS_VALUE, bold=True)

        R(c, KELP_STRIP, ryb, strip_w, data_rh)
        R(c, SB, ryb, sbf_w, data_rh)
        cmt = s.get("comment", "")
        if cmt: TV(c, SB + 2, ryb + 5, cmt, fs=FS_LEGEND, maxw=sbf_w - 4)

    # ═══════════════════════════════════════════════════════════
    # ZONE 6: BOTTOM
    # ═══════════════════════════════════════════════════════════
    bot_top = data_top - max_rows * data_rh

    # Section label
    SECTION_LABEL(c, LM, bot_top - sec_h, RM - LM, sec_h, "CHAIN OF CUSTODY RECORD / LABORATORY RECEIVING")
    bot_top -= sec_h

    # Instructions / Remarks
    inst_h = 28; half_w = (RM - LM) / 2
    R(c, LM, bot_top - inst_h, half_w, inst_h)
    T(c, LM + 2, bot_top - 8, "Additional Instructions from KELP:", fs=FS_LABEL, bold=True)
    TV(c, LM + 4, bot_top - 20, g("additional_instructions"), fs=6, maxw=half_w - 8)
    R(c, LM + half_w, bot_top - inst_h, half_w, inst_h)
    T(c, LM + half_w + 2, bot_top - 8, "Customer Remarks / Special Conditions / Possible Hazards:", fs=FS_LABEL, bold=True)
    TV(c, LM + half_w + 4, bot_top - 20, g("customer_remarks"), fs=6, maxw=half_w - 8)

    # Lab Receiving
    lr_top = bot_top - inst_h; lr_h = 14
    R(c, LM, lr_top - lr_h, RM - LM, lr_h)
    T(c, LM + 2, lr_top - 10, "# Coolers:", fs=FS_LABEL)
    TV(c, LM + 42, lr_top - 10, g("num_coolers"), fs=7)
    T(c, LM + 70, lr_top - 10, "Thermometer ID:", fs=FS_LABEL)
    TV(c, LM + 130, lr_top - 10, g("thermometer_id"), fs=7)
    T(c, LM + 185, lr_top - 10, "Temp. (\u00b0C):", fs=FS_LABEL)
    TV(c, LM + 225, lr_top - 10, g("temperature"), fs=7)
    roi = g("received_on_ice", "Yes")
    T(c, LM + 310, lr_top - 10, "Sample Received on ice:", fs=FS_LABEL)
    CB(c, LM + 400, lr_top - 12, checked=(roi == "Yes")); T(c, LM + 409, lr_top - 10, "Yes", fs=FS_LABEL)
    CB(c, LM + 432, lr_top - 12, checked=(roi == "No")); T(c, LM + 441, lr_top - 10, "No", fs=FS_LABEL)

    # Relinquished/Received (3 rows)
    rel_top = lr_top - lr_h; rel_h = 13
    for row_i in range(3):
        ry = rel_top - row_i * rel_h; ryb = ry - rel_h
        rw1 = 210; dtw = 90
        R(c, LM, ryb, rw1, rel_h)
        T(c, LM + 2, ryb + 3, "Relinquished by/Company: (Signature)", fs=FS_LEGEND)
        R(c, LM + rw1, ryb, dtw, rel_h)
        T(c, LM + rw1 + 2, ryb + 3, "Date/Time:", fs=FS_LEGEND)
        rcx = LM + rw1 + dtw
        R(c, rcx, ryb, rw1, rel_h)
        T(c, rcx + 2, ryb + 3, "Received by/Company: (Signature)", fs=FS_LEGEND)
        R(c, rcx + rw1, ryb, dtw, rel_h)
        T(c, rcx + rw1 + 2, ryb + 3, "Date/Time:", fs=FS_LEGEND)
        rsx = rcx + rw1 + dtw; rsw = RM - rsx
        R(c, rsx, ryb, rsw, rel_h)
        if row_i == 0:
            T(c, rsx + 2, ryb + 3, "Tracking #:", fs=FS_LEGEND)
            TV(c, rsx + 42, ryb + 3, g("tracking_number"), fs=6, maxw=rsw - 46)
        elif row_i == 1:
            dm = g("delivery_method")
            T(c, rsx + 2, ryb + 3, "Delivered by:", fs=FS_LEGEND)
            avail = rsw - 55
            for dn, dx in [("In-Person", rsx+52), ("FedEx", rsx+52+avail*0.30),
                           ("UPS", rsx+52+avail*0.55), ("Other", rsx+52+avail*0.75)]:
                CB(c, dx, ryb + 1, checked=(dn == dm), sz=6)
                T(c, dx + 8, ryb + 3, dn, fs=FS_LEGEND)

    # Disclaimer — with spacing below signature rows
    disc_y = rel_top - 3 * rel_h - 10
    T(c, LM + 5, disc_y,
      "Submitting a sample via this chain of custody constitutes acknowledgment and acceptance of the KELP\u2019s Terms and Conditions",
      fs=FS_LABEL)

    # ═══════════════════════════════════════════════════════════
    # BOLD OUTER BORDER around entire form
    # ═══════════════════════════════════════════════════════════
    form_bot = disc_y - 4
    c.setStrokeColor(black); c.setLineWidth(LW_OUTER)
    c.rect(LM, form_bot, RM - LM, hdr_top - form_bot, fill=0, stroke=1)

    # ═══════════════════════════════════════════════════════════
    # FOOTER — Document ID, Version, Compliance Notice
    # ═══════════════════════════════════════════════════════════
    footer_y = 6
    c.setFont("Helvetica", FS_FOOTER); c.setFillColor(black)
    c.drawString(LM, footer_y,
                 f"{DOC_ID}-{DOC_TITLE.replace(' ', '')}  |  Version {DOC_VERSION}  |  Effective: {DOC_EFF_DATE}")
    c.drawRightString(RM, footer_y,
                      "CONTROLLED DOCUMENT - Do not copy without authorization  |  Page 1 of 2")
    # thin line above footer
    c.setStrokeColor(black); c.setLineWidth(0.3)
    c.line(LM, footer_y + 8, RM, footer_y + 8)

    # ═══════════════════════════════════════════════════════════
    # PAGE 2: COC FILLING INSTRUCTIONS
    # ═══════════════════════════════════════════════════════════
    c.showPage()

    # Title
    c.setFont("Helvetica-Bold", 14); c.setFillColor(black)
    c.drawCentredString(PW/2, PH - 40, "Chain of Custody (COC) Instructions")
    c.setFont("Helvetica", 10); c.setFillColor(black)
    c.drawCentredString(PW/2, PH - 56, "Complete all relevant fields on the COC form. Incomplete information may cause delays.")

    # Two-column layout
    col1_x = 30; col2_x = 420; col_w = 370
    bfs = 9.5   # body font size
    hfs = 10    # heading font size
    bul = 9     # bullet indent

    def sec_heading(yp, text, col_x=col1_x):
        c.setFont("Helvetica-Bold", hfs); c.setFillColor(HDR_BLUE)
        c.drawString(col_x, yp, text)
        return yp - 14

    def bullet(yp, label, desc, col_x=col1_x, w=col_w):
        c.setFont("Helvetica-Bold", bfs); c.setFillColor(black)
        c.drawString(col_x + bul, yp, "\u2022")
        lbl_w = stringWidth(label, "Helvetica-Bold", bfs)
        c.drawString(col_x + bul + 8, yp, label)
        c.setFont("Helvetica", bfs)
        # Wrap description
        remaining = desc
        first = True
        x_start = col_x + bul + 8 + lbl_w + 3 if first else col_x + bul + 8
        avail_w = w - bul - 8 - (lbl_w + 3 if first else 0)
        while remaining:
            # Simple word wrapping
            words = remaining.split()
            line = ""
            rest_words = []
            for i, word in enumerate(words):
                test = line + (" " if line else "") + word
                if stringWidth(test, "Helvetica", bfs) <= avail_w:
                    line = test
                else:
                    rest_words = words[i:]
                    break
            else:
                rest_words = []
            c.drawString(x_start, yp, line)
            remaining = " ".join(rest_words)
            if remaining:
                yp -= 12
                x_start = col_x + bul + 8
                avail_w = w - bul - 8
                first = False
        return yp - 14

    # Column 1: Sections 1 & 2
    y = PH - 85
    y = sec_heading(y, "1. Client & Project Information:")
    items1 = [
        ("Company Name:", "Your company\u2019s name."),
        ("Street Address:", "Your mailing address."),
        ("Contact/Report To:", "Person designated to receive results."),
        ("Customer Project # and Project Name:", "Your project reference number and name."),
        ("Site Collection Info/Facility ID:", "Project location or facility ID."),
        ("Time Zone:", "Sample collection time zone (e.g., AK, PT, MT, CT, ET) for accurate hold times."),
        ("Purchase Order #:", "Your PO number for invoicing, if applicable."),
        ("Invoice To:", "Contact person for the invoice."),
        ("Invoice Email:", "Email address for the invoice."),
        ("Phone #:", "Your contact phone number."),
        ("E-mail:", "Your email for correspondence and the final report."),
        ("Data Deliverable:", "Required data deliverable level (Standard, Level II, III, IV, Other)."),
        ("Field Filtered:", "Indicate if samples were filtered in the field (Yes/No)."),
        ("Quote #:", "Quote number, if applicable."),
        ("DW PWSID # or WW Permit #:", "Relevant drinking water or wastewater permit numbers, if applicable."),
    ]
    for lbl, desc in items1:
        y = bullet(y, lbl, desc)

    y -= 4
    y = sec_heading(y, "2. Sample Information:")
    items2 = [
        ("Customer Sample ID:", "Unique sample identifier for the report."),
        ("Collected Date:", "Sample collection date (provide start and end dates for composites)."),
        ("Collected Time:", "Sample collection time (provide start and end times for composites)."),
        ("Comp/Grab:", "\"GRAB\" for single-point collection; \"COMP\" for combined samples."),
        ("Matrix:", "Sample type (e.g., DW, GW, WW, P, SW, OT)."),
        ("Container Size:", "Specify size (e.g., 1L, 500mL, Other)."),
    ]
    for lbl, desc in items2:
        y = bullet(y, lbl, desc)

    # Column 2: Sections 2 continued, 3, 4
    y2 = PH - 85
    items2b = [
        ("Container Preservative Type:", "Specify preservative (e.g., None, HNO3, H2SO4, Other)."),
        ("Analysis Requested:", "List tests or method numbers and check boxes for applicable samples."),
        ("Sample Comment:", "Notes about individual samples; identify MS/MSD samples here."),
        ("Residual Chlorine:", "Record results and units if measured."),
    ]
    for lbl, desc in items2b:
        y2 = bullet(y2, lbl, desc, col_x=col2_x, w=col_w)

    y2 -= 4
    y2 = sec_heading(y2, "3. Additional Information & Instructions:", col_x=col2_x)
    items3 = [
        ("Customer Remarks/Hazards:", "Note special instructions, potential hazards (attach SDS if possible), or requests for extra report copies."),
        ("Rush Request:", "For expedited results, select an option (Same Day to 5 Day). Pre-approval from the lab is required; surcharges apply."),
        ("Relinquished By/Received By:", "Sign and date at each transfer of custody."),
    ]
    for lbl, desc in items3:
        y2 = bullet(y2, lbl, desc, col_x=col2_x, w=col_w)

    y2 -= 4
    y2 = sec_heading(y2, "4. Sample Acceptance Policy Summary:", col_x=col2_x)
    c.setFont("Helvetica", bfs); c.setFillColor(black)
    c.drawString(col2_x, y2, "For samples to be accepted, ensure:")
    y2 -= 14
    acceptance = [
        "Complete COC documentation.",
        "Readable, unique sample ID on containers (indelible ink).",
        "Appropriate containers and sufficient volume.",
        "Receipt within holding time and temperature requirements.",
        "Containers are in good condition, seals intact (if used).",
        "Proper preservation, no headspace in volatile water samples.",
        "Adequate volume for MS/MSD if required.",
    ]
    for item in acceptance:
        c.setFont("Helvetica", bfs)
        c.drawString(col2_x + bul, y2, "\u2022  " + item)
        y2 -= 13

    y2 -= 6
    c.setFont("Helvetica", 8.5); c.setFillColor(black)
    # Wrap the closing paragraph
    closing = ("Failure to meet these may result in data qualifiers. A detailed policy is "
               "available from your Project Manager. Submitting samples implies acceptance "
               "of KELP Terms and Conditions.")
    words = closing.split()
    line = ""; cy = y2
    for word in words:
        test = line + (" " if line else "") + word
        if stringWidth(test, "Helvetica", 8.5) <= col_w:
            line = test
        else:
            c.drawString(col2_x, cy, line)
            cy -= 12; line = word
    if line:
        c.drawString(col2_x, cy, line)

    # Page 2 footer
    c.setFont("Helvetica", FS_FOOTER); c.setFillColor(black)
    c.setStrokeColor(black); c.setLineWidth(0.3)
    c.line(LM, 14, RM, 14)
    c.drawString(LM, 6,
                 f"{DOC_ID}-{DOC_TITLE.replace(' ', '')}  |  Version {DOC_VERSION}  |  Effective: {DOC_EFF_DATE}")
    c.drawRightString(RM, 6,
                      "CONTROLLED DOCUMENT - Do not copy without authorization  |  Page 2 of 2")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


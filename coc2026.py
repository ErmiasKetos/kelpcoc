"""
coc2026.py — KELP Chain-of-Custody Streamlit App v15
Main entry point for Streamlit Cloud deployment.

Uses:
  - coc_catalog.py for analyte catalog and constants
  - coc_pdf_engine.py for PDF generation
"""
import streamlit as st
from datetime import datetime
from coc_catalog import KELP_ANALYTE_CATALOG
from coc_pdf_engine import generate_coc_pdf

DOC_ID = "KELP-QMS-FORM-001"
DOC_VERSION = "1.1"

st.set_page_config(page_title="KELP Chain-of-Custody Generator", layout="wide")
st.title("\U0001f517 KELP Chain-of-Custody Generator")
st.caption(f"{DOC_ID} \u2014 Version {DOC_VERSION}")

# ── 1. Client Information ──
st.header("1. Client Information")
col1, col2 = st.columns(2)
with col1:
    company_name = st.text_input("Company Name")
    street_address = st.text_input("Street Address")
    project_number = st.text_input("Customer Project #")
    project_name = st.text_input("Project Name")
    site_info = st.text_input("Site Collection Info / Facility ID")
    county_state = st.text_input("County / State origin of sample(s)")
with col2:
    contact_name = st.text_input("Contact / Report To")
    phone = st.text_input("Phone #")
    email = st.text_input("E-Mail")
    cc_email = st.text_input("Cc E-Mail")
    invoice_to = st.text_input("Invoice To")
    invoice_email = st.text_input("Invoice E-Mail")

# ── 2. Project Details ──
st.header("2. Project Details")
col3, col4 = st.columns(2)
with col3:
    purchase_order = st.text_input("Purchase Order #")
    quote_number = st.text_input("Quote #")
    time_zone = st.selectbox("Sample Collection Time Zone", ["PT", "AK", "MT", "CT", "ET"])
    data_deliverable = st.selectbox("Data Deliverable", [
        "Level I (Std)", "Level II", "Level III", "Level IV", "Other"
    ])
with col4:
    regulatory_program = st.text_input("Regulatory Program (DW, RCRA, etc.)")
    reportable = st.selectbox("Reportable", ["Yes", "No"])
    rush = st.selectbox("Rush (Pre-approval required)", [
        "Standard (5-10 Day)", "Same Day", "1 Day", "2 Day", "3 Day", "4 Day", "5 Day"
    ])
    field_filtered = st.selectbox("Field Filtered", ["No", "Yes"])
    pwsid = st.text_input("DW PWSID # or WW Permit #")

# ── 3. Sample Information ──
st.header("3. Sample Information")
container_size = st.text_input("Container Size (e.g., 1L, 500mL)")
preservative_type = st.text_input("Preservative Type (e.g., None, HNO3, H2SO4)")
num_samples = st.number_input("Number of Samples", min_value=1, max_value=10, value=1)

samples = []
for i in range(num_samples):
    st.divider()
    st.subheader(f"Sample {i + 1}")

    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1:
        sid = st.text_input("Sample ID", key=f"sid_{i}")
        matrix = st.selectbox("Matrix", ["DW", "GW", "WW", "P", "SW", "OT"], key=f"mat_{i}")
    with sc2:
        comp_grab = st.selectbox("Comp/Grab", ["GRAB", "COMP"], key=f"cg_{i}")
        start_date = st.text_input("Start Date (MM/DD/YY)", key=f"sd_{i}")
    with sc3:
        start_time = st.text_input("Start Time", key=f"st_{i}")
        end_date = st.text_input("End Date (MM/DD/YY)", key=f"ed_{i}")
    with sc4:
        end_time = st.text_input("End Time", key=f"et_{i}")
        num_containers = st.text_input("# Containers", key=f"nc_{i}")

    # ── Analysis Requested: individual analyte selection by category ──
    st.markdown(f"**\U0001f9ea Analyses Requested \u2014 Sample {i + 1}**")
    st.caption("Expand a category to select specific analytes. Use 'Select all' for full panel.")

    sample_analyses = {}
    cat_cols = st.columns(2)
    cat_list = list(KELP_ANALYTE_CATALOG.keys())

    for ci, cat_name in enumerate(cat_list):
        cat_info = KELP_ANALYTE_CATALOG[cat_name]
        methods_display = ", ".join(cat_info["methods"])

        with cat_cols[ci % 2]:
            with st.expander(f"\U0001f9ea {cat_name}  ({methods_display})", expanded=False):
                select_all = st.checkbox(
                    f"Select all {cat_name} analytes",
                    key=f"all_{cat_name}_{i}",
                    value=False,
                )
                if select_all:
                    selected = st.multiselect(
                        "Analytes",
                        cat_info["analytes"],
                        default=cat_info["analytes"],
                        key=f"analytes_{cat_name}_{i}",
                    )
                else:
                    selected = st.multiselect(
                        "Analytes",
                        cat_info["analytes"],
                        key=f"analytes_{cat_name}_{i}",
                    )
                if selected:
                    sample_analyses[cat_name] = selected

    comment = st.text_input("Sample Comment", key=f"cmt_{i}")

    samples.append({
        "sample_id": sid,
        "matrix": matrix,
        "comp_grab": comp_grab,
        "start_date": start_date,
        "start_time": start_time,
        "end_date": end_date,
        "end_time": end_time,
        "num_containers": num_containers,
        "analyses": sample_analyses,
        "comment": comment,
    })

# ── 4. Analysis Summary Preview ──
st.header("4. Analysis Summary Preview")
any_analyses = False
for i, s in enumerate(samples):
    if s.get("analyses"):
        any_analyses = True
        sid_display = s["sample_id"] or f"Sample {i+1}"
        with st.expander(f"\U0001f4cb {sid_display} \u2014 Analysis Summary", expanded=True):
            for cat_name, analytes in s["analyses"].items():
                methods_str = ", ".join(KELP_ANALYTE_CATALOG[cat_name]["methods"])
                st.markdown(f"**{cat_name}** ({methods_str})")
                st.write(", ".join(analytes))

if not any_analyses:
    st.info("No analyses selected yet. Expand categories in each sample to pick analytes.")

# ── 5. Additional Info ──
st.header("5. Additional Information")
col5, col6 = st.columns(2)
with col5:
    additional_instructions = st.text_area("Additional Instructions from KELP")
    customer_remarks = st.text_area("Customer Remarks / Special Conditions / Possible Hazards")
with col6:
    num_coolers = st.text_input("# Coolers")
    delivery_method = st.selectbox("Delivery Method", ["In-Person", "FedEx", "UPS", "Other"])
    tracking_number = st.text_input("Tracking #")

# ── Logo upload ──
logo_file = st.file_uploader("Upload KELP Logo (optional)", type=["png", "jpg", "jpeg"])
logo_path = None
if logo_file:
    logo_path = f"/tmp/kelp_logo_{logo_file.name}"
    with open(logo_path, "wb") as f:
        f.write(logo_file.read())

# ── Generate PDF ──
st.divider()
if st.button("\U0001f5a8\ufe0f Generate Chain-of-Custody PDF", type="primary"):
    data = {
        "company_name": company_name, "street_address": street_address,
        "contact_name": contact_name, "phone": phone, "email": email,
        "cc_email": cc_email, "project_number": project_number,
        "project_name": project_name, "invoice_to": invoice_to,
        "invoice_email": invoice_email, "site_info": site_info,
        "purchase_order": purchase_order, "quote_number": quote_number,
        "county_state": county_state, "time_zone": time_zone,
        "data_deliverable": data_deliverable, "regulatory_program": regulatory_program,
        "reportable": reportable, "rush": rush, "field_filtered": field_filtered,
        "pwsid": pwsid, "container_size": container_size,
        "preservative_type": preservative_type, "samples": samples,
        "additional_instructions": additional_instructions,
        "customer_remarks": customer_remarks, "num_coolers": num_coolers,
        "delivery_method": delivery_method, "tracking_number": tracking_number,
    }

    pdf_buf = generate_coc_pdf(data, logo_path=logo_path)

    st.success("\u2705 COC generated! 3 pages: Form + Analysis Detail + Instructions")
    st.download_button(
        "\u2b07\ufe0f Download Chain-of-Custody PDF",
        pdf_buf,
        file_name=f"KELP_CoC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
        mime="application/pdf",
    )

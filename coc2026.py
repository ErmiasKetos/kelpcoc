"""
coc2026.py - KELP COC Streamlit App v3
- Date/time picker widgets (no manual typing)
- Matrix-aware method display
- Timezone selector
"""
import streamlit as st
import datetime
from coc_catalog import KELP_ANALYTE_CATALOG, get_methods_flat, generate_coc_id
from coc_pdf_engine import generate_coc_pdf

st.set_page_config(page_title="KELP COC Generator", layout="wide", page_icon="\U0001f4c4")
st.title("\U0001f4c4 KELP Chain-of-Custody Generator")

MATRIX_OPTIONS = ["DW", "GW", "WW", "SW", "P", "OT"]
MATRIX_LABELS = {
    "DW": "Drinking Water", "GW": "Ground Water", "WW": "Wastewater",
    "SW": "Surface Water", "P": "Product", "OT": "Other"
}
TZ_OPTIONS = ["PT", "AK", "MT", "CT", "ET"]
COMP_GRAB = ["GRAB", "COMP"]

# === CLIENT INFO ===
st.header("1\ufe0f\u20e3 Client Information")
c1, c2 = st.columns(2)
with c1:
    company_name = st.text_input("Company Name")
    client_street = st.text_input("Street Address")
    client_city_state_zip = st.text_input("City, State, ZIP")
    customer_project = st.text_input("Customer Project #")
    project_name = st.text_input("Project Name")
with c2:
    contact_name = st.text_input("Contact / Report To")
    phone = st.text_input("Phone #")
    email = st.text_input("E-Mail")
    cc_email = st.text_input("Cc E-Mail")
    invoice_to = st.text_input("Invoice To")

st.header("2\ufe0f\u20e3 Project Details")
c3, c4 = st.columns(2)
with c3:
    site_info = st.text_input("Site Collection Info / Facility ID")
    county_state = st.text_input("County / State origin of sample(s)")
    purchase_order = st.text_input("Purchase Order #")
    quote_number = st.text_input("Quote #")
with c4:
    invoice_email = st.text_input("Invoice E-mail")
    container_size = st.selectbox("Container Size", ["500mL", "1L", "250mL", "125mL", "100mL", "Other"])
    preservative = st.selectbox("Preservative Type", [
        "None", "HNO3", "H2SO4", "HCl", "NaOH",
        "Zn Acetate", "NaHSO4", "Sod.Thiosulfate", "Ascorbic Acid", "MeOH", "Other"
    ])

st.header("3\ufe0f\u20e3 Collection & Reporting Options")
c5, c6, c7 = st.columns(3)
with c5:
    time_zone = st.selectbox("Sample Collection Time Zone", TZ_OPTIONS, index=1)
    data_deliverable = st.selectbox("Data Deliverables", [
        "Level I (Std)", "Level II", "Level III", "Other"
    ])
    field_filtered = st.selectbox("Field Filtered", ["No", "Yes"])
with c6:
    reportable = st.selectbox("Reportable", ["Yes", "No"])
    rush = st.selectbox("Rush (Pre-approval required)", [
        "Standard (5-10 Day)", "Same Day", "1 Day", "2 Day", "3 Day", "4 Day", "5 Day", "Other"
    ])
with c7:
    received_on_ice = st.selectbox("Sample Received on Ice", ["Yes", "No"])
    delivery_method = st.selectbox("Delivery Method", ["In-Person", "FedEx", "UPS", "Courier", "Other"])

# === SAMPLES ===
st.header("4\ufe0f\u20e3 Sample Information")
num_samples = st.number_input("Number of Samples", 1, 10, 1)

samples = []
for i in range(num_samples):
    st.subheader(f"Sample {i+1}")
    sc1, sc2, sc3, sc4 = st.columns([3, 1, 1, 1])
    with sc1:
        sample_id = st.text_input("Customer Sample ID", key=f"sid_{i}")
    with sc2:
        matrix = st.selectbox("Matrix", MATRIX_OPTIONS, key=f"mat_{i}",
                              help=" | ".join([f"{k}={v}" for k, v in MATRIX_LABELS.items()]))
    with sc3:
        comp_grab = st.selectbox("Comp/Grab", COMP_GRAB, key=f"cg_{i}")
    with sc4:
        num_containers = st.number_input("# Containers", 1, 20, 1, key=f"nc_{i}")

    # Date/time pickers
    dt1, dt2, dt3, dt4 = st.columns(4)
    with dt1:
        if comp_grab == "COMP":
            start_date = st.date_input("Composite Start Date", value=None, key=f"sd_{i}")
        else:
            start_date = None
    with dt2:
        if comp_grab == "COMP":
            start_time = st.time_input("Composite Start Time", value=None, key=f"st_{i}")
        else:
            start_time = None
    with dt3:
        coll_date = st.date_input(
            "Collected Date" if comp_grab == "GRAB" else "Composite End Date",
            value=None, key=f"cd_{i}"
        )
    with dt4:
        coll_time = st.time_input(
            "Collected Time" if comp_grab == "GRAB" else "Composite End Time",
            value=None, key=f"ct_{i}"
        )

    # Residual chlorine
    rc1, rc2 = st.columns([1, 1])
    with rc1:
        res_cl_result = st.text_input("Residual Chlorine Result", key=f"rcr_{i}")
    with rc2:
        res_cl_units = st.selectbox("Residual Chlorine Units", ["", "mg/L", "ppm"], key=f"rcu_{i}")

    # Analysis selection with matrix-aware method display
    sample_analyses = {}
    st.markdown(f"**Analysis Requested** *(Matrix: {matrix} — {MATRIX_LABELS[matrix]})*")
    for ci, cat_name in enumerate(KELP_ANALYTE_CATALOG.keys()):
        cat_info = KELP_ANALYTE_CATALOG[cat_name]
        # Show appropriate method for this sample's matrix
        if matrix in ("DW",):
            method_str = ", ".join(cat_info["methods"].get("potable", []))
        else:
            method_str = ", ".join(cat_info["methods"].get("nonpotable", []))
        if cat_name != "Packages":
            with st.expander(f"\U0001f9ea {cat_name}  \u2014  {method_str}", expanded=False):
                sa = st.checkbox(f"Select all {cat_name}", key=f"all_{cat_name}_{i}", value=False)
                sel = st.multiselect(f"Analytes", cat_info["analytes"],
                    default=cat_info["analytes"] if sa else [], key=f"a_{cat_name}_{i}")
                if sel: sample_analyses[cat_name] = sel

    comment = st.text_input("Sample Comment", key=f"cmt_{i}")

    samples.append({
        "sample_id": sample_id, "matrix": matrix, "comp_grab": comp_grab,
        "start_date": start_date.strftime("%m/%d/%Y") if start_date else "",
        "start_time": start_time.strftime("%H:%M") if start_time else "",
        "end_date": coll_date.strftime("%m/%d/%Y") if coll_date else "",
        "end_time": coll_time.strftime("%H:%M") if coll_time else "",
        "num_containers": str(num_containers), "analyses": sample_analyses,
        "comment": comment,
        "res_cl_result": res_cl_result, "res_cl_units": res_cl_units,
    })

    # Preview
    if sample_analyses:
        with st.container():
            for cn, al in sample_analyses.items():
                cat_info = KELP_ANALYTE_CATALOG[cn]
                if matrix in ("DW",):
                    mstr = ", ".join(cat_info["methods"].get("potable", []))
                else:
                    mstr = ", ".join(cat_info["methods"].get("nonpotable", []))
                st.markdown(f"**{cn}** ({mstr}): {', '.join(al)}")
    else:
        st.info("No analyses selected yet.")

# === BOTTOM FIELDS ===
st.header("5\ufe0f\u20e3 Additional Information")
ac1, ac2 = st.columns(2)
with ac1:
    additional_instructions = st.text_area("Additional Instructions for KELP", height=60)
with ac2:
    customer_remarks = st.text_area("Customer Remarks / Special Conditions / Hazards", height=60)

# === GENERATE ===
st.divider()
if st.button("\U0001f4e4 Generate COC PDF", type="primary", use_container_width=True):
    coc_data = {
        "company_name": company_name, "client_address": client_street,
        "client_address_2": client_city_state_zip,
        "contact_name": contact_name, "phone": phone, "email": email, "cc_email": cc_email,
        "project_number": customer_project, "project_name": project_name,
        "invoice_to": invoice_to, "invoice_email": invoice_email,
        "site_info": site_info, "county_state": county_state,
        "purchase_order": purchase_order, "quote_number": quote_number,
        "container_size": container_size, "preservative_type": preservative,
        "time_zone": time_zone, "data_deliverable": data_deliverable,
        "field_filtered": field_filtered, "reportable": reportable, "rush": rush,
        "received_on_ice": received_on_ice, "delivery_method": delivery_method,
        "additional_instructions": additional_instructions,
        "customer_remarks": customer_remarks,
        "samples": samples,
    }
    buf, coc_id = generate_coc_pdf(coc_data)
    st.success(f"COC generated: **{coc_id}**")
    st.download_button(
        label="\U0001f4be Download COC PDF",
        data=buf.getvalue(),
        file_name=f"KELP_CoC_{coc_id}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
        mime="application/pdf",
        type="primary"
    )

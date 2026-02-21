"""
coc_catalog.py - KELP Analyte Catalog v3
Matrix-aware methods + hybrid chemical symbols
"""
import random, datetime

# Methods keyed by matrix type: "potable" covers DW; "nonpotable" covers GW, WW, SW, P, OT
KELP_ANALYTE_CATALOG = {
    "Metals": {
        "methods": {
            "potable": ["EPA 200.8"],
            "nonpotable": ["EPA 6020B"],
        },
        "analytes": [
            "Aluminum","Antimony","Arsenic","Barium","Beryllium","Boron",
            "Cadmium","Calcium","Chromium","Chromium (VI)","Cobalt","Copper",
            "Iron","Lead","Magnesium","Manganese","Mercury","Molybdenum",
            "Nickel","Potassium","Selenium","Silicon","Silver","Sodium",
            "Thallium","Thorium","Uranium","Vanadium","Zinc",
        ],
    },
    "Inorganics": {
        "methods": {
            "potable": ["EPA 300.1"],
            "nonpotable": ["EPA 300.1"],
        },
        "analytes": [
            "Bromate","Bromide","Chlorate","Chloride","Chlorite",
            "Cyanide - Available","Cyanide - Total","Fluoride",
            "Nitrate","Nitrite","Perchlorate","Phosphate - Ortho","Sulfate",
        ],
    },
    "Physical/General Chemistry": {
        "methods": {
            "potable": ["SM/EPA"],
            "nonpotable": ["SM/EPA"],
        },
        "analytes": [
            "pH","Temperature","Turbidity","Conductivity","Dissolved Oxygen",
            "Alkalinity","Hardness - Total","Total Dissolved Solids",
            "Total Suspended Solids","Total Solids",
            "BOD (5-day)","BOD - Carbonaceous","Chemical Oxygen Demand",
        ],
    },
    "Nutrients": {
        "methods": {
            "potable": ["EPA 350/351/365"],
            "nonpotable": ["EPA 350/351/365"],
        },
        "analytes": [
            "Ammonia (as N)","Kjeldahl Nitrogen - Total",
            "Phosphorus - Total","Sulfide (as S)","Sulfite (as SO3)",
        ],
    },
    "Organics": {
        "methods": {
            "potable": ["EPA 415/SM5540"],
            "nonpotable": ["EPA 415/SM5540"],
        },
        "analytes": ["Dissolved Organic Carbon","Total Organic Carbon","Surfactants (MBAS)"],
    },
    "PFAS Testing": {
        "methods": {
            "potable": ["EPA 537.1"],
            "nonpotable": ["EPA 1633A"],
        },
        "analytes": [
            "PFAS 3-Compound (PFNA, PFOA, PFOS)","PFAS 14-Compound",
            "PFAS 18-Compound","PFAS 25-Compound","PFAS 40-Compound",
        ],
    },
    "Disinfection": {
        "methods": {
            "potable": ["SM 4500-Cl"],
            "nonpotable": ["SM 4500-Cl"],
        },
        "analytes": [
            "Chlorine - Free","Chlorine - Free (DPD)","Chlorine - Total (DPD)",
            "Chlorine - Combined","Chloramines (Monochloramine)","Chlorine Dioxide",
        ],
    },
    "Packages": {
        "methods": {
            "potable": ["Multiple"],
            "nonpotable": ["Multiple"],
        },
        "analytes": [
            "Essential Home Water Test","Complete Homeowner Package",
            "Conventional Loan Testing Package","Real Estate Well Water Package",
            "Agricultural Irrigation Package","Food & Beverage Water Quality Package",
            "PFAS Home Safety Package",
        ],
    },
}

CAT_SHORT_MAP = {
    "Metals": "Metals",
    "Inorganics": "Inorganics",
    "Physical/General Chemistry": "Phys/Gen Chem",
    "Nutrients": "Nutrients",
    "Organics": "Organics",
    "PFAS Testing": "PFAS Testing",
    "Disinfection": "Disinfection",
    "Packages": "Packages",
}

# Matrices classified as potable vs nonpotable
POTABLE_MATRICES = {"DW"}
NONPOTABLE_MATRICES = {"GW", "WW", "SW", "P", "OT"}

SYMBOL_MAP = {
    # Metals
    "Aluminum": "Al", "Antimony": "Sb", "Arsenic": "As", "Barium": "Ba",
    "Beryllium": "Be", "Boron": "B", "Cadmium": "Cd", "Calcium": "Ca",
    "Chromium": "Cr", "Chromium (VI)": "Cr(VI)", "Cobalt": "Co", "Copper": "Cu",
    "Iron": "Fe", "Lead": "Pb", "Magnesium": "Mg", "Manganese": "Mn",
    "Mercury": "Hg", "Molybdenum": "Mo", "Nickel": "Ni", "Potassium": "K",
    "Selenium": "Se", "Silicon": "Si", "Silver": "Ag", "Sodium": "Na",
    "Thallium": "Tl", "Thorium": "Th", "Uranium": "U", "Vanadium": "V", "Zinc": "Zn",
    # Inorganics
    "Cyanide - Available": "CN- Avail.",
    "Cyanide - Total": "CN- Total",
    "Phosphate - Ortho": "Ortho-PO4",
    "Perchlorate": "ClO4-",
    # Nutrients
    "Ammonia (as N)": "NH3-N",
    "Kjeldahl Nitrogen - Total": "TKN",
    "Phosphorus - Total": "Total P",
    "Sulfide (as S)": "Sulfide-S",
    "Sulfite (as SO3)": "Sulfite-SO3",
    # Disinfection
    "Chloramines (Monochloramine)": "Chloramines (NH2Cl)",
    "Chlorine - Free (DPD)": "Cl2 Free (DPD)",
    "Chlorine - Total (DPD)": "Cl2 Total (DPD)",
    "Chlorine - Free": "Cl2 Free",
    "Chlorine - Combined": "Cl2 Combined",
    "Chlorine Dioxide": "ClO2",
    # Phys/Gen Chem
    "Total Dissolved Solids": "TDS",
    "Total Suspended Solids": "TSS",
    "Chemical Oxygen Demand": "COD",
    "BOD - Carbonaceous": "CBOD",
    # Organics
    "Dissolved Organic Carbon": "DOC",
    "Total Organic Carbon": "TOC",
    "Surfactants (MBAS)": "MBAS",
}

def to_symbol(analyte_name):
    """Convert analyte name to hybrid symbol if available."""
    return SYMBOL_MAP.get(analyte_name, analyte_name)


def get_methods_for_category(cat_name, matrices):
    """Get combined method string based on which matrices are present.
    
    Args:
        cat_name: Category name (full name from catalog)
        matrices: set of matrix codes present on the COC (e.g., {"DW", "WW"})
    
    Returns:
        Method string like "EPA 200.8" or "EPA 200.8/6020B" if mixed
    """
    if cat_name not in KELP_ANALYTE_CATALOG:
        return ""
    minfo = KELP_ANALYTE_CATALOG[cat_name]["methods"]
    
    has_potable = bool(matrices & POTABLE_MATRICES)
    has_nonpotable = bool(matrices & NONPOTABLE_MATRICES)
    
    methods = []
    if has_potable:
        methods.extend(minfo.get("potable", []))
    if has_nonpotable:
        for m in minfo.get("nonpotable", []):
            if m not in methods:
                methods.append(m)
    
    # If no matrix info available, show both if different
    if not methods:
        methods.extend(minfo.get("potable", []))
        for m in minfo.get("nonpotable", []):
            if m not in methods:
                methods.append(m)
    
    return ", ".join(methods)


def get_methods_flat(cat_name):
    """Get all unique methods for a category (for display in Streamlit)."""
    if cat_name not in KELP_ANALYTE_CATALOG:
        return ""
    minfo = KELP_ANALYTE_CATALOG[cat_name]["methods"]
    methods = []
    for ml in minfo.values():
        for m in ml:
            if m not in methods:
                methods.append(m)
    return ", ".join(methods)


def generate_coc_id():
    now = datetime.datetime.now()
    seq = random.randint(1000, 9999)
    return f"KELP-COC-{now.strftime('%y%m%d')}-{seq}"

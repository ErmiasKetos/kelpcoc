"""
coc_catalog.py - KELP Analyte Catalog with hybrid chemical symbols
"""
import random, datetime

KELP_ANALYTE_CATALOG = {
    "Metals": {
        "methods": ["EPA 200.8/6020B"],
        "analytes": [
            "Aluminum","Antimony","Arsenic","Barium","Beryllium","Boron",
            "Cadmium","Calcium","Chromium","Chromium (VI)","Cobalt","Copper",
            "Iron","Lead","Magnesium","Manganese","Mercury","Molybdenum",
            "Nickel","Potassium","Selenium","Silicon","Silver","Sodium",
            "Thallium","Thorium","Uranium","Vanadium","Zinc",
        ],
    },
    "Inorganics": {
        "methods": ["EPA 300.1"],
        "analytes": [
            "Bromate","Bromide","Chlorate","Chloride","Chlorite",
            "Cyanide - Available","Cyanide - Total","Fluoride",
            "Nitrate","Nitrite","Perchlorate","Phosphate - Ortho","Sulfate",
        ],
    },
    "Physical/General Chemistry": {
        "methods": ["SM/EPA"],
        "analytes": [
            "pH","Temperature","Turbidity","Conductivity","Dissolved Oxygen",
            "Alkalinity","Hardness - Total","Total Dissolved Solids",
            "Total Suspended Solids","Total Solids",
            "BOD (5-day)","BOD - Carbonaceous","Chemical Oxygen Demand",
        ],
    },
    "Nutrients": {
        "methods": ["EPA 350/351/365"],
        "analytes": [
            "Ammonia (as N)","Kjeldahl Nitrogen - Total",
            "Phosphorus - Total","Sulfide (as S)","Sulfite (as SO3)",
        ],
    },
    "Organics": {
        "methods": ["EPA 415/SM5540"],
        "analytes": ["Dissolved Organic Carbon","Total Organic Carbon","Surfactants (MBAS)"],
    },
    "PFAS Testing": {
        "methods": ["EPA 537/1633"],
        "analytes": [
            "PFAS 3-Compound (PFNA, PFOA, PFOS)","PFAS 14-Compound",
            "PFAS 18-Compound","PFAS 25-Compound","PFAS 40-Compound",
        ],
    },
    "Disinfection": {
        "methods": ["SM 4500-Cl"],
        "analytes": [
            "Chlorine - Free","Chlorine - Free (DPD)","Chlorine - Total (DPD)",
            "Chlorine - Combined","Chloramines (Monochloramine)","Chlorine Dioxide",
        ],
    },
    "Packages": {
        "methods": ["Multiple"],
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

# Hybrid chemical symbol map: only where it saves significant space
# Metals -> element symbols (universally recognized, huge space savings)
# Inorganics -> abbreviate cyanide & perchlorate
# Nutrients -> standard abbreviations (TKN, NH3-N)
# Disinfection -> Cl2, NH2Cl abbreviations
SYMBOL_MAP = {
    # Metals
    "Aluminum": "Al", "Antimony": "Sb", "Arsenic": "As", "Barium": "Ba",
    "Beryllium": "Be", "Boron": "B", "Cadmium": "Cd", "Calcium": "Ca",
    "Chromium": "Cr", "Chromium (VI)": "Cr(VI)", "Cobalt": "Co", "Copper": "Cu",
    "Iron": "Fe", "Lead": "Pb", "Magnesium": "Mg", "Manganese": "Mn",
    "Mercury": "Hg", "Molybdenum": "Mo", "Nickel": "Ni", "Potassium": "K",
    "Selenium": "Se", "Silicon": "Si", "Silver": "Ag", "Sodium": "Na",
    "Thallium": "Tl", "Thorium": "Th", "Uranium": "U", "Vanadium": "V", "Zinc": "Zn",
    # Inorganics - plain ASCII only
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
    # Disinfection - plain ASCII
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

def generate_coc_id():
    now = datetime.datetime.now()
    seq = random.randint(1000, 9999)
    return f"KELP-COC-{now.strftime('%y%m%d')}-{seq}"

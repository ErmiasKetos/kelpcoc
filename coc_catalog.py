"""
coc_catalog.py — KELP Analyte Catalog & Constants v16
Sourced from KELP CA ELAP Price List (SKU 9/24/25)
"""
import random
from datetime import datetime

KELP_ANALYTE_CATALOG = {
    "Metals": {
        "methods": ["EPA 200.8/6020B"],
        "analytes": [
            "Aluminum", "Antimony", "Arsenic", "Barium", "Beryllium", "Boron",
            "Cadmium", "Calcium", "Chromium", "Chromium (VI)", "Cobalt", "Copper",
            "Iron", "Lead", "Magnesium", "Manganese", "Mercury", "Molybdenum",
            "Nickel", "Potassium", "Selenium", "Silicon", "Silver", "Sodium",
            "Thallium", "Thorium", "Uranium", "Vanadium", "Zinc",
        ],
    },
    "Inorganics": {
        "methods": ["EPA 300.1"],
        "analytes": [
            "Bromate", "Bromide", "Chlorate", "Chloride", "Chlorite",
            "Cyanide - Available", "Cyanide - Total", "Fluoride",
            "Nitrate", "Nitrite", "Perchlorate",
            "Phosphate - Ortho", "Sulfate",
        ],
    },
    "Physical/General Chemistry": {
        "methods": ["SM/EPA"],
        "analytes": [
            "pH", "Temperature", "Turbidity", "Conductivity", "Dissolved Oxygen",
            "Alkalinity", "Hardness - Total",
            "Total Dissolved Solids", "Total Suspended Solids", "Total Solids",
            "BOD (5-day)", "BOD - Carbonaceous", "Chemical Oxygen Demand",
        ],
    },
    "Nutrients": {
        "methods": ["EPA 350/351/365"],
        "analytes": [
            "Ammonia (as N)", "Kjeldahl Nitrogen - Total",
            "Phosphorus - Total", "Sulfide (as S)", "Sulfite (as SO3)",
        ],
    },
    "Organics": {
        "methods": ["EPA 415/SM5540"],
        "analytes": [
            "Dissolved Organic Carbon", "Total Organic Carbon",
            "Surfactants (MBAS)",
        ],
    },
    "PFAS Testing": {
        "methods": ["EPA 537/1633"],
        "analytes": [
            "PFAS 3-Compound (PFNA, PFOA, PFOS)",
            "PFAS 14-Compound", "PFAS 18-Compound",
            "PFAS 25-Compound", "PFAS 40-Compound",
        ],
    },
    "Disinfection": {
        "methods": ["SM 4500-Cl"],
        "analytes": [
            "Chlorine - Free", "Chlorine - Free (DPD)", "Chlorine - Total (DPD)",
            "Chlorine - Combined", "Chloramines (Monochloramine)", "Chlorine Dioxide",
        ],
    },
    "Packages": {
        "methods": ["Multiple"],
        "analytes": [
            "Essential Home Water Test",
            "Complete Homeowner Package",
            "Conventional Loan Testing Package",
            "Real Estate Well Water Package",
            "Agricultural Irrigation Package",
            "Food & Beverage Water Quality Package",
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


def generate_coc_id():
    """
    Generate a unique COC document ID per TNI V1M2 / ISO 17025.
    Format: KELP-COC-YYMMDD-NNNN
    In production, NNNN would be a sequential counter from LIMS/database.
    """
    now = datetime.now()
    date_part = now.strftime("%y%m%d")
    seq = f"{random.randint(1, 9999):04d}"
    return f"KELP-COC-{date_part}-{seq}"

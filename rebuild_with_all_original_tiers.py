#!/usr/bin/env python3
"""
Rebuild workbook with ALL original Launch Inventory tiers
- Add all missing Tier 5, 6, 7, 8, 9 items
- Maintain all current items (294 + 9 CGM/ankle = 303)
- Add back missing categories to reach full BOC coverage
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime
import os

print("="*100)
print("REBUILDING WITH ALL ORIGINAL LAUNCH INVENTORY TIERS")
print("="*100)

# ============================================================================
# STEP 1: LOAD CURRENT INVENTORY (303 items)
# ============================================================================
print("\nSTEP 1: Loading current inventory (294 + 9 added)...")

current_df = pd.read_csv("/home/user/InitialInventoryScratch/MASTER_INVENTORY_PLAN.csv")
print(f"  Current inventory: {len(current_df)} products")

# Add the 9 items we just added (5 ankle + 4 CGM)
added_ankle = pd.DataFrame([
    {'HCPCS_Code': 'L1906', 'BOC_Category': 'OR03', 'Description': 'AFO - multiligamentous ankle support', 'Quantity': 6, 'Estimated_Unit_Cost': 80.0, 'Total_Cost': 480.0, 'Medicare_Rate': 117.69, 'Priority_Score': 80, 'Source': 'CUSTOMER', 'Customers': 'DR_NAS'},
    {'HCPCS_Code': 'L4361', 'BOC_Category': 'OR03', 'Description': 'Walking boot - pneumatic/vacuum (CAM WALKER)', 'Quantity': 6, 'Estimated_Unit_Cost': 70.0, 'Total_Cost': 420.0, 'Medicare_Rate': 103.88, 'Priority_Score': 100, 'Source': 'CUSTOMER', 'Customers': 'DR_NAS'},
    {'HCPCS_Code': 'L4350', 'BOC_Category': 'OR03', 'Description': 'Ankle control orthosis, stirrup style', 'Quantity': 2, 'Estimated_Unit_Cost': 80.0, 'Total_Cost': 160.0, 'Medicare_Rate': 103.88, 'Priority_Score': 80, 'Source': 'CUSTOMER', 'Customers': 'DR_NAS'},
    {'HCPCS_Code': 'L4370', 'BOC_Category': 'OR03', 'Description': 'Pneumatic full leg splint', 'Quantity': 2, 'Estimated_Unit_Cost': 110.0, 'Total_Cost': 220.0, 'Medicare_Rate': 150.0, 'Priority_Score': 80, 'Source': 'CUSTOMER', 'Customers': 'DR_NAS'},
    {'HCPCS_Code': 'L4387', 'BOC_Category': 'OR03', 'Description': 'Walking boot - non-pneumatic', 'Quantity': 2, 'Estimated_Unit_Cost': 135.0, 'Total_Cost': 270.0, 'Medicare_Rate': 103.88, 'Priority_Score': 80, 'Source': 'CUSTOMER', 'Customers': 'DR_NAS'},
    {'HCPCS_Code': 'E2103', 'BOC_Category': 'DM06', 'Description': 'Dexcom G7 CGM Receiver', 'Quantity': 5, 'Estimated_Unit_Cost': 400.0, 'Total_Cost': 2000.0, 'Medicare_Rate': None, 'Priority_Score': 70, 'Source': 'CUSTOMER', 'Customers': 'MOYHINOR (CGM_STRATEGY)'},
    {'HCPCS_Code': 'E2103', 'BOC_Category': 'DM06', 'Description': 'Freestyle Libre CGM Reader', 'Quantity': 5, 'Estimated_Unit_Cost': 400.0, 'Total_Cost': 2000.0, 'Medicare_Rate': None, 'Priority_Score': 70, 'Source': 'CUSTOMER', 'Customers': 'MOYHINOR (CGM_STRATEGY)'},
    {'HCPCS_Code': 'A4239', 'BOC_Category': 'DM06', 'Description': 'Dexcom G7 CGM Sensors, 30-day', 'Quantity': 10, 'Estimated_Unit_Cost': 90.0, 'Total_Cost': 900.0, 'Medicare_Rate': None, 'Priority_Score': 70, 'Source': 'CUSTOMER', 'Customers': 'MOYHINOR (CGM_STRATEGY)'},
    {'HCPCS_Code': 'A4239', 'BOC_Category': 'DM06', 'Description': 'Freestyle Libre CGM Sensors, 30-day', 'Quantity': 10, 'Estimated_Unit_Cost': 70.0, 'Total_Cost': 700.0, 'Medicare_Rate': None, 'Priority_Score': 70, 'Source': 'CUSTOMER', 'Customers': 'MOYHINOR (CGM_STRATEGY)'},
])

current_with_added = pd.concat([current_df, added_ankle], ignore_index=True)
print(f"  With recently added items: {len(current_with_added)} products")

# ============================================================================
# STEP 2: DEFINE ALL MISSING TIER ITEMS
# ============================================================================
print("\nSTEP 2: Defining missing tier items from original plan...")

# Missing from TIER 1
tier1_missing = [
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0163', 'Description': 'Commode chair - fixed arms', 'Quantity': 4, 'Est_Unit_Cost': 150},
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0165', 'Description': 'Commode chair - detachable arms', 'Quantity': 4, 'Est_Unit_Cost': 175},
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0168', 'Description': 'Commode chair - bariatric', 'Quantity': 2, 'Est_Unit_Cost': 250},
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0275', 'Description': 'Bedpan - standard', 'Quantity': 10, 'Est_Unit_Cost': 8},
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0276', 'Description': 'Bedpan - fracture', 'Quantity': 10, 'Est_Unit_Cost': 12},
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0160', 'Description': 'Sitz bath - portable', 'Quantity': 2, 'Est_Unit_Cost': 35},
    {'Tier': 'TIER 1', 'BOC_Category': 'DM02', 'HCPCS_Code': 'E0161', 'Description': 'Sitz bath - with faucet', 'Quantity': 1, 'Est_Unit_Cost': 65},
]

# Missing from TIER 3
tier3_missing = [
    {'Tier': 'TIER 3', 'BOC_Category': 'DM29', 'HCPCS_Code': 'E2001', 'Description': 'Suction pump - portable', 'Quantity': 3, 'Est_Unit_Cost': 350},
    {'Tier': 'TIER 3', 'BOC_Category': 'DM29', 'HCPCS_Code': 'A6590', 'Description': 'Suction canister', 'Quantity': 10, 'Est_Unit_Cost': 25},
    {'Tier': 'TIER 3', 'BOC_Category': 'DM29', 'HCPCS_Code': 'A7001', 'Description': 'Suction tubing set', 'Quantity': 20, 'Est_Unit_Cost': 8},
]

# Missing from TIER 4
tier4_missing = [
    {'Tier': 'TIER 4', 'BOC_Category': 'DM25', 'HCPCS_Code': 'A4224', 'Description': 'Insulin infusion set', 'Quantity': 25, 'Est_Unit_Cost': 12},
    {'Tier': 'TIER 4', 'BOC_Category': 'DM25', 'HCPCS_Code': 'A4225', 'Description': 'Insulin reservoir', 'Quantity': 20, 'Est_Unit_Cost': 15},
    {'Tier': 'TIER 4', 'BOC_Category': 'DM25', 'HCPCS_Code': 'K0552', 'Description': 'Insulin pump supplies', 'Quantity': 15, 'Est_Unit_Cost': 35},
]

# Missing from TIER 5 (expand nebulizers, add TENS, NMES, Pneumatic)
tier5_missing = [
    # Nebulizers (expand from 1 to full set)
    {'Tier': 'TIER 5', 'BOC_Category': 'R07', 'HCPCS_Code': 'E0572', 'Description': 'Nebulizer - aerosol compressor', 'Quantity': 3, 'Est_Unit_Cost': 85},
    {'Tier': 'TIER 5', 'BOC_Category': 'R07', 'HCPCS_Code': 'A7015', 'Description': 'Nebulizer mask - adult', 'Quantity': 30, 'Est_Unit_Cost': 4},
    {'Tier': 'TIER 5', 'BOC_Category': 'R07', 'HCPCS_Code': 'A7016', 'Description': 'Nebulizer mask - pediatric', 'Quantity': 15, 'Est_Unit_Cost': 4},
    {'Tier': 'TIER 5', 'BOC_Category': 'R07', 'HCPCS_Code': 'A7003', 'Description': 'Nebulizer tubing - disposable', 'Quantity': 40, 'Est_Unit_Cost': 3},
    {'Tier': 'TIER 5', 'BOC_Category': 'R07', 'HCPCS_Code': 'A7013', 'Description': 'Nebulizer filter - disposable', 'Quantity': 50, 'Est_Unit_Cost': 2},
    # TENS Units
    {'Tier': 'TIER 5', 'BOC_Category': 'DM22', 'HCPCS_Code': 'E0720', 'Description': 'TENS device - 2 lead', 'Quantity': 6, 'Est_Unit_Cost': 120},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM22', 'HCPCS_Code': 'E0730', 'Description': 'TENS device - 4 lead', 'Quantity': 4, 'Est_Unit_Cost': 180},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM22', 'HCPCS_Code': 'A4595', 'Description': 'TENS electrode pads', 'Quantity': 80, 'Est_Unit_Cost': 8},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM22', 'HCPCS_Code': 'A4557', 'Description': 'TENS lead wires', 'Quantity': 15, 'Est_Unit_Cost': 12},
    # NMES
    {'Tier': 'TIER 5', 'BOC_Category': 'DM16', 'HCPCS_Code': 'E0731', 'Description': 'NMES device - form fitting conduction garment', 'Quantity': 2, 'Est_Unit_Cost': 200},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM16', 'HCPCS_Code': 'E0744', 'Description': 'NMES device - orofacial', 'Quantity': 2, 'Est_Unit_Cost': 250},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM16', 'HCPCS_Code': 'E0745', 'Description': 'NMES device - other', 'Quantity': 1, 'Est_Unit_Cost': 300},
    # Pneumatic Compression
    {'Tier': 'TIER 5', 'BOC_Category': 'DM18', 'HCPCS_Code': 'E0651', 'Description': 'Pneumatic compression pump - non-segmental home', 'Quantity': 1, 'Est_Unit_Cost': 450},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM18', 'HCPCS_Code': 'E0652', 'Description': 'Pneumatic compression pump - segmental home', 'Quantity': 1, 'Est_Unit_Cost': 650},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM18', 'HCPCS_Code': 'E0671', 'Description': 'Pneumatic compression sleeve - segmental arm', 'Quantity': 3, 'Est_Unit_Cost': 120},
    {'Tier': 'TIER 5', 'BOC_Category': 'DM18', 'HCPCS_Code': 'E0672', 'Description': 'Pneumatic compression sleeve - segmental leg', 'Quantity': 5, 'Est_Unit_Cost': 140},
]

# Missing from TIER 6 (all of it)
tier6_missing = [
    # Support Surfaces
    {'Tier': 'TIER 6', 'BOC_Category': 'DM20', 'HCPCS_Code': 'E0199', 'Description': 'Dry pressure pad - standard mattress width', 'Quantity': 4, 'Est_Unit_Cost': 85},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM20', 'HCPCS_Code': 'E0196', 'Description': 'Gel pressure pad - standard mattress', 'Quantity': 3, 'Est_Unit_Cost': 125},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM20', 'HCPCS_Code': 'E0186', 'Description': 'Air pressure mattress', 'Quantity': 2, 'Est_Unit_Cost': 450},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM20', 'HCPCS_Code': 'E0197', 'Description': 'Air pressure pad - standard mattress', 'Quantity': 4, 'Est_Unit_Cost': 95},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM20', 'HCPCS_Code': 'E0198', 'Description': 'Water pressure pad - standard mattress', 'Quantity': 4, 'Est_Unit_Cost': 110},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM20', 'HCPCS_Code': 'E0277', 'Description': 'Low air loss mattress', 'Quantity': 1, 'Est_Unit_Cost': 850},
    # Heat & Cold
    {'Tier': 'TIER 6', 'BOC_Category': 'DM08', 'HCPCS_Code': 'E0215', 'Description': 'Electric heating pad - standard', 'Quantity': 10, 'Est_Unit_Cost': 25},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM08', 'HCPCS_Code': 'E0217', 'Description': 'Electric heating pad - moist', 'Quantity': 10, 'Est_Unit_Cost': 35},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM08', 'HCPCS_Code': 'E0230', 'Description': 'Ice cap or collar', 'Quantity': 10, 'Est_Unit_Cost': 15},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM08', 'HCPCS_Code': 'E0236', 'Description': 'Pump for water circulating pad', 'Quantity': 10, 'Est_Unit_Cost': 45},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM08', 'HCPCS_Code': 'E0239', 'Description': 'Hydrocollator unit - portable', 'Quantity': 15, 'Est_Unit_Cost': 55},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM08', 'HCPCS_Code': 'E0235', 'Description': 'Paraffin bath unit', 'Quantity': 1, 'Est_Unit_Cost': 185},
    # Infrared
    {'Tier': 'TIER 6', 'BOC_Category': 'DM11', 'HCPCS_Code': 'E0221', 'Description': 'Infrared heating pad system', 'Quantity': 3, 'Est_Unit_Cost': 175},
    {'Tier': 'TIER 6', 'BOC_Category': 'DM11', 'HCPCS_Code': 'A4639', 'Description': 'Infrared pad replacement', 'Quantity': 10, 'Est_Unit_Cost': 25},
]

# Missing from TIER 7 (expand orthotics, add specialized equipment)
tier7_missing = [
    # Expand orthotics (beyond just ankle)
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L1810', 'Description': 'Knee orthosis - elastic support', 'Quantity': 4, 'Est_Unit_Cost': 45},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L1820', 'Description': 'Knee orthosis - soft interface', 'Quantity': 4, 'Est_Unit_Cost': 65},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L1830', 'Description': 'Knee orthosis - rigid support', 'Quantity': 4, 'Est_Unit_Cost': 125},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L0621', 'Description': 'Lumbar orthosis - sagittal control', 'Quantity': 3, 'Est_Unit_Cost': 85},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L0625', 'Description': 'Lumbar orthosis - sagittal-coronal control', 'Quantity': 3, 'Est_Unit_Cost': 95},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L3908', 'Description': 'Wrist-hand-finger orthosis', 'Quantity': 4, 'Est_Unit_Cost': 45},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L3916', 'Description': 'Wrist-hand orthosis - includes fingers', 'Quantity': 4, 'Est_Unit_Cost': 55},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L3702', 'Description': 'Elbow orthosis - without joints', 'Quantity': 3, 'Est_Unit_Cost': 40},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L0120', 'Description': 'Cervical collar - soft foam', 'Quantity': 2, 'Est_Unit_Cost': 25},
    {'Tier': 'TIER 7', 'BOC_Category': 'OR03', 'HCPCS_Code': 'L0172', 'Description': 'Cervical collar - semi-rigid', 'Quantity': 2, 'Est_Unit_Cost': 45},
    # External Infusion Pumps
    {'Tier': 'TIER 7', 'BOC_Category': 'DM12', 'HCPCS_Code': 'E0781', 'Description': 'Ambulatory infusion pump', 'Quantity': 2, 'Est_Unit_Cost': 450},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM12', 'HCPCS_Code': 'E0779', 'Description': 'Stationary infusion pump', 'Quantity': 1, 'Est_Unit_Cost': 650},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM12', 'HCPCS_Code': 'A4305', 'Description': 'Infusion pump supply - drug cassette', 'Quantity': 10, 'Est_Unit_Cost': 25},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM12', 'HCPCS_Code': 'A4306', 'Description': 'Infusion pump supply - administration set', 'Quantity': 10, 'Est_Unit_Cost': 15},
    # NPWT
    {'Tier': 'TIER 7', 'BOC_Category': 'DM15', 'HCPCS_Code': 'E2402', 'Description': 'Negative pressure wound therapy pump', 'Quantity': 1, 'Est_Unit_Cost': 850},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM15', 'HCPCS_Code': 'A7000', 'Description': 'NPWT canister - disposable', 'Quantity': 10, 'Est_Unit_Cost': 35},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM15', 'HCPCS_Code': 'A6550', 'Description': 'NPWT dressing kit', 'Quantity': 10, 'Est_Unit_Cost': 45},
    # Osteogenesis Stimulators
    {'Tier': 'TIER 7', 'BOC_Category': 'DM17', 'HCPCS_Code': 'E0760', 'Description': 'Osteogenesis stimulator - ultrasonic', 'Quantity': 1, 'Est_Unit_Cost': 550},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM17', 'HCPCS_Code': 'E0747', 'Description': 'Osteogenesis stimulator - electromagnetic', 'Quantity': 1, 'Est_Unit_Cost': 450},
    # Traction
    {'Tier': 'TIER 7', 'BOC_Category': 'DM21', 'HCPCS_Code': 'E0855', 'Description': 'Cervical traction equipment', 'Quantity': 2, 'Est_Unit_Cost': 185},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM21', 'HCPCS_Code': 'E0830', 'Description': 'Ambulatory traction device', 'Quantity': 1, 'Est_Unit_Cost': 125},
    {'Tier': 'TIER 7', 'BOC_Category': 'DM21', 'HCPCS_Code': 'E0890', 'Description': 'Pelvic traction setup', 'Quantity': 1, 'Est_Unit_Cost': 250},
]

# Missing from TIER 8
tier8_missing = [
    {'Tier': 'TIER 8', 'BOC_Category': 'PE03', 'HCPCS_Code': 'B4160', 'Description': 'Enteral formula - pediatric', 'Quantity': 10, 'Est_Unit_Cost': 35},
    {'Tier': 'TIER 8', 'BOC_Category': 'PE03', 'HCPCS_Code': 'B4161', 'Description': 'Enteral formula - specialized infant', 'Quantity': 5, 'Est_Unit_Cost': 45},
    {'Tier': 'TIER 8', 'BOC_Category': 'PE04', 'HCPCS_Code': 'E0776', 'Description': 'Enteral feeding pump', 'Quantity': 2, 'Est_Unit_Cost': 350},
    {'Tier': 'TIER 8', 'BOC_Category': 'PE04', 'HCPCS_Code': 'B4034', 'Description': 'Enteral feeding bag', 'Quantity': 15, 'Est_Unit_Cost': 8},
    {'Tier': 'TIER 8', 'BOC_Category': 'PE04', 'HCPCS_Code': 'B4035', 'Description': 'Enteral feeding bag - gravity', 'Quantity': 10, 'Est_Unit_Cost': 6},
    {'Tier': 'TIER 8', 'BOC_Category': 'PE04', 'HCPCS_Code': 'B4088', 'Description': 'Enteral extension set', 'Quantity': 15, 'Est_Unit_Cost': 12},
    {'Tier': 'TIER 8', 'BOC_Category': 'PE04', 'HCPCS_Code': 'B4081', 'Description': 'Enteral feeding tube - nasogastric', 'Quantity': 10, 'Est_Unit_Cost': 15},
]

# Missing from TIER 9
tier9_missing = [
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A7520', 'Description': 'Tracheostomy tube - cuffed', 'Quantity': 5, 'Est_Unit_Cost': 35},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A7521', 'Description': 'Tracheostomy tube - uncuffed', 'Quantity': 5, 'Est_Unit_Cost': 30},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A4625', 'Description': 'Tracheostomy care kit', 'Quantity': 10, 'Est_Unit_Cost': 18},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A4626', 'Description': 'Tracheostomy cleaning brush', 'Quantity': 10, 'Est_Unit_Cost': 5},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A7505', 'Description': 'Trach speaking valve - one way', 'Quantity': 3, 'Est_Unit_Cost': 85},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A7506', 'Description': 'Trach speaking valve - one way with adapter', 'Quantity': 2, 'Est_Unit_Cost': 95},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A4481', 'Description': 'Tracheostomy filter', 'Quantity': 15, 'Est_Unit_Cost': 8},
    {'Tier': 'TIER 9', 'BOC_Category': 'PD08', 'HCPCS_Code': 'A4483', 'Description': 'Tracheostomy moisture exchanger', 'Quantity': 15, 'Est_Unit_Cost': 12},
]

# Combine all missing items
all_missing = []
for tier_list in [tier1_missing, tier3_missing, tier4_missing, tier5_missing, tier6_missing, tier7_missing, tier8_missing, tier9_missing]:
    all_missing.extend(tier_list)

missing_df = pd.DataFrame(all_missing)

# Calculate costs and add metadata
missing_df['Total_Cost'] = missing_df['Quantity'] * missing_df['Est_Unit_Cost']
missing_df['Source'] = 'LAUNCH_INVENTORY'
missing_df['Customers'] = None
missing_df['Priority_Score'] = 40  # Launch inventory default priority
missing_df['Medicare_Rate'] = None  # Will be filled by lookup later

# Rename columns to match (don't create duplicate)
missing_df = missing_df.rename(columns={'Est_Unit_Cost': 'Estimated_Unit_Cost'})
missing_df = missing_df[['BOC_Category', 'HCPCS_Code', 'Description', 'Quantity', 'Estimated_Unit_Cost', 'Total_Cost', 'Medicare_Rate', 'Priority_Score', 'Source', 'Customers']]

print(f"  Missing items defined: {len(missing_df)} products")
print(f"    TIER 1 (DM02 Commodes): {len(tier1_missing)} items")
print(f"    TIER 3 (DM29 Suction): {len(tier3_missing)} items")
print(f"    TIER 4 (DM25 Insulin Pumps): {len(tier4_missing)} items")
print(f"    TIER 5 (R07, DM22, DM16, DM18): {len(tier5_missing)} items")
print(f"    TIER 6 (DM20, DM08, DM11): {len(tier6_missing)} items")
print(f"    TIER 7 (OR03, DM12, DM15, DM17, DM21): {len(tier7_missing)} items")
print(f"    TIER 8 (PE03, PE04): {len(tier8_missing)} items")
print(f"    TIER 9 (PD08): {len(tier9_missing)} items")

missing_cost = missing_df['Total_Cost'].sum()
print(f"\n  Additional Budget Required: ${missing_cost:,.2f}")

# ============================================================================
# STEP 3: COMBINE ALL INVENTORY
# ============================================================================
print("\nSTEP 3: Combining all inventory...")

# Reset indices to avoid duplicate index issues
current_with_added = current_with_added.reset_index(drop=True)
missing_df = missing_df.reset_index(drop=True)

# Debug: Check columns
print(f"\n  DEBUG: Current columns: {list(current_with_added.columns)}")
print(f"  DEBUG: Missing columns: {list(missing_df.columns)}")

# Ensure all columns are present in both DataFrames (fill missing ones with None)
all_columns = list(set(current_with_added.columns) | set(missing_df.columns))
for col in all_columns:
    if col not in current_with_added.columns:
        current_with_added[col] = None
    if col not in missing_df.columns:
        missing_df[col] = None

# Reorder columns to match
missing_df = missing_df[current_with_added.columns]

# Combine current + missing
full_inventory = pd.concat([current_with_added, missing_df], ignore_index=True)

print(f"  Current (with recent additions): {len(current_with_added)} products")
print(f"  Missing from original tiers: {len(missing_df)} products")
print(f"  TOTAL COMPREHENSIVE: {len(full_inventory)} products")

current_cost = current_with_added['Total_Cost'].sum()
new_total_cost = full_inventory['Total_Cost'].sum()

print(f"\n  Budget Summary:")
print(f"    Previous total: ${current_cost:,.2f}")
print(f"    Adding: ${missing_cost:,.2f}")
print(f"    NEW TOTAL: ${new_total_cost:,.2f}")

# ============================================================================
# STEP 4: LOAD MEDICARE RATES AND MATCH
# ============================================================================
print("\nSTEP 4: Matching Medicare rates...")

medicare = pd.read_excel(
    "/home/user/InitialInventoryScratch/Medicare_Rates_Normalized_Structure_Validated.xlsx",
    sheet_name="Medicare Rates - Normalized"
)

# Create lookup dictionaries
medicare_nu_urban = medicare[
    (medicare['Modifier 1'] == 'NU') &
    (medicare['Geographic Tier'] == 'Urban') &
    (medicare['Delivery Method'] == 'Standard')
].set_index('HCPCS Code')['Rate ($)'].to_dict()

medicare_nomod_urban = medicare[
    (medicare['Modifier 1'].isna()) &
    (medicare['Geographic Tier'] == 'Urban') &
    (medicare['Delivery Method'] == 'Standard')
].set_index('HCPCS Code')['Rate ($)'].to_dict()

def get_medicare_rate(hcpcs_code):
    if pd.isna(hcpcs_code):
        return None
    if hcpcs_code in medicare_nu_urban:
        return medicare_nu_urban[hcpcs_code]
    elif hcpcs_code in medicare_nomod_urban:
        return medicare_nomod_urban[hcpcs_code]
    else:
        return None

# Fill in Medicare rates for new items
for idx, row in full_inventory.iterrows():
    if pd.isna(row.get('Medicare_Rate')):
        rate = get_medicare_rate(row['HCPCS_Code'])
        full_inventory.at[idx, 'Medicare_Rate'] = rate

matched = full_inventory['Medicare_Rate'].notna().sum()
total = len(full_inventory)
print(f"  Medicare rates matched: {matched}/{total} ({matched/total*100:.1f}%)")

# ============================================================================
# STEP 5: CALCULATE PROFIT METRICS
# ============================================================================
print("\nSTEP 5: Calculating profit metrics...")

full_inventory['Profit_Margin_%'] = full_inventory.apply(
    lambda row: ((row['Medicare_Rate'] - row['Estimated_Unit_Cost']) / row['Medicare_Rate'])
    if pd.notna(row['Medicare_Rate']) and row['Medicare_Rate'] > 0 else None,
    axis=1
)

full_inventory['Profit_Per_Unit'] = full_inventory.apply(
    lambda row: row['Medicare_Rate'] - row['Estimated_Unit_Cost']
    if pd.notna(row['Medicare_Rate']) else None,
    axis=1
)

full_inventory['Total_Revenue_Potential'] = full_inventory.apply(
    lambda row: row['Quantity'] * row['Medicare_Rate']
    if pd.notna(row['Medicare_Rate']) else 0,
    axis=1
)

full_inventory['Total_Profit_Potential'] = full_inventory.apply(
    lambda row: row['Quantity'] * row['Profit_Per_Unit']
    if pd.notna(row['Profit_Per_Unit']) else 0,
    axis=1
)

# ============================================================================
# STEP 6: SAVE UPDATED MASTER PLAN
# ============================================================================
print("\nSTEP 6: Saving updated master inventory plan...")

output_csv = "/home/user/InitialInventoryScratch/MASTER_INVENTORY_PLAN_COMPLETE_ALL_TIERS.csv"
full_inventory.to_csv(output_csv, index=False)
print(f"  ✓ Saved to: {output_csv}")

# ============================================================================
# STEP 7: SUMMARY BY BOC CATEGORY
# ============================================================================
print("\n" + "="*100)
print("BOC CATEGORY COVERAGE SUMMARY")
print("="*100)

boc_summary = full_inventory.groupby('BOC_Category').agg({
    'HCPCS_Code': 'count',
    'Quantity': 'sum',
    'Total_Cost': 'sum'
}).rename(columns={'HCPCS_Code': 'SKU_Count', 'Quantity': 'Total_Units', 'Total_Cost': 'Total_Investment'})

boc_summary = boc_summary.sort_values('Total_Investment', ascending=False)

print(f"\nCategories Now Used: {len(boc_summary)}")
print("\nBreakdown:")
for boc, row in boc_summary.iterrows():
    print(f"  {boc:6s} - {int(row['SKU_Count']):3d} SKUs, {int(row['Total_Units']):5d} units, ${row['Total_Investment']:,.2f}")

print(f"\n{'='*100}")
print(f"FINAL TOTALS")
print("="*100)
print(f"  Total Products: {len(full_inventory)}")
print(f"  Total BOC Categories: {len(boc_summary)}/34 ({len(boc_summary)/34*100:.1f}%)")
print(f"  Total Investment: ${new_total_cost:,.2f}")
print(f"  Total Revenue Potential: ${full_inventory['Total_Revenue_Potential'].sum():,.2f}")
print(f"  Total Profit Potential: ${full_inventory['Total_Profit_Potential'].sum():,.2f}")
print("="*100)

print("\n✅ COMPLETE: All original Launch Inventory tiers added back")
print(f"   Ready to rebuild workbook with {len(full_inventory)} products across {len(boc_summary)} BOC categories")

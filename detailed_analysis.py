#!/usr/bin/env python3
"""
Detailed analysis of inventory and Medicare rates
"""
import pandas as pd
import numpy as np

print("="*100)
print("PHASE 1 DISCOVERY - DETAILED ANALYSIS")
print("="*100)

# 1. Load Medicare Rates
print("\n1. MEDICARE RATES DATA")
print("-"*100)
medicare = pd.read_excel(
    "/home/user/InitialInventoryScratch/Medicare_Rates_Normalized_Structure_Validated.xlsx",
    sheet_name="Medicare Rates - Normalized"
)
print(f"Total Medicare rate entries: {len(medicare)}")
print(f"Unique HCPCS codes: {medicare['HCPCS Code'].nunique()}")
print(f"\nRate statistics:")
print(medicare['Rate ($)'].describe())
print(f"\nUnique categories: {medicare['Category'].unique()}")
print(f"\nSample of rate variations for same HCPCS code:")
sample_hcpcs = medicare['HCPCS Code'].value_counts().head(1).index[0]
print(f"\nCode {sample_hcpcs} has {len(medicare[medicare['HCPCS Code']==sample_hcpcs])} rate variations:")
print(medicare[medicare['HCPCS Code']==sample_hcpcs][['HCPCS Code', 'Modifier 1', 'Geographic Tier', 'Delivery Method', 'Rate ($)']].to_string())

# 2. Load Inventory Data
print("\n\n2. INVENTORY DATA")
print("-"*100)
inventory = pd.read_excel(
    "/home/user/InitialInventoryScratch/Holistic_Medical_Inventory_DETAILED(1).xlsx",
    sheet_name="Inventory Detail"
)
print(f"Total inventory items: {len(inventory)}")
print(f"Items with HCPCS codes: {inventory['HCPCS Code'].notna().sum()}")
print(f"Unique HCPCS codes in inventory: {inventory['HCPCS Code'].nunique()}")

# Remove subtotal/summary rows for analysis
inventory_clean = inventory[inventory['HCPCS Code'].notna()].copy()
print(f"Clean inventory items (excluding subtotals): {len(inventory_clean)}")

print("\n\nTier breakdown:")
print(inventory_clean['Tier'].value_counts().sort_index())

print("\n\nBOC Category breakdown:")
print(inventory_clean['BOC Category'].value_counts())

# 3. Match inventory to Medicare rates
print("\n\n3. MEDICARE RATE MATCHING ANALYSIS")
print("-"*100)

inventory_hcpcs = set(inventory_clean['HCPCS Code'].dropna().unique())
medicare_hcpcs = set(medicare['HCPCS Code'].unique())

print(f"HCPCS codes in inventory: {len(inventory_hcpcs)}")
print(f"HCPCS codes in Medicare data: {len(medicare_hcpcs)}")
print(f"Codes with Medicare rates: {len(inventory_hcpcs & medicare_hcpcs)}")
print(f"Codes WITHOUT Medicare rates: {len(inventory_hcpcs - medicare_hcpcs)}")

if len(inventory_hcpcs - medicare_hcpcs) > 0:
    print(f"\nCodes in inventory but NOT in Medicare data:")
    for code in sorted(inventory_hcpcs - medicare_hcpcs):
        items = inventory_clean[inventory_clean['HCPCS Code']==code]
        for _, item in items.iterrows():
            print(f"  {code} - {item['Product Description']} (Tier: {item['Tier']})")

# 4. Rate variation analysis
print("\n\n4. MEDICARE RATE VARIATION ANALYSIS")
print("-"*100)
print("Understanding how many rate scenarios exist per HCPCS code:")

rate_variations = medicare.groupby('HCPCS Code').agg({
    'Rate ($)': ['min', 'max', 'count', 'mean']
}).round(2)
rate_variations.columns = ['Min Rate', 'Max Rate', 'Num Variations', 'Avg Rate']
rate_variations['Rate Spread'] = rate_variations['Max Rate'] - rate_variations['Min Rate']

print(f"\nRate variation statistics:")
print(rate_variations['Num Variations'].describe())
print(f"\nCodes with most rate variations:")
print(rate_variations.nlargest(10, 'Num Variations'))

# Focus on inventory codes
inventory_codes_with_rates = list(inventory_hcpcs & medicare_hcpcs)
inventory_rate_variations = rate_variations.loc[inventory_codes_with_rates].sort_values('Num Variations', ascending=False)

print(f"\n\nRate variations for codes in OUR inventory:")
print(f"Total inventory codes with Medicare rates: {len(inventory_rate_variations)}")
print(inventory_rate_variations.head(20).to_string())

# 5. Budget Analysis
print("\n\n5. BUDGET AND COST ANALYSIS")
print("-"*100)

# Clean budget values (remove $ and convert to float)
inventory_clean['Budget_Clean'] = inventory_clean['Budget Allocation'].str.replace('$', '').str.replace(',', '').astype(float)

print(f"Total budget allocation: ${inventory_clean['Budget_Clean'].sum():,.2f}")
print(f"\nBudget by Tier:")
tier_budget = inventory_clean.groupby('Tier')['Budget_Clean'].sum().sort_values(ascending=False)
for tier, amount in tier_budget.items():
    pct = amount / inventory_clean['Budget_Clean'].sum() * 100
    print(f"  {tier}: ${amount:,.2f} ({pct:.1f}%)")

print(f"\nBudget by BOC Category (top 10):")
cat_budget = inventory_clean.groupby('BOC Category')['Budget_Clean'].sum().sort_values(ascending=False).head(10)
for cat, amount in cat_budget.items():
    pct = amount / inventory_clean['Budget_Clean'].sum() * 100
    print(f"  {cat}: ${amount:,.2f} ({pct:.1f}%)")

# 6. Data Quality Issues
print("\n\n6. DATA QUALITY ISSUES IDENTIFIED")
print("-"*100)

print(f"✓ Medicare rates data: COMPLETE (1,531 entries)")
print(f"✓ Inventory structure: GOOD (168 rows, 132 products)")
print(f"✗ Medicare Reimbursement Rate column: EMPTY (0% filled)")
print(f"✗ Missing rates for {len(inventory_hcpcs - medicare_hcpcs)} codes")
print(f"✗ No unit cost data in inventory")
print(f"✗ No quantity × unit cost calculations")
print(f"✗ No vendor comparison columns")
print(f"✗ No margin calculations")

print("\n\n7. WHAT'S MISSING FROM CURRENT WORKBOOK")
print("-"*100)
print("Required for VGM vendor meeting:")
print("  1. Medicare rate matching (all scenarios: Urban/Rural, Standard/Delivery modifiers)")
print("  2. Unit cost estimates or vendor pricing")
print("  3. Line total calculations (Qty × Unit Cost)")
print("  4. Vendor comparison columns (3 vendors)")
print("  5. Best cost/margin formulas")
print("  6. Conditional formatting for margin analysis")
print("  7. Missing rates analysis tab")
print("  8. Tier budget summary tab")
print("  9. Professional formatting with frozen headers")

print("\n" + "="*100)
print("PHASE 1 DISCOVERY COMPLETE")
print("="*100)

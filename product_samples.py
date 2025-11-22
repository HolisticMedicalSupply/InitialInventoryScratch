#!/usr/bin/env python3
"""
Get sample products to understand data structure
"""
import pandas as pd

inventory = pd.read_excel(
    "/home/user/InitialInventoryScratch/Holistic_Medical_Inventory_DETAILED(1).xlsx",
    sheet_name="Inventory Detail"
)

medicare = pd.read_excel(
    "/home/user/InitialInventoryScratch/Medicare_Rates_Normalized_Structure_Validated.xlsx",
    sheet_name="Medicare Rates - Normalized"
)

print("SAMPLE PRODUCTS WITH MEDICARE RATE VARIATIONS\n")
print("="*120)

# Clean inventory
inventory_clean = inventory[inventory['HCPCS Code'].notna()].copy()

# Sample: Wheelchair (likely has multiple rates)
print("\nEXAMPLE 1: WHEELCHAIR K0001")
print("-"*120)
wheelchair = inventory_clean[inventory_clean['HCPCS Code']=='K0001'].iloc[0]
print(f"Product: {wheelchair['Product Description']}")
print(f"BOC Category: {wheelchair['BOC Category']}")
print(f"Quantity: {wheelchair['Quantity']}")
print(f"Budget: {wheelchair['Budget Allocation']}")
print(f"\nMedicare rates for K0001:")
k0001_rates = medicare[medicare['HCPCS Code']=='K0001'][['HCPCS Code', 'Modifier 1', 'Geographic Tier', 'Delivery Method', 'Rate ($)', 'Description']]
print(k0001_rates.to_string(index=False))

print("\n\nEXAMPLE 2: WALKER E0143")
print("-"*120)
walker = inventory_clean[inventory_clean['HCPCS Code']=='E0143'].iloc[0]
print(f"Product: {walker['Product Description']}")
print(f"Quantity: {walker['Quantity']}")
print(f"Budget: {walker['Budget Allocation']}")
print(f"\nMedicare rates for E0143:")
e0143_rates = medicare[medicare['HCPCS Code']=='E0143'][['HCPCS Code', 'Modifier 1', 'Geographic Tier', 'Delivery Method', 'Rate ($)', 'Description']]
print(e0143_rates.to_string(index=False))

print("\n\nEXAMPLE 3: WOUND DRESSING A6209")
print("-"*120)
foam = inventory_clean[inventory_clean['HCPCS Code']=='A6209'].iloc[0]
print(f"Product: {foam['Product Description']}")
print(f"Quantity: {foam['Quantity']}")
print(f"Budget: {foam['Budget Allocation']}")
print(f"\nMedicare rates for A6209:")
a6209_rates = medicare[medicare['HCPCS Code']=='A6209'][['HCPCS Code', 'Modifier 1', 'Geographic Tier', 'Delivery Method', 'Rate ($)', 'Description']]
print(a6209_rates.to_string(index=False))

print("\n\nEXAMPLE 4: GLUCOSE MONITOR E0607")
print("-"*120)
glucose = inventory_clean[inventory_clean['HCPCS Code']=='E0607'].iloc[0]
print(f"Product: {glucose['Product Description']}")
print(f"Quantity: {glucose['Quantity']}")
print(f"Budget: {glucose['Budget Allocation']}")
print(f"\nMedicare rates for E0607:")
e0607_rates = medicare[medicare['HCPCS Code']=='E0607'][['HCPCS Code', 'Modifier 1', 'Geographic Tier', 'Delivery Method', 'Rate ($)', 'Description']]
print(e0607_rates.to_string(index=False))

print("\n\nEXAMPLE 5: CODE WITH NO MEDICARE RATE - A4520 (Adult Diaper)")
print("-"*120)
diaper = inventory_clean[inventory_clean['HCPCS Code']=='A4520'].iloc[0]
print(f"Product: {diaper['Product Description']}")
print(f"Quantity: {diaper['Quantity']}")
print(f"Budget: {diaper['Budget Allocation']}")
print(f"\nMedicare rates for A4520:")
a4520_rates = medicare[medicare['HCPCS Code']=='A4520']
if len(a4520_rates) == 0:
    print("  ❌ NO MEDICARE RATES FOUND - This code is not covered by Medicare Part B")
else:
    print(a4520_rates[['HCPCS Code', 'Modifier 1', 'Geographic Tier', 'Rate ($)', 'Description']].to_string(index=False))

print("\n\n" + "="*120)

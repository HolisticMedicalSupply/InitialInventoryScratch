#!/usr/bin/env python3
"""
Master Inventory Plan Builder
Consolidates requirements from:
1. Customer Needs
2. Launch Inventory structure
3. DME Catalog priorities
4. BOC Approved codes (978 codes)
5. Medicare rates database

Goal: Maximize SKU count within $50-60K budget
"""

import pandas as pd
import numpy as np
from datetime import datetime

print("=" * 80)
print("MASTER INVENTORY PLAN BUILDER")
print("=" * 80)
print()

# ============================================================================
# STEP 1: Load all data sources
# ============================================================================

print("Step 1: Loading data sources...")
print("-" * 80)

# Load approved HCPCS codes (skip to detailed section starting at line 41)
print("Loading approved codes...")
approved_codes = pd.read_csv('ApprovedCategoriesAndCodes copy.csv', skiprows=41)
approved_codes.columns = ['BOC_Category', 'BOC_Description', 'HCPCS_Code', 'Description', 'Requires_Accreditation', 'Effective_Date']
print(f"  ✓ Loaded {len(approved_codes)} approved HCPCS codes")
print(f"  ✓ Unique BOC categories: {approved_codes['BOC_Category'].nunique()}")

# Load Medicare rates
print("Loading Medicare rates...")
medicare_rates = pd.read_excel('Medicare_Rates_Normalized_Structure_Validated.xlsx',
                               sheet_name='Medicare Rates - Normalized')
# Rename columns to match expected format
medicare_rates.columns = medicare_rates.columns.str.replace(' ($)', '', regex=False).str.replace(' ', '_')
print(f"  ✓ Loaded {len(medicare_rates)} Medicare rate entries")
print(f"  ✓ Unique HCPCS codes with rates: {medicare_rates['HCPCS_Code'].nunique()}")

# Create Medicare rate lookup (use NU/Urban as primary rate)
medicare_primary = medicare_rates[
    (medicare_rates['Modifier_1'] == 'NU') &
    (medicare_rates['Geographic_Tier'] == 'Urban')
].copy()

# For codes without NU+Urban, use first available rate
all_codes_with_rates = medicare_rates[['HCPCS_Code', 'Rate']].drop_duplicates('HCPCS_Code')
medicare_lookup = pd.concat([
    medicare_primary[['HCPCS_Code', 'Rate']],
    all_codes_with_rates[~all_codes_with_rates['HCPCS_Code'].isin(medicare_primary['HCPCS_Code'])]
]).drop_duplicates('HCPCS_Code')

print(f"  ✓ Created Medicare lookup for {len(medicare_lookup)} codes")
print()

# ============================================================================
# STEP 2: Define customer-specific high-priority items
# ============================================================================

print("Step 2: Defining customer-specific priorities...")
print("-" * 80)

customer_priorities = {
    'MOYHINOR': {
        'K0001': {'qty': 2, 'budget': 800, 'note': 'Kids wheelchair'},
        'E0607': {'qty': 10, 'budget': 1000, 'note': 'Glucose monitors'},
        'A5500': {'qty': 6, 'budget': 900, 'note': 'Diabetic shoes'},
    },
    'RAMBOM': {
        'B4150': {'qty': 40, 'budget': 1200, 'note': 'Enteral formula'},
        'E0130': {'qty': 5, 'budget': 500, 'note': 'Rollators'},
        'E0143': {'qty': 3, 'budget': 600, 'note': 'Wheelchairs'},
        'E0570': {'qty': 6, 'budget': 900, 'note': 'Nebulizers'},
    },
    'DR_NAS': {
        'L4360': {'qty': 15, 'budget': 1500, 'note': 'CAM walker boots'},
        'L1902': {'qty': 20, 'budget': 800, 'note': 'Ankle braces'},
        'A5500': {'qty': 10, 'budget': 600, 'note': 'Post-op shoes'},
    },
    'WALTERS': {
        'E0240': {'qty': 20, 'budget': 2000, 'note': 'Shower chairs'},
        'A4520': {'qty': 3000, 'budget': 3500, 'note': 'Adult diapers (non-Medicare)'},
        'A4554': {'qty': 2000, 'budget': 250, 'note': 'Chux/underpads (non-Medicare)'},
        'E0607': {'qty': 15, 'budget': 1500, 'note': 'Glucose monitors'},
        'A4253': {'qty': 1500, 'budget': 750, 'note': 'Glucose test strips'},
        'A4259': {'qty': 1000, 'budget': 200, 'note': 'Lancets'},
    }
}

# Flatten customer priorities
customer_codes = {}
for customer, items in customer_priorities.items():
    for code, data in items.items():
        if code not in customer_codes:
            customer_codes[code] = {'qty': 0, 'budget': 0, 'customers': []}
        customer_codes[code]['qty'] += data['qty']
        customer_codes[code]['budget'] += data['budget']
        customer_codes[code]['customers'].append(customer)

print(f"  ✓ Identified {len(customer_codes)} unique HCPCS codes from customer requests")
print(f"  ✓ Total customer budget: ${sum(d['budget'] for d in customer_codes.values()):,.0f}")
print()

# ============================================================================
# STEP 3: Define Launch Inventory structure priorities
# ============================================================================

print("Step 3: Defining Launch Inventory priorities...")
print("-" * 80)

# Core product categories from Launch Inventory with budget allocation
launch_inventory_structure = {
    'TIER_1_MOBILITY': {
        'budget': 11500,
        'categories': ['M01', 'M05', 'M06', 'M06A'],
        'priority': 1,
    },
    'TIER_2_WOUND_CARE': {
        'budget': 10750,
        'categories': ['S01', 'S04'],
        'priority': 1,
    },
    'TIER_3_INCONTINENCE': {
        'budget': 8250,
        'categories': ['PD09'],
        'priority': 1,
    },
    'TIER_4_DIABETES': {
        'budget': 10300,
        'categories': ['DM05', 'DM06', 'S02'],
        'priority': 1,
    },
    'TIER_5_RESPIRATORY': {
        'budget': 8600,
        'categories': ['R07', 'DM11', 'DM12'],
        'priority': 2,
    },
    'TIER_6_PRESSURE_RELIEF': {
        'budget': 6800,
        'categories': ['DM20', 'DM26'],
        'priority': 2,
    },
    'TIER_7_ORTHOTICS': {
        'budget': 6000,
        'categories': ['OR03'],
        'priority': 1,
    },
    'TIER_8_ENTERAL': {
        'budget': 3500,
        'categories': ['PE03', 'PE04'],
        'priority': 2,
    },
    'TIER_9_MISC': {
        'budget': 2050,
        'categories': ['DM02', 'DM13'],
        'priority': 3,
    },
}

total_launch_budget = sum(tier['budget'] for tier in launch_inventory_structure.values())
print(f"  ✓ Launch inventory structure: 9 tiers")
print(f"  ✓ Total tier budget: ${total_launch_budget:,.0f}")
print()

# ============================================================================
# STEP 4: Merge approved codes with Medicare rates
# ============================================================================

print("Step 4: Merging approved codes with Medicare rates...")
print("-" * 80)

# Merge approved codes with Medicare rates
master_list = approved_codes.merge(
    medicare_lookup,
    left_on='HCPCS_Code',
    right_on='HCPCS_Code',
    how='left'
)

print(f"  ✓ Master list: {len(master_list)} codes")
print(f"  ✓ Codes with Medicare rates: {master_list['Rate'].notna().sum()} ({master_list['Rate'].notna().sum()/len(master_list)*100:.1f}%)")
print(f"  ✓ Codes without Medicare rates: {master_list['Rate'].isna().sum()} ({master_list['Rate'].isna().sum()/len(master_list)*100:.1f}%)")
print()

# ============================================================================
# STEP 5: Calculate priority scores
# ============================================================================

print("Step 5: Calculating priority scores...")
print("-" * 80)

def calculate_priority_score(row, customer_codes, launch_structure):
    """
    Calculate priority score (0-100) based on:
    - Customer request: 40 points
    - Launch inventory tier: 30 points
    - Medicare rate availability: 20 points
    - High-volume category: 10 points
    """
    score = 0
    code = row['HCPCS_Code']
    boc_cat = row['BOC_Category']

    # Customer request bonus (40 points max)
    if code in customer_codes:
        num_customers = len(customer_codes[code]['customers'])
        score += min(40, num_customers * 20)  # 20 pts per customer, max 40

    # Launch inventory tier priority (30 points max)
    for tier_name, tier_data in launch_structure.items():
        if boc_cat in tier_data['categories']:
            if tier_data['priority'] == 1:
                score += 30
            elif tier_data['priority'] == 2:
                score += 20
            elif tier_data['priority'] == 3:
                score += 10
            break

    # Medicare rate availability (20 points)
    if pd.notna(row['Rate']):
        score += 20

    # High-volume category bonus (10 points)
    # Categories with many codes = wider coverage
    high_volume_cats = ['S01', 'S04', 'OR03', 'PD09', 'M06A', 'M07A']
    if boc_cat in high_volume_cats:
        score += 10

    return score

master_list['Priority_Score'] = master_list.apply(
    lambda row: calculate_priority_score(row, customer_codes, launch_inventory_structure),
    axis=1
)

print(f"  ✓ Priority scores calculated")
print(f"  ✓ Score range: {master_list['Priority_Score'].min():.0f} - {master_list['Priority_Score'].max():.0f}")
print(f"  ✓ Mean score: {master_list['Priority_Score'].mean():.1f}")
print()

# ============================================================================
# STEP 6: Build inventory plan
# ============================================================================

print("Step 6: Building inventory plan...")
print("-" * 80)

# Sort by priority score
master_list_sorted = master_list.sort_values('Priority_Score', ascending=False).copy()

# Start with customer-requested items (guaranteed inclusion)
inventory_plan = []
running_budget = 0
budget_limit = 60000  # $60K max

print("\nPhase 1: Adding customer-requested items...")
for code, data in customer_codes.items():
    if code in master_list_sorted['HCPCS_Code'].values:
        row = master_list_sorted[master_list_sorted['HCPCS_Code'] == code].iloc[0]

        # Use customer budget and quantity
        item = {
            'HCPCS_Code': code,
            'BOC_Category': row['BOC_Category'],
            'Description': row['Description'],
            'Quantity': data['qty'],
            'Estimated_Unit_Cost': data['budget'] / data['qty'],
            'Total_Cost': data['budget'],
            'Medicare_Rate': row['Rate'] if pd.notna(row['Rate']) else None,
            'Priority_Score': row['Priority_Score'],
            'Source': 'CUSTOMER',
            'Customers': ', '.join(data['customers'])
        }

        inventory_plan.append(item)
        running_budget += data['budget']

print(f"  ✓ Added {len(inventory_plan)} customer-requested items")
print(f"  ✓ Running budget: ${running_budget:,.0f}")

# Phase 2: Add high-priority items from Launch Inventory structure
print("\nPhase 2: Adding Launch Inventory core items...")

# Focus on high-count categories for SKU diversity
for tier_name, tier_data in sorted(launch_inventory_structure.items(),
                                   key=lambda x: x[1]['priority']):
    tier_budget = tier_data['budget']
    tier_categories = tier_data['categories']

    # Get codes in this tier's categories, excluding already added
    tier_codes = master_list_sorted[
        (master_list_sorted['BOC_Category'].isin(tier_categories)) &
        (~master_list_sorted['HCPCS_Code'].isin([item['HCPCS_Code'] for item in inventory_plan]))
    ].copy()

    if len(tier_codes) == 0:
        continue

    # Distribute tier budget across codes (prefer more SKUs with lower quantities)
    # Strategy: Stock 1-3 units of each code for maximum variety
    tier_items_added = 0
    tier_budget_used = 0

    for idx, row in tier_codes.iterrows():
        if running_budget >= budget_limit:
            break

        # Estimate unit cost from Medicare rate
        if pd.notna(row['Rate']):
            # Assume we buy at 60% of Medicare rate (40% margin target)
            est_unit_cost = row['Rate'] * 0.60
        else:
            # For items without Medicare rates, use category averages
            category_avg_rates = master_list_sorted[
                (master_list_sorted['BOC_Category'] == row['BOC_Category']) &
                (master_list_sorted['Rate'].notna())
            ]['Rate'].mean()

            if pd.notna(category_avg_rates):
                est_unit_cost = category_avg_rates * 0.60
            else:
                est_unit_cost = 50  # Default $50 if no data

        # Determine quantity (1-3 units for diversity, more for high-priority)
        if row['Priority_Score'] >= 60:
            qty = 3
        elif row['Priority_Score'] >= 40:
            qty = 2
        else:
            qty = 1

        item_cost = est_unit_cost * qty

        # Check if we have budget
        if running_budget + item_cost <= budget_limit:
            item = {
                'HCPCS_Code': row['HCPCS_Code'],
                'BOC_Category': row['BOC_Category'],
                'Description': row['Description'],
                'Quantity': qty,
                'Estimated_Unit_Cost': est_unit_cost,
                'Total_Cost': item_cost,
                'Medicare_Rate': row['Rate'] if pd.notna(row['Rate']) else None,
                'Priority_Score': row['Priority_Score'],
                'Source': 'LAUNCH_INVENTORY',
                'Customers': ''
            }

            inventory_plan.append(item)
            running_budget += item_cost
            tier_budget_used += item_cost
            tier_items_added += 1

    print(f"  {tier_name}: Added {tier_items_added} items (${tier_budget_used:,.0f})")

print(f"\n  ✓ Total items after Launch Inventory: {len(inventory_plan)}")
print(f"  ✓ Running budget: ${running_budget:,.0f}")

# Phase 3: Fill remaining budget with high-SKU categories
print("\nPhase 3: Maximizing SKU count with remaining budget...")

remaining_budget = budget_limit - running_budget
high_sku_categories = ['S01', 'S04', 'OR03', 'PD09', 'M06A']  # Categories with many codes

items_added = 0
for cat in high_sku_categories:
    if running_budget >= budget_limit:
        break

    cat_codes = master_list_sorted[
        (master_list_sorted['BOC_Category'] == cat) &
        (master_list_sorted['Rate'].notna()) &  # Only items with Medicare rates
        (~master_list_sorted['HCPCS_Code'].isin([item['HCPCS_Code'] for item in inventory_plan]))
    ].copy()

    for idx, row in cat_codes.iterrows():
        if running_budget >= budget_limit:
            break

        # Stock 1 unit for diversity
        est_unit_cost = row['Rate'] * 0.60

        if running_budget + est_unit_cost <= budget_limit:
            item = {
                'HCPCS_Code': row['HCPCS_Code'],
                'BOC_Category': row['BOC_Category'],
                'Description': row['Description'],
                'Quantity': 1,
                'Estimated_Unit_Cost': est_unit_cost,
                'Total_Cost': est_unit_cost,
                'Medicare_Rate': row['Rate'],
                'Priority_Score': row['Priority_Score'],
                'Source': 'SKU_DIVERSITY',
                'Customers': ''
            }

            inventory_plan.append(item)
            running_budget += est_unit_cost
            items_added += 1

print(f"  ✓ Added {items_added} items for SKU diversity")
print(f"  ✓ Final running budget: ${running_budget:,.0f}")
print()

# ============================================================================
# STEP 7: Create final inventory DataFrame
# ============================================================================

print("Step 7: Finalizing inventory plan...")
print("-" * 80)

inventory_df = pd.DataFrame(inventory_plan)

# Calculate margins
inventory_df['Profit_Margin_%'] = inventory_df.apply(
    lambda row: ((row['Medicare_Rate'] - row['Estimated_Unit_Cost']) / row['Medicare_Rate'] * 100)
    if pd.notna(row['Medicare_Rate']) and row['Medicare_Rate'] > 0 else None,
    axis=1
)

inventory_df['Profit_Per_Unit'] = inventory_df.apply(
    lambda row: (row['Medicare_Rate'] - row['Estimated_Unit_Cost'])
    if pd.notna(row['Medicare_Rate']) else None,
    axis=1
)

inventory_df['Total_Revenue_Potential'] = inventory_df.apply(
    lambda row: row['Quantity'] * row['Medicare_Rate']
    if pd.notna(row['Medicare_Rate']) else None,
    axis=1
)

inventory_df['Total_Profit_Potential'] = inventory_df.apply(
    lambda row: row['Quantity'] * row['Profit_Per_Unit']
    if pd.notna(row['Profit_Per_Unit']) else None,
    axis=1
)

# Sort by priority score
inventory_df = inventory_df.sort_values('Priority_Score', ascending=False)

# Save to CSV
output_file = 'MASTER_INVENTORY_PLAN.csv'
inventory_df.to_csv(output_file, index=False)

print(f"  ✓ Saved master inventory plan to {output_file}")
print()

# ============================================================================
# STEP 8: Generate summary statistics
# ============================================================================

print("=" * 80)
print("MASTER INVENTORY PLAN SUMMARY")
print("=" * 80)
print()

print(f"Total SKU Count: {len(inventory_df)}")
print(f"Total Investment: ${inventory_df['Total_Cost'].sum():,.2f}")
print(f"Total Units: {inventory_df['Quantity'].sum():,.0f}")
print()

print("By Source:")
for source in inventory_df['Source'].unique():
    count = len(inventory_df[inventory_df['Source'] == source])
    budget = inventory_df[inventory_df['Source'] == source]['Total_Cost'].sum()
    print(f"  {source}: {count} SKUs (${budget:,.2f})")
print()

print("By BOC Category:")
cat_summary = inventory_df.groupby('BOC_Category').agg({
    'HCPCS_Code': 'count',
    'Total_Cost': 'sum',
    'Quantity': 'sum'
}).rename(columns={'HCPCS_Code': 'SKU_Count'}).sort_values('Total_Cost', ascending=False)

for cat, row in cat_summary.head(15).iterrows():
    print(f"  {cat}: {int(row['SKU_Count'])} SKUs, {int(row['Quantity'])} units (${row['Total_Cost']:,.0f})")
print()

print("Margin Analysis (items with Medicare rates):")
items_with_margins = inventory_df[inventory_df['Profit_Margin_%'].notna()]
if len(items_with_margins) > 0:
    print(f"  Items with Medicare rates: {len(items_with_margins)} ({len(items_with_margins)/len(inventory_df)*100:.1f}%)")
    print(f"  Average margin: {items_with_margins['Profit_Margin_%'].mean():.1f}%")
    print(f"  Median margin: {items_with_margins['Profit_Margin_%'].median():.1f}%")
    print(f"  Total revenue potential: ${items_with_margins['Total_Revenue_Potential'].sum():,.2f}")
    print(f"  Total profit potential: ${items_with_margins['Total_Profit_Potential'].sum():,.2f}")
print()

print("Top 20 Products by Priority Score:")
print(inventory_df[['HCPCS_Code', 'Description', 'BOC_Category', 'Quantity',
                    'Total_Cost', 'Priority_Score', 'Source']].head(20).to_string(index=False))
print()

print("=" * 80)
print("✓ Master inventory plan complete!")
print("=" * 80)

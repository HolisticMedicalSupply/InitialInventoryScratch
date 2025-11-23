#!/usr/bin/env python3
"""
Validate Complete Initial Inventory Plan Workbook
Checks against all requirements from Customer Needs and Launch Inventory PDFs
"""

import pandas as pd
import openpyxl

print("="*80)
print("VALIDATING COMPLETE INITIAL INVENTORY PLAN WORKBOOK")
print("="*80)

# Load the final CSV
df = pd.read_csv('MASTER_INVENTORY_PLAN_COMPLETE_FINAL.csv')

print(f"\n[1/6] Loading inventory data...")
print(f"   Total items loaded: {len(df)}")

# Critical items checklist
print("\n[2/6] Verifying critical customer-requested items...")

critical_items = {
    'L4361': {'name': 'CAM walker', 'customer': 'DR_NAS', 'required': True},
    'L1902': {'name': 'Ankle gauntlet', 'customer': 'DR_NAS', 'required': True},
    'L1906': {'name': 'Multiligamentous support', 'customer': 'DR_NAS', 'required': True},
    'L4350': {'name': 'Ankle control orthosis', 'customer': 'DR_NAS', 'required': True},
    'L4370': {'name': 'Pneumatic full leg splint', 'customer': 'DR_NAS', 'required': True},
    'L4387': {'name': 'Walking boot non-pneumatic', 'customer': 'DR_NAS', 'required': True},
    'E0607': {'name': 'Blood glucose monitor', 'customer': 'MOYHINOR/WALTERS', 'required': True},
    'E2103': {'name': 'CGM receiver (Dexcom/Libre)', 'customer': 'MOYHINOR (CGM)', 'required': True},
    'A4239': {'name': 'CGM sensors', 'customer': 'MOYHINOR (CGM)', 'required': True},
    'A4253': {'name': 'Glucose test strips', 'customer': 'WALTERS', 'required': True},
    'A4520': {'name': 'Adult diapers/briefs', 'customer': 'WALTERS', 'required': True},
    'A4554': {'name': 'Disposable underpads/chux', 'customer': 'WALTERS', 'required': True},
    'E0143': {'name': 'Rollator walker', 'customer': 'RAMBOM', 'required': True},
    'E0570': {'name': 'Nebulizer', 'customer': 'RAMBOM', 'required': True},
    'B4150': {'name': 'Enteral formula', 'customer': 'RAMBOM', 'required': True},
    'K0001': {'name': 'Standard wheelchair', 'customer': 'MOYHINOR/RAMBOM', 'required': True},
    'E0240': {'name': 'Shower chair', 'customer': 'WALTERS', 'required': True},
}

critical_found = 0
critical_missing = []

for hcpcs, info in critical_items.items():
    if hcpcs in df['HCPCS_Code'].values:
        critical_found += 1
        print(f"   ✅ {hcpcs} - {info['name']} ({info['customer']})")
    else:
        critical_missing.append(f"{hcpcs} - {info['name']} ({info['customer']})")
        print(f"   ❌ {hcpcs} - {info['name']} ({info['customer']}) - MISSING!")

print(f"\n   Customer Items Score: {critical_found}/{len(critical_items)} ({critical_found/len(critical_items)*100:.1f}%)")

# High-priority missing items from analysis
print("\n[3/6] Verifying previously-missing high-priority items...")

high_priority_additions = {
    'E2601': 'Wheelchair cushion <22"',
    'E2602': 'Wheelchair cushion ≥22"',
    'E2603': 'Skin protection cushion',
    'A4338': 'Foley catheter 2-way latex',
    'A4340': 'Foley catheter silicone',
    'A4357': 'Bedside drainage bag',
    'A4358': 'Urinary leg bag',
    'A4349': 'Male external catheter',
    'A4553': 'Non-disposable underpads',
    'A4256': 'Glucose control solution',
    'A4258': 'Lancing device',
    'A4557': 'TENS lead wires',
    'A7004': 'Disposable nebulizer',
    'A7002': 'Suction tubing',
    'A4600': 'Compression sleeve replacement',
    'A4559': 'Osteogenesis coupling gel',
}

high_priority_found = 0
high_priority_missing = []

for hcpcs, name in high_priority_additions.items():
    if hcpcs in df['HCPCS_Code'].values:
        high_priority_found += 1
        print(f"   ✅ {hcpcs} - {name}")
    else:
        high_priority_missing.append(f"{hcpcs} - {name}")
        print(f"   ❌ {hcpcs} - {name} - MISSING!")

print(f"\n   High Priority Additions Score: {high_priority_found}/{len(high_priority_additions)} ({high_priority_found/len(high_priority_additions)*100:.1f}%)")

# BOC Category Coverage
print("\n[4/6] Analyzing BOC category coverage...")

boc_counts = df['BOC_Category'].value_counts().sort_index()
print(f"\n   Total BOC Categories: {len(boc_counts)}")
print(f"\n   Category Distribution:")
for boc, count in boc_counts.items():
    print(f"      {boc}: {count} items")

# Tier Coverage
print("\n[5/6] Analyzing tier coverage...")

tier_mapping = {
    'M05': 'TIER 1', 'M01': 'TIER 1', 'M06': 'TIER 1', 'M06A': 'TIER 1', 'DM02': 'TIER 1',
    'S01': 'TIER 2', 'S04': 'TIER 2',
    'PD09': 'TIER 3', 'DM29': 'TIER 3',
    'DM05': 'TIER 4', 'DM06': 'TIER 4', 'DM25': 'TIER 4',
    'R07': 'TIER 5', 'DM22': 'TIER 5', 'DM16': 'TIER 5', 'DM18': 'TIER 5',
    'DM20': 'TIER 6', 'DM08': 'TIER 6', 'DM11': 'TIER 6',
    'OR03': 'TIER 7', 'DM12': 'TIER 7', 'DM15': 'TIER 7', 'DM17': 'TIER 7', 'DM21': 'TIER 7',
    'PE03': 'TIER 8', 'PE04': 'TIER 8',
    'PD08': 'TIER 9'
}

df['Tier'] = df['BOC_Category'].map(tier_mapping)
tier_counts = df[df['Tier'].notna()]['Tier'].value_counts().sort_index()

print(f"\n   Tier Distribution:")
for tier, count in tier_counts.items():
    tier_investment = df[df['Tier'] == tier]['Total_Cost'].sum()
    print(f"      {tier}: {count} items, ${tier_investment:,.2f} investment")

# Financial Analysis
print("\n[6/6] Financial summary...")

total_investment = df['Total_Cost'].sum()
total_revenue = df['Total_Revenue_Potential'].sum()
total_profit = df['Total_Profit_Potential'].sum()
avg_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
items_with_medicare = len(df[df['Medicare_Rate'] > 0])

print(f"""
   Financial Metrics:
   - Total Investment: ${total_investment:,.2f}
   - Total Revenue Potential: ${total_revenue:,.2f}
   - Total Profit Potential: ${total_profit:,.2f}
   - Average Margin: {avg_margin:.1f}%
   - Items with Medicare Coverage: {items_with_medicare}/{len(df)} ({items_with_medicare/len(df)*100:.1f}%)
""")

# Final Scores
print("\n" + "="*80)
print("VALIDATION SUMMARY")
print("="*80)

customer_score = critical_found / len(critical_items) * 100
high_priority_score = high_priority_found / len(high_priority_additions) * 100
overall_score = (customer_score + high_priority_score) / 2

print(f"✅ Customer-Requested Items: {critical_found}/{len(critical_items)} ({customer_score:.1f}%)")
print(f"✅ High-Priority Additions: {high_priority_found}/{len(high_priority_additions)} ({high_priority_score:.1f}%)")
print(f"✅ Total Items in Plan: {len(df)}")
print(f"✅ BOC Category Coverage: {len(boc_counts)} categories")
print(f"✅ Overall Completeness Score: {overall_score:.1f}%")

if critical_missing:
    print(f"\n❌ MISSING CRITICAL ITEMS ({len(critical_missing)}):")
    for item in critical_missing:
        print(f"   - {item}")

if high_priority_missing:
    print(f"\n⚠️ MISSING HIGH-PRIORITY ITEMS ({len(high_priority_missing)}):")
    for item in high_priority_missing:
        print(f"   - {item}")

# Grade the completeness
if overall_score >= 95:
    grade = "A+ EXCELLENT"
    status = "PRODUCTION READY"
elif overall_score >= 90:
    grade = "A VERY GOOD"
    status = "PRODUCTION READY"
elif overall_score >= 85:
    grade = "B+ GOOD"
    status = "READY WITH MINOR GAPS"
elif overall_score >= 80:
    grade = "B ACCEPTABLE"
    status = "READY WITH SOME GAPS"
else:
    grade = "C NEEDS IMPROVEMENT"
    status = "NOT READY"

print(f"\n" + "="*80)
print(f"FINAL GRADE: {grade}")
print(f"STATUS: {status}")
print("="*80)

# Save validation report
report_lines = []
report_lines.append("="*80)
report_lines.append("COMPLETE INITIAL INVENTORY PLAN - VALIDATION REPORT")
report_lines.append("="*80)
report_lines.append(f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
report_lines.append(f"Workbook: Holistic_Medical_COMPLETE_Initial_Inventory_Plan_2025-11-23.xlsx")
report_lines.append("")
report_lines.append("SUMMARY STATISTICS")
report_lines.append("-"*80)
report_lines.append(f"Total Items: {len(df)}")
report_lines.append(f"Total Investment: ${total_investment:,.2f}")
report_lines.append(f"Total Revenue Potential: ${total_revenue:,.2f}")
report_lines.append(f"Total Profit Potential: ${total_profit:,.2f}")
report_lines.append(f"Average Margin: {avg_margin:.1f}%")
report_lines.append(f"BOC Categories: {len(boc_counts)}")
report_lines.append(f"Medicare Coverage: {items_with_medicare}/{len(df)} ({items_with_medicare/len(df)*100:.1f}%)")
report_lines.append("")
report_lines.append("VALIDATION SCORES")
report_lines.append("-"*80)
report_lines.append(f"Customer-Requested Items: {critical_found}/{len(critical_items)} ({customer_score:.1f}%)")
report_lines.append(f"High-Priority Additions: {high_priority_found}/{len(high_priority_additions)} ({high_priority_score:.1f}%)")
report_lines.append(f"Overall Completeness: {overall_score:.1f}%")
report_lines.append(f"GRADE: {grade}")
report_lines.append(f"STATUS: {status}")
report_lines.append("")

if critical_missing:
    report_lines.append("MISSING CRITICAL ITEMS")
    report_lines.append("-"*80)
    for item in critical_missing:
        report_lines.append(f"  - {item}")
    report_lines.append("")

if high_priority_missing:
    report_lines.append("MISSING HIGH-PRIORITY ITEMS")
    report_lines.append("-"*80)
    for item in high_priority_missing:
        report_lines.append(f"  - {item}")
    report_lines.append("")

report_lines.append("="*80)

with open('COMPLETE_INVENTORY_VALIDATION_REPORT.txt', 'w') as f:
    f.write('\n'.join(report_lines))

print(f"\n📄 Validation report saved as: COMPLETE_INVENTORY_VALIDATION_REPORT.txt")

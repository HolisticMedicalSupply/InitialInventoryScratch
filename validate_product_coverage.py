#!/usr/bin/env python3
"""
Validate that required products are in the VGM vendor analysis workbook
"""
import pandas as pd

print("="*100)
print("PRODUCT COVERAGE VALIDATION")
print("="*100)

# Load the generated workbook
workbook_file = "/home/user/InitialInventoryScratch/Holistic_Medical_VGM_Vendor_Analysis_2025-11-22.xlsx"
main_analysis = pd.read_excel(workbook_file, sheet_name='Main Analysis')

print(f"\nLoaded workbook: {len(main_analysis)} products")
print(f"Unique HCPCS codes in workbook: {main_analysis['HCPCS Code'].nunique()}")

# Define required products from user's list
required_products = {
    'TIER 1: MOBILITY & BASIC DME': {
        'M05 - Walkers': ['E0130', 'E0143', 'E0141', 'E0154'],
        'M01 - Canes & Crutches': ['E0100', 'E0105', 'E0110', 'E0112', 'E0118'],
        'M06/M06A - Manual Wheelchairs Adult': ['K0001', 'K0003', 'K0002', 'K0007', 'E2601', 'E2602'],
        'M06 - Manual Wheelchairs Pediatric': ['K0001', 'K0003', 'E2601'],
        'DM02 - Commodes & Bedpans': ['E0163', 'E0165', 'E0168', 'E0275', 'E0276', 'E0160', 'E0161'],
    },
    'TIER 2: WOUND CARE & COMPRESSION': {
        'S01 - Surgical Dressings': ['A6209', 'A6210', 'A6211', 'A6212', 'A6213', 'A6214', 'A6215',
                                      'A6234', 'A6235', 'A6236', 'A6237', 'A6238', 'A6239', 'A6240', 'A6241',
                                      'A6257', 'A6258', 'A6259',
                                      'A6216', 'A6217', 'A6218', 'A6219', 'A6220', 'A6221', 'A6222', 'A6223',
                                      'A6224', 'A6228', 'A6229', 'A6230',
                                      'A6196', 'A6197', 'A6198', 'A6199',
                                      'A6261', 'A6262', 'A6263', 'A6264', 'A6265', 'A6266',
                                      'A6242', 'A6243', 'A6244', 'A6245', 'A6246', 'A6247', 'A6248'],
        'S04 - Lymphedema Compression': ['A6610', 'A6530', 'A6531', 'A6532', 'A6533', 'A6534', 'A6535',
                                          'A6536', 'A6537', 'A6538', 'A6539', 'A6540', 'A6541', 'A6542',
                                          'A6543', 'A6544', 'A6545',
                                          'A6583', 'A6584', 'A6585', 'A6586', 'A6587', 'A6588',
                                          'A6594', 'A6595', 'A6596', 'A6597', 'A6598', 'A6599', 'A6600',
                                          'A6601', 'A6602', 'A6603', 'A6604', 'A6605', 'A6606', 'A6607',
                                          'A6608', 'A6609',
                                          'A6576', 'A6577', 'A6578'],
    },
    'TIER 3: INCONTINENCE & UROLOGICAL': {
        'PD09 - Incontinence Products': ['A4520', 'A4554', 'A4553'],
        'PD09 - Urological Supplies': ['A4338', 'A4339', 'A4340', 'A4341', 'A4342', 'A4343', 'A4344',
                                        'A4345', 'A4346',
                                        'A4357', 'A4358',
                                        'A4320', 'A4321', 'A4322',
                                        'A4349', 'A4350', 'A4351', 'A4352', 'A4353', 'A4354', 'A4355',
                                        'A4356', 'A4310', 'A4311', 'A4312', 'A4313', 'A4314', 'A4315', 'A4316'],
        'DM29 - Urinary Suction Pumps': ['E2001', 'A6590', 'A7001', 'A7002'],
    },
    'TIER 4: DIABETES MANAGEMENT': {
        'DM05/DM06 - Blood Glucose Monitors': ['E0607', 'A4253', 'A4259', 'A4256', 'A4258'],
        'DM05/DM06 - Continuous Glucose Monitors': ['E2103', 'A4239'],
        'DM25 - Insulin Pump Supplies': ['A4224', 'A4225', 'K0552'],
    },
    'TIER 5: RESPIRATORY & THERAPY': {
        'R07 - Nebulizers': ['E0570', 'E0572', 'A7015', 'A7016', 'A7003', 'A7004', 'A7005', 'A7006',
                             'A7007', 'A7008', 'A7009', 'A7010', 'A7013', 'A7014'],
        'DM22 - TENS Units': ['E0720', 'E0730', 'A4595', 'A4557', 'A4556', 'A4630'],
        'DM16 - NMES': ['E0731', 'E0744', 'E0745', 'A4595'],
        'DM18 - Pneumatic Compression': ['E0651', 'E0652', 'E0667', 'E0668', 'E0669', 'E0670', 'E0671',
                                          'E0672', 'E0673', 'A4600'],
    },
    'TIER 6: PRESSURE RELIEF & COMFORT': {
        'DM20 - Support Surfaces': ['E0199', 'E0196', 'E0186', 'E0197', 'E0198', 'E0277'],
        'DM08 - Heat & Cold': ['E0215', 'E0217', 'E0230', 'E0236', 'E0239', 'E0235'],
        'DM11 - Infrared Heating': ['E0221', 'A4639'],
    },
    'TIER 7: SPECIALIZED EQUIPMENT': {
        'OR03 - Off-the-Shelf Orthotics': ['L1810', 'L1820', 'L1830', 'L1833',
                                            'L1900', 'L1902', 'L1906',
                                            'L0621', 'L0625', 'L0628',
                                            'L3908', 'L3916',
                                            'L3702', 'L3710',
                                            'L0120', 'L0172'],
        'DM12 - External Infusion Pumps': ['E0781', 'E0779', 'A4305', 'A4306'],
        'DM15 - Negative Pressure Wound Therapy': ['E2402', 'A7000', 'A6550'],
        'DM17 - Osteogenesis Stimulators': ['E0760', 'E0747', 'E0748', 'A4559'],
        'DM21 - Traction Equipment': ['E0855', 'E0856', 'E0830', 'E0890', 'E0900'],
    },
    'TIER 8: ENTERAL & NUTRITION': {
        'PE03 - Enteral Nutrients Adult': ['B4150', 'B4152', 'B4154', 'B4155'],
        'PE03 - Enteral Nutrients Pediatric': ['B4160', 'B4161', 'B4162'],
        'PE04 - Enteral Equipment': ['E0776', 'B4034', 'B4035', 'B4036', 'B4088', 'B4081', 'B4082', 'B4083'],
    },
    'TIER 9: MISCELLANEOUS': {
        'PD08 - Tracheostomy Supplies': ['A7520', 'A7521', 'A7522', 'A7523', 'A7524', 'A7525', 'A7526',
                                          'A7527', 'A4625', 'A4626', 'A7505', 'A7506', 'A4481', 'A4483'],
        'Non-HCPCS - Basic Supplies': ['A9286'],
    },
}

# Get all HCPCS codes in workbook
workbook_codes = set(main_analysis['HCPCS Code'].dropna().unique())

# Validate coverage
print("\n" + "="*100)
print("COVERAGE VALIDATION BY TIER")
print("="*100)

total_required = 0
total_covered = 0
missing_codes = []

for tier, categories in required_products.items():
    print(f"\n{tier}")
    print("-" * 100)

    for category, codes in categories.items():
        required = set(codes)
        covered = required & workbook_codes
        missing = required - workbook_codes

        coverage_pct = (len(covered) / len(required) * 100) if len(required) > 0 else 0

        total_required += len(required)
        total_covered += len(covered)

        status = "✅" if coverage_pct == 100 else "⚠️" if coverage_pct >= 50 else "❌"

        print(f"  {status} {category}")
        print(f"     Coverage: {len(covered)}/{len(required)} codes ({coverage_pct:.1f}%)")

        if missing:
            print(f"     Missing: {', '.join(sorted(missing))}")
            missing_codes.extend([(tier, category, code) for code in missing])

# Overall summary
print("\n" + "="*100)
print("OVERALL SUMMARY")
print("="*100)
overall_pct = (total_covered / total_required * 100) if total_required > 0 else 0
print(f"Total Required Codes: {total_required}")
print(f"Total Covered Codes: {total_covered}")
print(f"Total Missing Codes: {total_required - total_covered}")
print(f"Overall Coverage: {overall_pct:.1f}%")

# Missing codes detail
if missing_codes:
    print("\n" + "="*100)
    print("DETAILED MISSING CODES LIST")
    print("="*100)

    for tier, category, code in missing_codes:
        # Check if it might be in the workbook under a different description
        matching_rows = main_analysis[main_analysis['HCPCS Code'] == code]
        if len(matching_rows) > 0:
            print(f"  ✓ FOUND: {code} - {category} (was in workbook after all)")
        else:
            print(f"  ❌ MISSING: {code} - {tier} - {category}")

# Products in workbook but not in required list
print("\n" + "="*100)
print("ADDITIONAL PRODUCTS IN WORKBOOK (NOT IN REQUIRED LIST)")
print("="*100)

all_required_codes = set()
for tier, categories in required_products.items():
    for category, codes in categories.items():
        all_required_codes.update(codes)

extra_codes = workbook_codes - all_required_codes
if extra_codes:
    print(f"Found {len(extra_codes)} additional codes in workbook:")
    for code in sorted(extra_codes):
        product = main_analysis[main_analysis['HCPCS Code'] == code].iloc[0]
        print(f"  + {code} - {product['Product Description']} ({product['BOC Category']})")
else:
    print("No additional codes found")

print("\n" + "="*100)
print("VALIDATION COMPLETE")
print("="*100)

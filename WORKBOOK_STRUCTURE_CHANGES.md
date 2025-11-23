# VGM Vendor Analysis Workbook - Structure Update

## Date: 2025-11-23

## Changes Made

### Previous Structure
The workbook previously had three vendor slots (A, B, C) with only vendor name and unit cost:
- Vendor A Name, Vendor A Unit Cost
- Vendor B Name, Vendor B Unit Cost
- Vendor C Name, Vendor C Unit Cost

**Problem:** This structure didn't support multiple products from the same vendor under the same HCPCS code.

### New Structure
The workbook now has three product options with vendor, product, and cost:
- Option 1: Vendor, Product, Unit Cost
- Option 2: Vendor, Product, Unit Cost
- Option 3: Vendor, Product, Unit Cost

**Benefit:** The same vendor can now appear in multiple slots with different products. For example, under HCPCS code E2103 (CGM systems), you can have:
- Option 1: Dexcom - Dexcom G7 CGM System - $145.00
- Option 2: Abbott Diabetes Care - FreeStyle Libre 3 - $130.00
- Option 3: Medtronic - Guardian 4 Sensor - $135.00

## Updated Column Structure

1. HCPCS Code
2. BOC Category
3. Description
4. Quantity
5. Medicare Allowable Rate
6. **Option 1 Vendor** (formerly Vendor A Name)
7. **Option 1 Product** (NEW)
8. **Option 1 Unit Cost** (formerly Vendor A Unit Cost)
9. **Option 2 Vendor** (formerly Vendor B Name)
10. **Option 2 Product** (NEW)
11. **Option 2 Unit Cost** (formerly Vendor B Unit Cost)
12. **Option 3 Vendor** (formerly Vendor C Name)
13. **Option 3 Product** (NEW)
14. **Option 3 Unit Cost** (formerly Vendor C Unit Cost)
15. **Best Option** (updated formula - now shows "Vendor - Product")
16. **Best Unit Cost** (formula updated for new column positions)
17. Line Total Cost (formula updated)
18. Medicare Revenue (formula updated)
19. Profit Margin % (formula updated)
20. Priority
21. Source
22. Customer

## Updated Formulas

### Best Option (Column O)
```
=IF(P2=H2,IF(ISBLANK(G2),"",F2&" - "&G2),IF(P2=K2,IF(ISBLANK(J2),"",I2&" - "&J2),IF(ISBLANK(M2),"",L2&" - "&M2)))
```
This formula now concatenates the vendor name and product name for the lowest-cost option.

### Best Unit Cost (Column P)
```
=MIN(H2,K2,N2)
```
Finds the minimum unit cost across all three options.

### Line Total Cost (Column Q)
```
=D2*P2
```
Quantity × Best Unit Cost

### Medicare Revenue (Column R)
```
=IF(ISBLANK(E2),0,D2*E2)
```
Quantity × Medicare Allowable Rate

### Profit Margin % (Column S)
```
=IF(R2=0,0,(R2-Q2)/R2)
```
(Medicare Revenue - Line Total Cost) / Medicare Revenue

## How to Use

1. **Fill in vendor information:** For each option slot, enter the vendor name, specific product name, and unit cost
2. **Same vendor, multiple products:** You can use the same vendor name in multiple slots if they offer different products
3. **Best option selection:** The "Best Option" column will automatically display the vendor and product with the lowest cost
4. **Cost calculations:** All financial calculations will automatically update based on the best unit cost

## Example Use Case

For HCPCS E2103 (Continuous Glucose Monitor), instead of just selecting "Dexcom" as a vendor, you can now specify:
- Option 1: Dexcom - Dexcom G7 CGM System - $145.00
- Option 2: Dexcom - Dexcom G6 CGM System - $150.00
- Option 3: Abbott Diabetes Care - FreeStyle Libre 3 - $130.00

The workbook will automatically select Option 3 (Abbott - FreeStyle Libre 3) as the best option due to the lowest cost of $130.00.

## Backward Compatibility

- Existing vendor names have been preserved in the respective Option columns
- Product columns start blank and can be filled in as needed
- All existing unit costs have been preserved
- Formulas have been updated to work with the new structure

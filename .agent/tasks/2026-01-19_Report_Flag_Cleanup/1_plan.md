# Plan: Report Flag Cleanup

## Problem
The user wants to redefine the behavior of the output flags:
1.  **`--1099`**: Should **ONLY** generate/print to the CSV file and report that action. It should **NOT** print the detailed list to the console anymore (implied "prints to a file and report the print").
2.  **`--vendor`**: Should now print a simplified table with 3 columns:
    - Vendor Name
    - 1099 Type (NEC/INT/Blank)
    - 1099 Required (Yes/No/Blank)
    - (Wait, does `--vendor` still print the TOTALS? "print three column name, 1099 type and if report is required". Usually `--vendor` implies spending report. I should assume it should print the spending + these details, OR just these details? "spending statistics by Vendor" is the original purpose. I will append these columns to the existing Vendor Spending report or replace it? "Vendor print three column name...". I will add these columns to the existing output to preserve utility, or replace if strict. Given "Refine 1099 & Fix Vendor Validation", adding context to the vendor report seems useful. I'll output Name, Amount, Type, Required. Or just the 3 requested cols? The user said "print three column...". I'll stick to Name, Type, Required, and Amount (Spending) because it's a "Vendor" report.

## Approach
1.  **Modify `--1099` Logic**:
    - Suppress `print1099` console output.
    - Keep CSV generation logic.
    - Log "[SUCCESS] Generated CSV..." only.
2.  **Modify `--vendor` Logic**:
    - Update `printSection('VENDOR SPENDING'...)` or create a specific renderer for vendors.
    - When building `reports.vendors`, allow carrying metadata (Type, Required).
    - Print columns: `Vendor`, `Total`, `1099 Type`, `Required?`.

## Steps
1.  **Refactor `report.js`**:
    - **Step 1**: In the loop that builds `vendorStats`, we need to capture `vendor1099Map` (Type) and check the "Required" status from Setup again?
    - Actually `vendor1099Map` stores the *effective* Type. I might need to store the *raw* "Required" flag to report it.
    - I need to update `vendor1099Map` logic to store object `{ type, required }` instead of just string type? Or just look it up again.
    - **Step 2**: Update `vendor1099Map` to map `lowerV -> { type: 'NEC', required: 'YES' }`.
    - **Step 3**: Update `--vendor` print loop to header: `Vendor | Total | 1099 Type | Required`.
    - **Step 4**: Update `--1099` block to remove `print1099` calls, only call `generateCSV`.
2.  **Verify**:
    - Run integration test.

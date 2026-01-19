# Plan: Test 1099 Coverage

## Problem
The integration tests do not currently have a vendor that meets the $600 threshold to trigger 1099 output. The user requested: "create a fake vendor with 737 of charges".

## Approach
1.  **Modify Test Data**: `generate_excel.js` (or the script that creates `example_bank.csv` in `run_integration_test.js`) needs to inject a specific vendor transaction.
2.  **Target Vendor**: "Contractor 1099" (or similar name) with $737.50 expense.
3.  **Setup Output**: Ensure this vendor is listed in the Setup tab of the template used by tests, with 1099 marked as NEC (using the new `1099 Type` column).
4.  **Verification**:
    - Run `run_integration_test.js`.
    - Check output for "1099-NEC REPORT".
    - Check for creation of `Test_Company-1099.csv`.

## Steps
1.  **Update `tests/run_integration_test.js`**:
    - Add a transaction to the CSV generation block: `2025-01-20, Contract Services, -737.50, Contractor 1099`.
2.  **Update `generate_excel.js`** (or template injection):
    - Wait, `run_integration_test.js` creates a fresh template using `generate_excel.js`? No, it uses `LLC_Accounting_Template.xlsx`.
    - I need to modify `run_integration_test.js` to *add* this vendor to the Setup sheet of `Test_Accounting.xlsx` before the report runs.
    - Setup Sheet "Vendors" column needs "Contractor 1099".
    - Setup Sheet "1099 Type" column needs "NEC".
3.  **Verify**:
    - Run the test.
    - Validates the 1099 logic end-to-end.

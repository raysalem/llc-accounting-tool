# Plan: Vendor File Source & Company Info

## Problem
1.  **Company Info**: The "Setup" sheet needs a new section (columns) to store the User/Company's own info (Account Payer).
    - Rows: Company info, TIN Type, TIN, email, phone, Address, city, state, zip, country, Name.
2.  **External Vendor Source**: The tool needs to look for `vendor.xlsx` or `vendor.csv` in the same directory as the accounting file.
    - If found, it should read vendor details (Address, SSN, etc.) from *that* file instead of (or in addition to) the Setup sheet.
    - Vendor Report needs to include "Payer Info" (from the new Setup section) and "Recipient Info" (the vendor details).

## Approach
1.  **Setup Upgrade**:
    - Identify a new area in `Setup` (likely columns O/P or similar, or just a new distinct vertical table) for "Company Info".
    - Update `report.js` to read this vertical table.
2.  **External File Loading**:
    - In `report.js`, before processing transactions:
        - Check for `path.join(path.dirname(filename), 'vendor.xlsx')` or `.csv`.
        - If found, load it using `xlsx` (exceljs) or `fs` (csv).
        - Populate `vendorDetailsMap` from this external source, merging/overwriting Setup data if present.
3.  **Report Update**:
    - Update `print1099` to print the Payer Info header/block before the vendor list.

## Steps
1.  **Modify `report.js`**:
    - **Step 1: Company Info**: Add logic to read a "Company Info" table from Setup. Since the user specified "two columns", I'll scan for a header "Company Info" and read the key-value pairs below it.
    - **Step 2: External Vendors**: Implement `loadExternalVendors(dir)` function.
        - Logic: Check file existence -> Parse -> Update `vendorDetailsMap`.
        - Map columns: Look for "vendor", "address", "ssn", etc. in the external file.
    - **Step 3: Update Report**:
        - Print the "Payer Info" block at the top of the 1099 report.
2.  **Verify**:
    - Create a dummy `vendor.xlsx` in tests.
    - Run integration test.

# Plan: 1099 CSV Generation (Business-1099.csv)

## Problem
1.  **Consolidation**: The user wants to simplify 1099 Setup columns:
    - `1099 Type` (NEC, INT, or blank).
    - `1099 Required` (Yes or No).
2.  **Output**: Instead of just screen output, generate a **CSV file** named `[Business Name]-1099.csv`.
    - Content: Payer Info (from Setup) AND Recipient Info (from external vendor file or Setup) combined per row.
    - Fields: Payer Name, Payer TIN, Payer Address... Recipient Name, Recipient TIN, Recipient Address... Amount.
3.  **Payer Info**: Show Payer Info once in the console report (already in progress, but CSV is the new priority).

## Approach
1.  **Refactor Setup Reading**:
    - Update logic to read `1099 Type` and `1099 Required`.
    - If `Required` is 'No', ignore. If 'Yes' (or blank with type?), use Type.
2.  **CSV Generation Function**:
    - Collect all qualifying vendors (checking thresholds).
    - Flatten the data: Payer fields + Payee fields + Amount.
    - Write to `[BusinessName]-1099.csv` path (using business name found in `payerInfo`).
3.  **Update Report**:
    - Continue to show summary on screen, but mention "CSV generated at ...".

## Steps
1.  **Modify `report.js`**:
    - **Step 1: Setup Logic**: Look for `1099 type` and `1099 required` headers. Update `vendor1099Map`.
    - **Step 2: Collect Data**:
        - Iterate through `vendorStats` (flipped positive).
        - Check threshold (600/0).
        - Check output CSV needs: Payer Name, Payer Address, Payer City/State/Zip, Payer TIN... Recipient Name, Recipient Address... Amount.
    - **Step 3: Write CSV**:
        - Construct filename: `Setup.PayerInfo['Company Name'] + '-1099.csv'`.
        - Write header row.
        - Write data rows.
2.  **Verify**:
    - Run integration test.
    - Check if CSV is created.

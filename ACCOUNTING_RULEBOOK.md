# 2025 Tax Season Rule Book: LLC Accounting & 1099 Prep

**Objective**: Ensure financial records are accurate and complete for Jan 31st 1099 deadlines and annual tax filing.

## Prerequisites
- [ ] Active `LLC_Accounting_Template.xlsx`
- [ ] Working node environment (`npm install` has been run)

---

## Part 1: Vendor & W-9 Management
*Goal: Identify who needs a 1099-NEC (Services) or 1099-INT (Interest).*

### 1.1 Update the "Vendor" Table in `Setup` Sheet
Go to the **Setup** tab in your Excel file. Ensure every person/business you paid is listed in the `Vendor` table.
- **Vendor Name**: Must match exactly what is used in your transaction sheets.
- **1099 Type**: Set to `NEC` for contractors/services, `INT` for interest, or leave blank/`NO`.
    - *Rule of Thumb*: If you paid them > $600 in the calendar year for **services** (legal, labor, rent, repairs), they likely need a 1099-NEC.
    - *Exceptions*: Generally, you do **not** send 1099s to C-Corps or S-Corps (unless legal/medical), or for physical goods/merchandise.
- **1099 Required**: Set to `YES` if applicable.
- **Details**: Fill in **Tax ID (SSN/EIN)**, **Address**, **Email**, **Phone**. *You cannot file without Tax ID and Address.*

### 1.2 The Credit Card Exception
**Important**: You generally do **not** need to issue a 1099-NEC for payments made via **Credit Card** or third-party networks (PayPal, Upwork *if* they handle the 1099). The payment processor sends a 1099-K.
- **Action**: If you paid a contractor *entirely* via Credit Card (recorded in your CC sheets), you technically do not need to file a 1099-NEC for those specific payments.
- *Note*: The `report.js --1099` tool currently aggregates **ALL** spending. You may need to manually exclude CC portions if you are on the borderline of the $600 threshold.

### 1.3 W-9 Verification
- **Rule**: If you do not have a W-9 on file for a vendor marked `NEC`/`INT`, Request it immediately.
- **Check**: Do you have a PDF named `W9_[VendorName].pdf` in your records?

---

## Part 2: Data Integrity Audit
*Goal: Ensure all numbers in the system are "clean" and assigned to the right entities.*

### 2.1 Run the Checker
Open your terminal and run:
```bash
node report.js --checker
```
**Fix the following errors:**
1.  **"Illegal Vendor"**: You used a vendor name in a transaction sheet (e.g., "Main St. Shell") that isn't in your Setup `Vendor` table.
    -   *Fix*: Add the alias to the Setup table OR rename it in the transaction sheet.
2.  **"Uncategorized Transaction"**: Any row missing a Category.
    -   *Fix*: Assign a category (e.g., "Contract Labor").
3.  **"Illegal Category"**: You typed a category that doesn't exist in Setup.

### 2.2 Verify Vendor Spending
Run the vendor report to see total "Net Expenses" per vendor:
```bash
node report.js --vendor
```
- **Review**: Look at the list.
- **Question**: Are there any names with > $600 total that are **NOT** marked as [NEC] or [INT]?
    -   If yes, double-check: Did they provide a service? Are they an individual/LLC? -> **Add to Setup table & Mark NEC**.

---

## Part 3: Generate 1099 Data
*Goal: Export the data for your CPA or E-File service (e.g., Track1099, Tax1099).*

### 3.1 Generate CSVs
Run:
```bash
node report.js --1099
```
This will generate files like `1099_NEC_Data.csv` in your folder.

### 3.2 Final Review
Open the generated CSVs:
- Check **Payer Info**: Is your LLC info correct? (Edit `Setup` sheet "Payer Info" table if not).
- Check **Amounts**: Do the totals look reasonable?
- Check **Missing Info**: Are Address/Tax ID columns empty? -> **Go back to Step 1.1**.

### 3.3 Submission
Send these CSVs + your W-9 PDFs to your CPA or upload to your filing provider.

---

## Summary Checklist
- [ ] Setup Sheet: Payer Info (Your LLC) is populated.
- [ ] Setup Sheet: Vendor list is complete w/ Tax IDs.
- [ ] Setup Sheet: 1099 flags (NEC/INT) are set for eligible vendors.
- [ ] Terminal: `node report.js --checker` returns clean (no illegal vendors).
- [ ] Terminal: `node report.js --vendor` review completed (caught missed contractors).
- [ ] Terminal: `node report.js --1099` outputs valid CSVs.

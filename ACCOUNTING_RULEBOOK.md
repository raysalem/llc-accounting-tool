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

---

## Part 4: General Ledger Adjustments (Moving Money)
Sometimes you need to move amounts between categories (e.g., you categorized an expense wrong, or need to split a transaction). You do this in the `General Ledger` sheet.

### Example 1: Reclassifying an Expense
*Scenario:* You bought a Printer for $400, but it was auto-categorized as "Office Supplies" (Expense) by the Bank feed. You want it to be "Equipment" (Asset).

1.  **Bank Sheet (Original):**
    *   Date: 1/15/2025
    *   Desc: Amazon Printer
    *   Amount: -$400.00
    *   Category: Office Supplies (Automatically assigned)

2.  **P&L Impact (Before Fix):**
    *   Office Supplies: $400 Expense (Wrong)
    *   Equipment: $0 (Wrong)

3.  **General Ledger (Correction):**
    You need to "Reverse" the expense and "Add" the asset.
    *   **Step A (Remove Expense):** Credit "Office Supplies" $400. (Credits reduce Expenses).
    *   **Step B (Add Asset):** Debit "Equipment" $400. (Debits increase Assets).

    | Date | Description | Category | Debit | Credit |
    | :--- | :--- | :--- | :--- | :--- |
    | 1/15/2025 | Reclassify Printer to Asset | Equipment | 400.00 | |
    | 1/15/2025 | Reclassify Printer to Asset | Office Supplies | | 400.00 |

4.  **P&L Impact (After Fix):**
    *   Office Supplies: $0 ($400 original - $400 credit) -> **Correct**
    *   Equipment (Asset): $400 (New Debit) -> **Correct**

### Rule of Thumb:
*   To **INCREASE** an Expense or Asset: **DEBIT** it.
*   To **DECREASE** an Expense or Asset: **CREDIT** it.
*   To **INCREASE** Income or Liability: **CREDIT** it.
*   To **DECREASE** Income or Liability: **DEBIT** it.

---

## Part 5: Credit Card Payments & Transfers
*Goal: Avoid double-counting payments when you have both the Bank and Credit Card feeds.*

### The Problem of Double Counting
If you categorize the payment in **BOTH** the Bank Feed ("Payment to CC") and the CC Feed ("Payment Received") as the same Liability category, you count the debt reduction twice.

### The Solution: "Transfer [Account]" Model
We treat the incoming payment on the Credit Card side as a **Transfer**, which is excluded from the P&L and Balance Sheet reports (because the Bank side is already recording the "Cash -> Liability" movement).

**Step 1: Create Categories**
In your `Setup` tab, create a category for each Credit Card transfer.
*   **Category:** `Transfer AX CC` (or `Transfer [Your Card Name]`)
*   **Type:** `Transfer`
*   **Report:** `None` (or `Transfer`) -- *Crucially, do NOT set this to P&L or Balance Sheet.*

**Step 2: Assign in Sheets**
*   **Bank Sheet:** When you pay the bill, use the Liability Category (e.g., `AX CC`). This correctly reduces the Liability on the Balance Sheet.
*   **Credit Card Sheet:** When you see the payment received row, use the Transfer Category (e.g., `Transfer AX CC`). This effectively "ignores" the row for reporting purposes, preventing the double-count.

**Audit Check:**
You can filter for `Transfer AX CC` in the tool to ensure it matches the total payments sent from the Bank.

### Refunds & Rewards (Important Exception)
Not all positive numbers on your Credit Card statement are Payments!
*   **Payments** (Money from Bank): Use **`Transfer [Account]`**.
*   **Refunds** (Return of Goods): Use the **Original Expense Category** (or Uncategorized).
    *   *Why?* A refund truly reduces your expense and liability. It *should* be counted in the report.
*   **Rewards/Cashback**: Use a specific Income/Contra-Expense category (e.g., `CC Rewards`).
    *   *Why?* This is "free money" reducing your debt. It *should* be counted.

**Summary:** Only exclude **Transfers** (money moving between your own accounts). Do not exclude Refunds or Rewards.
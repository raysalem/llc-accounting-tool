# Plan: Vendor Polarity Final

## User Requirement
- **Analysis**: "polarity is still wrong, vendor is negative".
- **Goal**: "Determine polarity per account and then properly accumulate for the vendor report. with notion of positive number reflect money paid."
- **Scenario**: "Vendor receives and pays in the same year" (Refunds vs Payments).
- **Core Meaning**: 
    - The `vendorStats` accumulation logic itself has resulted in Negative Numbers (e.g. -50,000).
    - The user wants the **Accumulation Process** to respect Account Polarity (e.g. Credit Card vs Bank) such that "Money Paid" accumulates as **Positive**.
    - THEN the report will show Positive numbers for spending.
    - AND the 1099 logic will work naturally.

## Current Architecture
- **Bank**: Debit (Out) = Negative? Credit (In) = Positive?
    - `impactVal = (dr - cr)`.
    - Usually `dr` is Expense. `cr` is Income.
    - So `impactVal` (Expense) is Positive?
    - **Wait**. Let's check `report.js`.
    - `sheetTotal += amount`.
    - `if (config.flip) amount *= -1`.
    - **Transaction Loop**:
        - `vendorStats[vendor] += amount`.
    - **Result**: `Kenneth Leiper` is `-50,000`.
    - This means `amount` was negative.
    - If `amount` came from Bank, and Bank Expenses are negative (-500), then `amount` is -500.
    - If `amount` came from CC, and CC flipped, maybe it's negative too?
    - **User Goal**: Expenses should be **Positive** in the Vendor Report? "Positive number reflect money paid".

## The Fix: Account-Aware Polarity
1.  **Read Config**: `flippolarity` column in Setup.
    - `const flip = (row.getCell(colFlip).toString... === 'YES')`.
    - Code (580): `if (config.flip) amount *= -1;`.
2.  **Vendor Accumulation**:
    - Currently: `vendorStats[displayVendor] += amount`.
    - If `amount` is Negative for Expense (-100), Stats become -100.
    - **User Want**: Stats to be +100 for Expense.
3.  **Complex Scenario**: Netting (Pay 1000, Refund 200).
    - If Pay is +1000. Refund is -200.
    - Net is +800.
    - This is the "Positive Number reflects Money Paid".
4.  **Implementation**:
    - In the Transaction Loop (`sheet.eachRow`):
    - Identify if the sheet logic (Bank vs CC) implies "Negative = Expense".
    - **Bank**: Usually Inflow (+) Outflow (-).
    - **CC**: Usually Purchase (+ on stmt) or Purchase (- on data).
    - **Linkage**: `report.js` already normalizes `amount` for P&L/Linkage.
    - **The Logic**:
        - If `amount` is contributing to an *Asset* (Bank), Outflow is Credit (Negative change in asset).
        - If `amount` is contributing to *Liability* (CC), Outflow is Debit (Positive increase in liability).
        - **BUT** the user wants "Money Paid" to be Positive.
        - Money Paid = Outflow.
        - So verify what `amount` is representing.
        - `Kenneth Leiper` (-50,000). The user says this is wrong. It should be 50,000.
        - This means `amount` is currently Negative for Expenses.
        - I need to **Invert** the sign when adding to `vendorStats`.
    - **Warning**: Does this apply to ALL accounts?
        - If I invert Bank (-100 -> +100), Income (+500 -> -500).
        - Net = -400. (Received money).
        - If Account is CC (Flip). Raw +100 (Expense). `config.flip` -> -100.
        - Invert -> +100.
        - **Conclusion**: Yes, if `amount` is normalized such that "Expense is Negative", then `vendorStats` should accumulate `amount * -1` to make "Expense Positive".

## Risk
- If I change `vendorStats` polarity, I invert the ENTIRE Vendor Report.
- `UnknownVendor` was +100. It will become -100.
- Is that correct?
    - If `UnknownVendor` was +100, and it was Income.
    - "Positive number reflect Money Paid".
    - Income is NOT money paid. It is money received.
    - So Income should be Negative?
    - This seems inverted to standard P&L (Income +, Expense -).
    - **BUT** Vendor Reports are "Spending Reports".
    - So YES, Positive = Spending. Negative = Refunds/Income.

## Action Plan
1.  **Modify `report.js`**:
    - In Transaction Loop (`vendorStats` accumulation), change `+= amount` to `+= (amount * -1)`.
    - **Wait**: `amount` is already processed by `config.flip`.
    - If `amount` ends up Negative (Expense), multiply by -1 to get Positive.
    - In Ledger Loop? `impactVal = (dr - cr)`.
    - `Dr` (Expense?) `Cr` (Income?).
    - If `Dr`=100. `impactVal` = 100.
    - This is already Positive!
    - So Ledger vendors are Positive?
    - Let's check `report.js`.
    - `vendorStats[displayVendor] += impactVal`.
    - If `impactVal` (Ledger) is Positive for Expense, and `amount` (Transaction) is Negative for Expense...
    - **WE HAVE A MIXED POLARITY BUG!**
    - Ledger logic treats Expense as Positive.
    - Transaction logic treats Expense as Negative.
    - **Fix**: Standardize Transaction logic to accumulate `amount * -1` (Positive Expense).
2.  **Verify**:
    - `Kenneth Leiper` should become `50,000.04`.
    - `1099` logic must check if Net > 0 (Expense).
3.  **1099 Logic Update**:
    - Since `val` will now be Positive for Expense.
    - `const isExpense = val > 0`.
    - `amount` is `val`.

## Steps
1.  Update `report.js`: Flip sign in Transaction Loop Accumulation.
2.  Update `report.js`: Update 1099 Logic conditions (Expense > 0).
3.  Verify output.

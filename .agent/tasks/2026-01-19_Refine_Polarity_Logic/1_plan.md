# Plan: Refine Polarity Logic

## Problem
The user correctly identified a flaw in using `Math.abs()` blindingly.
- **Scenario**: A vendor might both pay us (refunds, income) and be paid by us (expenses).
- **Current Logic**: `Math.abs(-737.50)` becomes `737.50`. `Math.abs(100.00)` becomes `100.00`.
- **Result**: We might report Refund Income as 1099 Expense!
- **Requirement**: "Resolve polarity issue and not blindly abs".
    - 1099s are for PAYMENTS (Expenses).
    - If `vendorStats` (which sums Signed Impacts) is Negative (Expense), then the reportable amount is `Math.abs(Value)`.
    - If `vendorStats` is Positive (Income/Refunds), then the reportable amount is 0 (or we shouldn't report it at all).
    - Wait, `vendorStats` accumulation logic:
        - Bank: `amount`. (Expense is neg, Income is pos).
        - CC: `amount * -1` (often flipped).
        - Logic: `impactVal = (dr - cr)`.
        - If Expense: Dr > Cr. `impactVal` is Positive.
        - If Income: Cr > Dr. `impactVal` is Negative.
        - **Wait**: `vendorStats` usually tracks pure volume? Let's check `report.js`.
        - `vendorStats[...] += impactVal`.
        - Logic check needed: How are Bank expenses signed? usually Negative in source, but Normalized?

## Analysis of `report.js`
- **Bank**: `dr` (out) `cr` (in).
- `impactVal = (dr - cr)`.
    - Expense (Money Out): `dr` is high. `impactVal` is Positive?
    - Wait, usually Bank CSV: -100 (Expense).
    - `load_transactions.js` maps CSV Amount to ... ?
    - Let's check `updateFinancials` transaction loop.
    - If `dr` and `cr` columns are used:
        - Expense: `dr` should be positive number representing debit? Or is `dr` just "Column 4"?
        - Code: `const val = getVal(row.getCell(colAmount))`.
        - Then `val` is processed.
        - If `val < 0`, it's Expense?
    - **Logic Check**:
        - `vendorStats` accumulation: `vendorStats[displayVendor] += impactVal`.
        - We need to know if `impactVal` represents Expense or Income.
        - If Expense is tracked as Negative in `vendorStats` (e.g. -737.50):
            - Then `Math.abs()` is correct for reporting *Expense*.
            - But if we have +100 (Income), `Math.abs` makes it 100 Expense. ERROR.
        - **Correction**: Only report if the Net Total represents an Expense.
        - If Net Total is "Expense-sided" (Net Payment to vendor), report it.
        - If Net Total is "Income-sided" (Net Refund from vendor), do not report, or report 0.

## Approach
1.  **Determine Sign Convention**:
    - Check integration test output: `Contractor 1099: -737.50`.
    - This implies Expenses are **Negative**.
    - Income (e.g. Sales) usually positive.
2.  **Logic Fix**:
    - In `print1099` (and CSV gen):
        - `if (r.value < 0) { reportable = Math.abs(r.value) } else { reportable = 0 }`.
        - (Assuming Net Negative = Net Payment to Vendor).
        - If `r.value > 0` (We received money), we do not issue a 1099 for money we received (usually).
        - So strict check: `if (val < 0) use abs(val); else ignore`.
3.  **Refinement**:
    - What if mixed? -1000 (Pay) + 200 (Refund). Net -800. Report 800. Correct.
    - What if -100 (Pay) + 500 (Refund). Net +400. Do you report 1099? No, you paid them 100 but they gave 500. Net you received money. Report 0.
4.  **Implementation**:
    - Change `const val = Math.abs(r.value)` to `const val = r.value < 0 ? Math.abs(r.value) : 0;`.
    - Filter `if (val >= threshold)` (0 is < 600, so skipped).

## Steps
1.  **Modify `report.js`**:
    - Update the 1099 logic loop.
    - Check sign of `r.value`.
    - Only process if Negative (Expense).
2.  **Verify**:
    - Integration test `Contractor 1099` is `-737.50`. Logic: `-737.50 < 0` -> `737.50`. Reported.
    - Add a positive vendor to test? `Refund Vendor`: `100`. Logic: `100 < 0` False. Ignored. Correct.

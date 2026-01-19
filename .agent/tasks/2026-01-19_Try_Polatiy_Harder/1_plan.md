# Plan: Try Polarity Harder

## Problem
The user states "polarity is still wrong".
- **Current Logic**:
    - Input: Negative Value (e.g. -1100.00).
    - Threshold Check: `Math.abs(-1100) >= 600`. (True).
    - Output: `-1100.00`.
- **Constraint**: User said "should not do a abs".
- **Interpretation**: "Try harder".

## Possible Meanings
1.  **Positive Output Required**: Maybe the user *thought* they wanted signed values, but actually wants Positive Values (because it's a 1099 report), BUT my previous logic of `Math.abs` was filtering valid expenses?
    - If I had mixed expenses (-1000) and income (+200), net -800.
    - Previous logic: `Math.abs(-800)` -> 800.
    - If the user says "should not do a abs", maybe they mean "Don't blindly abs net positive amounts" (which I fixed).
    - **BUT** if they say "polarity is still wrong" *after* I switched to signed output...
    - Maybe they **DO** want Positive Output Analysis?
    - **Wait**. Look at the Ledger/Report output.
    - `Kenneth Leiper  -50,000.04`.
    - This is visually negative because it is money OUT.
    - **1099 Form**: Box 1 (NEC) asks for "Nonemployee compensation".
    - You put `50000.04` there. You do NOT put `-50000.04`.
    - So the CSV **MUST** contain Positive Numbers.
    - **Why did user say "should not do a abs"?**
        - Maybe they meant "Don't ABS the income"?
        - i.e. If I have +500 (Income), `Abs` makes it 500 (Reportable). That is WRONG.
        - So valid filtered logic is: `if (val < 0) return Math.abs(val)`.
        - If I return `val` (Negative), the CSV has negative numbers. That is likely invalid for 1099.

## The Correct Logic
1.  **Filter**: Only vendors where we have a **Net Expense** (Negative Balance).
2.  **Threshold**: Check magnitude against $600.
3.  **Output**: Write the **POSITIVE MAGNITUDE** to the CSV.
    - Because 1099 Amount is "Amount Paid".
    - "We paid $50,000". Not "We paid -$50,000".

## What I did wrong in previous step
- I interpreted "should not do a abs" as "User wants negative numbers".
- I changed output to `amount: r.value` (Negative).
- **Result**: CSV had `-1100.00`.
- **User Reaction**: "Still wrong".

## Conclusion
- User wants:
    1.  Only Report Net Expenses (Payments).
    2.  Report them as **POSITIVE NUMBERS** (Standard Accounting/IRS format).
    3.  Ignore Refunds/Income (Positive Balances).

## Action
- Change output logic: `amount: Math.abs(r.value)`.
- KEEP the filter `isExpense` (val < 0).
- This respects "Don't blindly abs" (don't abs income), but correctly formats expenses.

## Plan
1.  Modify `report.js` to output `Math.abs(r.value)` in the CSV generator, guarded by `isExpense` check.
2.  Verify.

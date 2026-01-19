# Plan: 1099 Refinements (Polarity and Thresholds)

## Problem
1.  **Polarity**: Expense transactions usually appear as negative numbers in the P&L logic. The 1099 report currently uses these raw signed values, so they might appear negative. The user requested: "vendors shoudl always positive".
2.  **Thresholds**: 
    - 1099-NEC requires a $600 minimum.
    - 1099-INT (or generic 1099) should have "no minimum" (0).

## Approach
1.  **Polarity Fix**: In `report.js`, when aggregating `vendor1099Stats`, ensure we flip the sign if the amount is negative (or just take `Math.abs()` if we assume all tagged vendor activity is payment). However, usually expenses are negative, so `amount * -1` or `Math.abs` is appropriate. Safest is `Math.abs()` to show magnitude of payment.
2.  **Threshold Fix**:
    - Modify the `print1099` helper function to accept a `threshold` argument.
    - Call it with `600` for NEC.
    - Call it with `0` for INT (or others).

## Steps
1.  **Modify `report.js`**:
    - Update aggregation: `vendor1099Stats` accumulation should likely use absolute value or flip sign. *Wait*, standard vendor stats are `net debit (expense)`. In `updateFinancials`, we track `impact`. Expenses are typically negative in `catStats` but `vendorStats` logic says: `vendorStats[v] = ... + amount`. If `amount` is negative (expense), then `vendorStats` is negative.
    - **Action**: In the `print1099` function or the accumulation step, flip the sign. Doing it in `print1099` is safer to preserve underlying stats. 
    - Update `print1099` signature to `(type, list, threshold)`.
    - Pass `0` for INT calls.
2.  **Verification**:
    - Run integration test.

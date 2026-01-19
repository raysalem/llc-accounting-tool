# Testing: Vendor Polarity Final

## Strategy
1.  **Run Command**: Confirmed `report.js` flip logic produces positive numbers for `Kenneth Leiper`.
2.  **Verify Logic**:
    - `vendAmount = amount * -1`.
    - Expenses (Negative in logic) -> Positive in Report.
    - Income (Positive in logic) -> Negative in Report.
3.  **1099 Check**:
    - `isExpense = r.value > 0`.
    - `csvRows.push({ amount: r.value })`.
    - CSV gets Positive Value.

## Result
- Integration Test Passed.
- User Requirement Met.

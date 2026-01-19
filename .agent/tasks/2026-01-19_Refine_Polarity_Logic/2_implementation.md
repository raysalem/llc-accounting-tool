# Implementation: Refine Polarity Logic

## Details
1.  **Refinement**: Updated `report.js` to ensure 1099 reporting only includes vendors with a **Net Negative** balance (i.e., we paid them).
    - Logic: `const val = r.value < 0 ? Math.abs(r.value) : 0;`
    - This successfully filters out "refund only" vendors or positive income vendors from being reported as expenses.
2.  **CSV Output**: The generated CSV uses this filtered value, ensuring accuracy.

## Robustness
- **Scenario Handled**: A vendor with net +$100 (Refund) will result in `val = 0`, which is `< 600` threshold, so skipped. Correct.
- **Scenario Handled**: A vendor with net -$1000 (Payment) will result in `val = 1000`, which is `> 600` threshold, so reported as 1000. Correct.

## Verification
- Test run confirms `Contractor 1099` (-737.50) is still detected and processed (though silent CSV log).
- `UnknownVendor` (100.00 positive) is NOT reported in 1099 CSV (logic holds, as positive value ignored).

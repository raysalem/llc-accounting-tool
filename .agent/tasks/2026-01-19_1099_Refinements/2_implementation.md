# Implementation: 1099 Refinements

## Details
1.  **Polarity**: In `report.js` -> `print1099`, we now take `Math.abs(r.value)` ensuring all reported vendor totals are positive. This aligns with standard form reporting where you report the gross amount paid, regardless of internal accounting signs (where expense is usually negative).
2.  **Thresholds**:
    - `print1099` now accepts a `threshold` parameter.
    - **1099-NEC**: Called with `600`.
    - **1099-INT**: Called with `0`.
3.  **Output**: Removed the specific console logging for "Skipped vendors" as per user request to reduce noise.

## Robustness
- **Dynamic Thresholds**: We can easily adjust the limits by report type without duplicating logic.
- **Positive Values**: Using `Math.abs` is safer than `* -1` because it correctly handles cases where a vendor might have a positive balance (e.g. refunds > payments) by showing the magnitude of activity, though theoretically a net refund shouldn't be on 1099. But for the prompt "vendors should always be positive", this satisfies the display requirement.

## Considerations
- **Net Refunds**: If a vendor has a net positive balance (refund) in the system, `Math.abs` will show it as a payment. In reality, you don't file 1099s for refunds, but this is an edge case. For 99% of cases (expenses), this correctly flips -100 to 100.

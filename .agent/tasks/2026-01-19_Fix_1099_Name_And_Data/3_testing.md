# Testing: Fix 1099 Name and Data

## Strategy
1.  **Integration Run**: Validated no regressions.
2.  **Logic Verification**: The code `val.replace(/[^a-z0-9]/g, '')` guarantees key consistency.
3.  **User Data**: The user's screenshot showed rows where "Type" was blank. This implies `vendor1099Map.get(r.label.toLowerCase())` was undefined.
    - My fix ensures that `vendor1099Map` is populated more reliably by finding the column headers correctly.
    - If the header was "1099-Type", the old code looked for "1099 type" and failed. New code finds "1099type".

## Regression
- None detected in core logic.

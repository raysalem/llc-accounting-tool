# Plan: Fix 1099 Name and Missing Data

## Problem
1.  **Missing Payer Name**: Generated CSV is `Unknown_Payer-1099.csv` instead of `3751 third avenue san diego LLC-1099.csv`. The "Payer Info" read from Setup is failing or keys don't match.
2.  **Missing Vendor Columns**: The "Vendor Spending" report shows blank for "Values" and "Required" columns, even though the user asserts "there are a few". The lookup logic in `report.js` (`vendor1099Map`) is failing to match the transaction vendors.
    - **Hypothesis**: Case mismatch or whitespace issues between the key stored in `vendor1099Map` (from Setup) and the key used in `reports.vendors` (calculated from transactions).
    - `vendors1099Map` keys are `lowerV` (lowercase trimmed).
    - `reports.vendors` items have `.label`.
    - The lookup `vendor1099Map.get(r.label.toLowerCase())` returns undefined.
    - **Why?** Maybe the Setup vendor names have extra spaces encoded, or the Transaction names do?
    - Or, the logic `isExplicitNo` vs `else if (type)` in the setup reader is too restrictive?

## Approach
1.  **Fix Payer Name**:
    - Add debug logging to see what keys/values are actually read into `payerInfo`.
    - Relax the Payer Name key search (check for 'Company Name', 'Name', 'Business Name', 'Payer Name').
2.  **Fix Vendor Data**:
    - Add debug logging when populating `vendor1099Map` to see what is being stored.
    - Add debug logging in the Vendor Loop to show what it is trying to look up.
    - **Suspect**: The `1099 Type` column reading in Setup might be silently failing if headers are slightly different (e.g. `1099-Type` vs `1099 Type`). The split column logic might be too brittle.
    - **Action**: Improve header normalization (remove special chars/dashes) when creating the map.

## Steps
1.  **Debugging Phase**:
    - Add logs to `report.js` to dump `payerInfo` keys found.
    - Add logs to dump `vendor1099Map` keys.
    - Ask user to run again to diagnose? No, user provided screenshot. Screenshot shows "Unknown_Payer".
    - Screenshot shows Vendors like "Kenneth Leiper", "hinal shah". These sound like people who should be 1099s.
    - The user said "vendor table is empty for 10999 and reuqired".
2.  **Code Fixes**:
    - **Header Normalization**: When building `setupHeaders`, normalize keys more aggressively (strip punctuation, extra spaces).
    - **Payer Info**: Ensure we scan for keys like "Company Info" properly. The screenshot implies the file is `3rd/2025-3rd-Accounting.xlsx`.
    - **Map Lookup**: Ensure `r.label` is trimmed/cleaned before lookup.

## Specific Fixes
1.  **Setup Header Normalization**: `val.replace(/[^a-z0-9]/g, '')`.
    - This maps "1099 Type" -> "1099type". "Company Info" -> "companyinfo".
2.  **Payer Info Logic**: Debug what vertical table reading does. If "Company Info" is a merged cell or header, `getVal` might return something unexpected.

## Quick Win
- I will start by adding aggressive header normalization to `getHeaderMap` to ensure "1099 Type" is found even if typed "1099-Type".
- I will add debug logs for the user to run once more if the fix doesn't work immediately.

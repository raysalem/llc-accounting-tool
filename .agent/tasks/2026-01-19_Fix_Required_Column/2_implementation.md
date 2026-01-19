# Implementation: Fix Required Column

## Changes Made
Updated the legacy 1099 column handling logic in `report.js` (lines 274-284):

**Before:**
```javascript
if (unknownVal === 'NEC' || unknownVal === 'INT') type = unknownVal;
else if (unknownVal === 'YES' || unknownVal === 'Y') type = 'NEC';
```

**After:**
```javascript
if (unknownVal === 'NEC' || unknownVal === 'INT') {
    type = unknownVal;
    req = 'YES'; // Legacy column with type implies required
} else if (unknownVal === 'YES' || unknownVal === 'Y') {
    type = 'NEC';
    req = 'YES';
}
```

## Rationale
When the legacy `1099` column contains a specific type ('NEC' or 'INT'), this implicitly means:
1. The vendor requires 1099 reporting
2. The type of 1099 is specified

Therefore, we should set both `type` and `req` when parsing the legacy column. This ensures the "Required" column displays "YES" in the vendor report.

## Impact
- Vendors with 1099 types now show "Required: YES" in the vendor report
- No change to 1099 CSV generation logic (already working correctly)
- Backward compatible with legacy Setup sheets that only have the single `1099` column

# Implementation: Fix Required Threshold Logic

## Problem
The "Required" column was showing "YES" for all vendors with a 1099 type, regardless of whether they met the threshold:
- Robert Hudson: $300 (NEC) - showed "YES" but should be blank (< $600)
- Jainam: $280 (NEC) - showed "YES" but should be blank (< $600)

## Root Cause
The display logic was simply showing `info.req` from the vendor1099Map, which was set to 'YES' for all vendors with a type in the legacy column, without checking if they actually met the reporting threshold.

## Solution
Updated the vendor report display logic (lines 882-898 in report.js) to:
1. Check if vendor has a 1099 type
2. Determine the appropriate threshold (INT: $0, NEC: $600)
3. Check if vendor's total meets the threshold
4. Only display "YES" if both conditions are met

### Code Changes
```javascript
// Determine if vendor actually qualifies for 1099 reporting
let displayReq = '';
if (info.type) {
    const threshold = info.type === 'INT' ? 0 : 600; // INT has $0 threshold, NEC has $600
    const meetsThreshold = r.value > 0 && r.value >= threshold;
    displayReq = meetsThreshold ? 'YES' : '';
}
```

## Test Coverage
Created `tests/test_1099_threshold.js` to verify:
- ✅ NEC vendor with $700 → Shows "Required: YES"
- ✅ NEC vendor with $500 → Shows type but NO "Required: YES"
- ✅ INT vendor with $100 → Shows "Required: YES"
- ✅ INT vendor with $0 → Shows type but NO "Required: YES"

## Impact
- Vendor report now accurately reflects which vendors **qualify** for 1099 reporting
- CSV generation already had correct threshold logic, so no changes needed there
- Backward compatible with all existing functionality

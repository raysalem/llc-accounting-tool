# Plan: Fix Required Column Display

## Problem
The "Required" column is empty for Peter Wienold (and likely others) even though:
- The "1099 Type" column shows "NEC" correctly
- The user states "peter should have a 1099 report requirement"

## Analysis
Looking at the vendor report output:
```
Peter Wienold                        3,000.00  NEC
```

The "Required" column (3rd column after Total and 1099 Type) is blank.

## Root Cause Investigation
From the earlier debugging (Step 742), we saw:
- The user's Setup sheet has a **legacy `1099` column** (column 10)
- It does NOT have separate `1099 Type` and `1099 Required` columns
- The legacy column contains values like 'NEC' and 'INT'

The current logic (from Step 290):
```javascript
vendor1099Map.set(lowerV, { type, req });
```

Where:
- `type` comes from the legacy column ('NEC')
- `req` comes from `col1099Req` which is **undefined** (column doesn't exist)
- So `req` is always `''` (empty string)

## Display Logic
In the vendor report printer (around line 872):
```javascript
const info = vendor1099Map.get(r.label.toLowerCase()) || { type: '', req: '' };
console.log(
    `${r.label...}` +
    `${r.value...}` +
    `  ${(info.type || '').padEnd(12)}` +
    `${(info.req || '').padEnd(10)}`
);
```

Since `info.req` is `''`, the Required column prints blank.

## Solution
The user expects that if a vendor has a 1099 Type (NEC/INT), the "Required" column should show "YES" (or similar indicator).

**Options:**
1. **Implicit Logic**: If `type` is set (NEC/INT), default `req` to 'YES'
2. **Explicit Check**: Only if the legacy column says 'YES', set req to 'YES'
3. **User Clarification**: Ask what the legacy column value means

Given the context (legacy column has 'NEC'/'INT' values), the most logical interpretation is:
- If the legacy column has 'NEC' or 'INT', it means **both** Type=NEC/INT **and** Required=YES
- If the legacy column is blank, it means no 1099 requirement

## Implementation Plan
1. Update the Setup reading logic to set `req = 'YES'` when `type` is derived from the legacy column
2. Verify the output shows "YES" in the Required column
3. Run the command to confirm

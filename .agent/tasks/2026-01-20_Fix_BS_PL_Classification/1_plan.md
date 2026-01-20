# Plan: Fix BS/PL Classification Issue

## Problem
User reports: "alot of categories should be PL and not BS" when running `--bs-sub` on RMP prop.lnk file.

## Investigation
Looking at the output:
- **P&L Report** shows: Granada, CPA, management, repair, taxes, etc.
- **BS Report** shows: 39th, AX CC, Owner's Draw, Owner's Equity, Savings Account, usbank

The category "39th" appears in Balance Sheet but has P&L-type transactions:
- rent (income)
- mortgage (expense)
- property taxes (expense)
- Insurance (expense)

## Root Cause Analysis
The code reads the "Report" column from the Setup sheet (line 248):
```javascript
const report = colReport ? getVal(row.getCell(colReport)) : '';
uniqueCategories.set(lower, {
    report,  // This determines if it goes to P&L or BS
    ...
});
```

Then categories are filtered (lines 869-877):
```javascript
const pnlNames = Array.from(uniqueCategories.values())
    .filter(conf => conf.report === 'P&L')
    ...

const bsNames = Array.from(uniqueCategories.values())
    .filter(conf => conf.report === 'BS' || conf.report === 'Balance Sheet')
    ...
```

**Conclusion**: This is a **DATA ISSUE** in the user's Setup sheet, not a code bug. The Setup sheet has "39th" marked as "BS" or "Balance Sheet" when it should be marked as "P&L".

## Solution Options

### Option 1: User Fixes Setup Sheet (Recommended)
- User needs to open the Excel file
- Go to Setup tab
- Find the "39th" category row
- Change the "Report" column from "BS" to "P&L"
- Save the file

### Option 2: Add Validation/Warning
Add a checker warning that flags categories that:
- Are marked as "BS" but have mostly income/expense subcategories
- Are marked as "P&L" but have mostly asset/liability subcategories

This would help users catch misconfigurations.

### Option 3: Add Debug Output
Add a `--debug-categories` flag that shows:
- Category name
- Report type from Setup
- Number of transactions
- Sample subcategories

This would help users diagnose classification issues.

## Recommended Action
1. **Immediate**: Inform user this is a Setup sheet data issue
2. **Short-term**: Add helpful error message or warning
3. **Long-term**: Consider Option 2 (validation) to catch these issues automatically

## Verification Steps
1. Ask user to check Setup sheet "Report" column for "39th"
2. If it says "BS", change to "P&L"
3. Re-run `--bs-sub` and `--pl-sub` to verify

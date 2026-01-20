# Plan: Fix Subcategory Validation

## Problem
The subcategory "deposit" appears in the `--pl-sub` report even though it's not defined in the Setup sheet. The `--checker` flag also doesn't flag it as an illegal subcategory.

## Root Cause
Looking at `report.js` lines 600-611:
- **Category validation exists** (line 600-602): Checks if category is in `validCategories`
- **Subcategory validation is MISSING**: Line 610-611 just uses `subCatVal` without any validation

```javascript
const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';
catStats[displayCat].subCats[sName] = (catStats[displayCat].subCats[sName] || 0) + amount;
```

The code never checks if `sName` is a valid subcategory for that category.

## Expected Behavior
1. **Setup Sheet Structure**: Each category row can have a `Sub-Category` column that defines valid subcategories for that category
2. **Validation**: When processing transactions, if a subcategory is used, it should be validated against the Setup sheet
3. **Reporting**: Invalid subcategories should be flagged in `--checker` output

## Current Setup Sheet Reading
From lines 243-250, the code reads:
```javascript
const subCatVal = colSubCategory ? getVal(row.getCell(colSubCategory)) : '';
uniqueCategories.set(lower, {
    report,
    accountType: typeVal,
    subCategory: subCatVal,  // This stores ONE subcategory per category
    displayName: trimmed
});
```

**Issue**: The current structure only stores ONE subcategory per category in the Setup sheet, but in reality, a category can have MULTIPLE valid subcategories.

## Solution Options

### Option 1: Build a Set of Valid Subcategories (Recommended)
1. Create a `Set` called `validSubCategories` to track all subcategories defined in Setup
2. During Setup reading, add each subcategory to this set
3. During transaction processing, validate subcategories against this set
4. Add to `illegalSubCategories` array if not found

### Option 2: Category-Specific Subcategory Validation
1. Change `uniqueCategories` to store an array/set of valid subcategories per category
2. Validate that subcategory is valid for the specific category
3. More complex but more accurate

## Implementation Plan (Option 1 - Simpler)
1. Add `const validSubCategories = new Set();` near line 139
2. Add `const illegalSubCategories = [];` near line 145
3. In Setup reading loop (line 243), if `subCatVal` exists, add to `validSubCategories`
4. In transaction processing (after line 610), validate `sName` against `validSubCategories`
5. In checker output, report illegal subcategories

## Testing
- Run against user's file: `node report.js "...RMP prop.lnk" --checker`
- Verify "deposit" is flagged as illegal subcategory
- Verify it still shows in `--pl-sub` (data should display even if invalid)
- Create automated test with known invalid subcategory

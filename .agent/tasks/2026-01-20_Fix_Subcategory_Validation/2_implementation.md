# Implementation: Fix Subcategory Validation

## Changes Made

### 1. Added Data Structures (Lines 133-149)
```javascript
const validSubCategories = new Set(); // Set of all valid subcategories from Setup
const illegalSubCategories = [];      // Track invalid subcategories used in transactions
```

### 2. Populate Valid Subcategories During Setup Reading (Lines 243-252)
```javascript
// Track valid subcategories
if (subCatVal) {
    const subCatStr = subCatVal.toString().trim();
    if (subCatStr) {
        validSubCategories.add(subCatStr.toLowerCase());
    }
}
```

When reading the Setup sheet, any subcategory defined in the `Sub-Category` column is added to the `validSubCategories` set.

### 3. Validate Subcategories During Transaction Processing (Lines 610-626)
```javascript
const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';

// Validate subcategory if one is provided
if (subCatVal && sName !== '(No Sub-Cat)') {
    const sLower = sName.toLowerCase();
    if (!validSubCategories.has(sLower)) {
        illegalSubCategories.push({ 
            value: sName, 
            category: displayCat,
            sheet: sheet.name, 
            row: r, 
            date: displayDate 
        });
    }
}

catStats[displayCat].subCats[sName] = (catStats[displayCat].subCats[sName] || 0) + amount;
```

When processing transactions, if a subcategory is used, it's validated against the `validSubCategories` set. Invalid ones are tracked with their category context.

### 4. Report Illegal Subcategories (Lines 1036-1064)
```javascript
const hasIssues = ... || illegalSubCategories.length > 0;
...
const subCats = new Set(illegalSubCategories.filter(x => x.sheet === s).map(x => x.value));
if (subCats.size) console.log(`  [!] Illegal Sub-Categories: ${Array.from(subCats).join(', ')}`);

if (showChecker) {
    illegalSubCategories.filter(x => x.sheet === s).forEach(x => 
        console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL SUB-CATEGORY "${x.value}" in category "${x.category}"`)
    );
}
```

Added illegal subcategories to the DATA INTEGRITY ISSUES output, showing:
- Summary line with all illegal subcategories
- Detailed `--checker` output with row numbers, dates, and category context

## Design Decisions

1. **Global Validation**: Used a single `Set` for all subcategories rather than per-category validation. This is simpler and matches the current Setup sheet structure where subcategories are defined once per category row.

2. **Display Regardless**: Invalid subcategories still appear in reports (like `--pl-sub`). The validation only flags them as issues - it doesn't hide the data.

3. **Category Context**: When reporting illegal subcategories, we include which category they were used with, making it easier for users to understand the issue.

## Impact
- Subcategories are now validated just like categories, vendors, and customers
- `--checker` flag now catches illegal subcategories
- No change to report display logic - data still shows even if invalid
- Backward compatible with existing workbooks

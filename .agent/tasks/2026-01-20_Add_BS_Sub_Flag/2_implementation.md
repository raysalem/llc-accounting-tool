# Implementation: Add --bs-sub Flag

## Changes Made

### 1. Added Flag Definition (Line 38)
```javascript
const showBSSub = args.includes('--bs-sub');
```

### 2. Updated Help Menu (Line 66)
```
--bs-sub        (Optional) Print detailed Balance Sheet with sub-category breakdowns.
```

### 3. Added to Known Flags (Line 77)
```javascript
const knownFlags = [
    '--save', '--pl', '--bs', '--vendor', '--customer', '--pl-sub', '--bs-sub', ...
];
```

### 4. Updated Specific Filter (Line 87)
```javascript
const specificFilter = ... || showBSSub || ...;
```

### 5. Implemented BS-Sub Display Logic (Lines 917-937)
```javascript
if (showAll || showBS || showBSSub) {
    console.log(`\n--- BALANCE SHEET ---`);
    if (!reports.bs.length) console.log('(No Data)');
    else {
        const max = Math.max(...reports.bs.map(r => r.label.length), 10);
        reports.bs.forEach(r => {
            console.log(`${r.label.padEnd(max + 5)} : ${r.value...}`);
            // Sub-Category Detail
            if (showBSSub && catStats[r.label] && catStats[r.label].subCats) {
                const subs = catStats[r.label].subCats;
                const subKeys = Object.keys(subs).filter(k => Math.abs(subs[k]) > 0.01);
                if (!(subKeys.length === 1 && subKeys[0] === '(No Sub-Cat)')) {
                    subKeys.sort().forEach(sub => {
                        console.log(`   > ${sub.padEnd(max + 1)} : ${subs[sub]...}`);
                    });
                }
            }
        });
    }
}
```

## Design Pattern
Followed the exact same pattern as `--pl-sub`:
1. Check if flag is set
2. Display main category with total
3. If `showBSSub` is true, display subcategories indented with ">"
4. Filter out subcategories with near-zero values (< 0.01)
5. Skip display if only subcategory is "(No Sub-Cat)"

## Example Output
```
--- BALANCE SHEET ---
39th                 :      -18,561.44
   > 3rd-party-parking :       -1,200.00
   > Insurance        :       -2,958.00
   > deposit          :       -2,500.00
   > mortgage         :      -42,200.00
   > rent             :       90,970.00
```

## Impact
- Users can now see Balance Sheet subcategory breakdowns
- Consistent with existing `--pl-sub` functionality
- No changes to data processing, only display logic

# Testing: Add --bs-sub Flag

## Test 1: User's Actual Data
**Command:** `node report.js "\\192.168.1.90\Documents Private\taxes\2025\RMP prop.lnk" --bs-sub`

**Result:** ✅ SUCCESS
```
--- BALANCE SHEET ---
39th                 :      -18,561.44
   > 3rd-party-parking :       -1,200.00
   > Insurance        :       -2,958.00
   > Utilities-water  :       -1,580.00
   > deposit          :       -2,500.00
   > mortgage         :      -42,200.00
   > pest             :       -1,190.00
   > property taxes   :      -19,280.56
   > rent             :       90,970.00
   > repair           :       -1,500.00
AX CC                :       -3,421.87
Owner's Draw         :            0.00
Owner's Equity       :       12,300.00
Savings Account      :           -0.00
usbank               :        7,265.32
```

The Balance Sheet now shows subcategory breakdowns, just like `--pl-sub`.

## Test 2: Integration Test
**Command:** `node tests/run_integration_test.js`

**Result:** ✅ PASSED
- All phases completed successfully
- No regressions detected
- Balance Sheet displays correctly

## Test 3: Help Menu
**Command:** `node report.js --help`

**Result:** ✅ SUCCESS
- `--bs-sub` appears in help menu
- Description is clear and consistent with `--pl-sub`

## Verification Summary
- ✅ `--bs-sub` flag works correctly
- ✅ Subcategories display with proper indentation
- ✅ Follows same pattern as `--pl-sub`
- ✅ No regressions in existing functionality
- ✅ Integration test passes

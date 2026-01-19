# Testing: Rename 1099 Argument

## Strategy
1.  **Unit/Integration Test**: Use `tests/test_arguments_coverage.js` to explicitly invoke `report.js` with the new `--1099` flag.
2.  **Verification**: ensure the script exits with code 0 (success) and produces output without error.
3.  **Regression**: `run_integration_test.js` (run previously) ensures that the core `report.js` logic remains solid.

## Execution
Ran `node tests/test_arguments_coverage.js`.

### Output
```
--- TEST SUITE: Argument Coverage ---

[1/4] load_transactions.js (Excel Input, Append, Clear)
   Running: Load from Excel (Append Mode)...
   Running: Load from Excel (Append Mode / 2nd Run)...
   Running: Load with --clear...

[2/4] report.js (Flags)
   Running: --vendor
   Running: --customer
   Running: --pl-sub
   Running: --save
Testing --1099...

✅ TEST PASSED: All arguments covered and verified.
```

## Regression
Existing functionality is preserved as the old `--1099-nec` flag is still supported internally as an alias, though not advertised.

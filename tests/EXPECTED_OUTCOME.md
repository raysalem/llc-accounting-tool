# Integration Test Overview

`tests/run_integration_test.js` checks the whole pipeline: template creation, CSV loading, categorization, ledger entries, reporting, the tax-year filter and the integrity checker. It asserts on the report output and exits non-zero if any check fails.

Run every test with:
```bash
npm test
```

## Test Components

- `example_bank.csv`: salary deposit, consulting income and a rent payment.
- `example_cc.csv`: coffee, office supplies and cloud services on a credit card.
- `run_integration_test.js`: the test runner.

## Scenario

1. **Template**: `generate_excel.js` creates the workbook. Transaction sheets have TOTAL/SUBTOTAL in rows 1–2 and the header in row 3.
2. **Load**: the bank and CC CSVs are loaded with `--clear`.
3. **Categorize**: rows are categorized, and a $737.50 payment to an NEC vendor (`Contractor 1099`) is added.
4. **Ledger** (balanced double entry):
   - Owner investment: Dr Checking Account 1,000 / Cr Owner Equity 1,000
   - Audit adjustment: Dr Office 50 / Cr Checking Account 50

## Expected Results (`--year=2025`)

| Check | Math | Expected |
|---|---|---|
| Sales | 5,000 + 2,500 | 7,500.00 |
| Office | 120 + 45 (CC) + 50 (ledger) | 215.00 expense |
| Net income | 7,500 − 1,500 − 737.50 − 15.50 − 215 | 5,032.00 |
| Bank balance | 5,000 − 1,500 + 2,500 − 737.50 + 1,000 − 50 | 6,212.50 |
| CC liability | 15.50 + 120 + 45 | 180.50 |
| Balance sheet | 6,212.50 = 180.50 + 1,000 + 5,032 | `[OK] (A = L + E)` |
| Clean run exit code | | 0 |

## Other Checks

- **1099**: `Contractor 1099` is over the 2025 $600 NEC threshold and is listed. It has no TIN or address, so the run reports "Incomplete Data" and exits non-zero.
- **Tax year**: with `--year=2026` the header shows the $2,000 NEC threshold, all 11 rows from 2025 (7 transactions + 4 ledger rows) are skipped, and net income is 0.00.
- **Integrity checker**: after adding a row with an unknown category (`IllegalCat`) and an uncategorized row, `--checker` flags both and exits non-zero.

## Test Artifact
The test saves `tests/Full_Accounting_Test_Case.xlsx` with the full setup, transactions and ledger entries.

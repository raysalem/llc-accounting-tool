# Test Plan

Run everything with:
```bash
npm test
```
`tests/run_all.js` runs every test file and exits non-zero if any fail. CI runs it on every push and pull request to `main` (Node 20 and 22).

Tests write their workbooks to a temporary directory, so a test run never changes files in the repo. To refresh the committed example `tests/Full_Accounting_Test_Case.xlsx`, run `SAVE_ARTIFACT=1 node tests/run_integration_test.js`.

## Coverage Map

| Area | Scenario | Test file |
|---|---|---|
| Parsing | Amount formats (`$`, commas, `(x)`, `x-`, `CR`), tax-year parsing, dates, 1099 thresholds | `test_accounting_lib.js` |
| Loading | CSV amounts and Dec 31 / Jan 1 dates land correctly in UTC, Los Angeles and Tokyo time zones | `test_loader_formats.js` |
| End to end | Template → load → categorize → ledger → P&L / BS with exact totals | `run_integration_test.js` |
| Tax year | `--year` skips other years; 2026 uses the $2,000 threshold | `run_integration_test.js`, `test_loader_formats.js` |
| Transfers | CC payment counted once; refunds; transfer category on the wrong sheet | `test_transfers.js` |
| Ledger | Missing date stops the run; unbalanced ledger stops the run | `test_ledger_strictness.js`, `test_ledger_unbalanced.js` |
| Balance sheet | Asset polarity for ledger entries | `test_asset_polarity.js` |
| 1099 | Required flag at / under the threshold, after refunds, 2025 vs 2026 | `test_1099_boundaries.js`, `test_1099_threshold.js` |
| 1099 | Missing vendor TIN/address flagged; `1099 Data` sheet amount | `run_integration_test.js`, `test_comprehensive_report.js` |
| Checker | Unknown category, uncategorized rows, non-zero exit | `run_integration_test.js`, `test_comprehensive_report.js` |
| Reports | `--pl-sub`, `--bs-sub`, `--vendor-sub`, `--customer-sub`, `--details` | `test_comprehensive_report.js`, `test_details_extended.js`, `test_pl_sub_display.js` |
| Reports | `--details` includes Ledger rows and doesn't change totals | `test_details_ledger.js` |
| CLI | `load_transactions.js` append / `--clear` / `--help`; `report.js` flags and `--save` | `test_arguments_coverage.js` |

## Integration Test (`run_integration_test.js`)

**Scenario**
1. `generate_excel.js` creates the workbook. Transaction sheets have TOTAL/SUBTOTAL in rows 1–2 and the header in row 3.
2. `tests/example_bank.csv` and `tests/example_cc.csv` are loaded with `--clear`.
3. Rows are categorized, and a $737.50 payment to an NEC vendor (`Contractor 1099`) is added.
4. Balanced ledger entries are added:
   - Owner investment: Dr Checking Account 1,000 / Cr Owner Equity 1,000
   - Audit adjustment: Dr Office 50 / Cr Checking Account 50

**Expected results (`--year=2025`)**

| Check | Math | Expected |
|---|---|---|
| Sales | 5,000 + 2,500 | 7,500.00 |
| Office | 120 + 45 (CC) + 50 (ledger) | 215.00 expense |
| Net income | 7,500 − 1,500 − 737.50 − 15.50 − 215 | 5,032.00 |
| Bank balance | 5,000 − 1,500 + 2,500 − 737.50 + 1,000 − 50 | 6,212.50 |
| CC liability | 15.50 + 120 + 45 | 180.50 |
| Balance sheet | 6,212.50 = 180.50 + 1,000 + 5,032 | `[OK] (A = L + E)` |
| Exit code | | 0 |

**Other checks**
- **1099**: `Contractor 1099` is over the 2025 $600 threshold and is listed. It has no TIN or address, so the run reports "Incomplete Data" and exits non-zero.
- **Tax year**: `--year=2026` shows the $2,000 threshold, skips all 11 rows from 2025 (7 transactions + 4 ledger rows), and net income is 0.00.
- **Checker**: after adding a row with an unknown category (`IllegalCat`) and an uncategorized row, `--checker` flags both and exits non-zero.

## Transfers and Refunds (`test_transfers.js`)

**Setup**: bank sheet linked to `Checking` (asset). Amex sheet linked to `AX CC` (liability), with Flip = Yes because the statement shows charges as positive. `Transfer AX CC` has Report = `Transfer` and Transfer Account = `AX CC`.

| Sheet | Row | Category |
|---|---|---|
| Bank | +1,000 client payment | Sales |
| Bank | −100 Amex autopay | AX CC |
| Amex | 300 office supplies | Office Supplies |
| Amex | −100 payment received | Transfer AX CC |
| Amex | −50 refund | Office Supplies |

**Expected**
- **Checking**: 1,000 − 100 = 900.00
- **AX CC**: 300 − 50 (refund) − 100 (bank payment only) = 150.00. If the Amex-side payment were also counted, it would show 50.00.
- **Office Supplies**: −250.00, because the refund reduces the expense.
- **Net income**: 750.00, and the balance sheet balances (900 = 150 + 750).
- **Wrong sheet**: using `Transfer AX CC` on the Bank sheet prints `CRITICAL WARNING … expects Transfer Account "AX CC" … Sheet "Bank"`.

## Known Gaps

- **Unbalanced sample books**: several feature tests (`test_1099_threshold`, `test_details_extended`, `test_pl_sub_display`, `test_arguments_coverage`) use sample books that don't balance, so they ignore `report.js`'s exit code. Giving them balanced books would let them assert exit 0.
- **Word-only checks**: some older checks only look for a word in the output. New tests should assert the number on the same line.
- **Warnings don't change the exit code**: a `CRITICAL WARNING` (e.g. a transfer category on the wrong sheet) doesn't make `report.js` exit non-zero. Only errors and integrity issues do.

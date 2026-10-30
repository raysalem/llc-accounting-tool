# Test Plan

Run everything with:
```bash
npm test
```
`tests/run_all.js` runs every test file and exits non-zero if any fail. CI runs it on every push and pull request to `main` (Node 20 and 22).

All sample books balance and every Bank/CC sheet carries the correct End Balance (computed from the test's own data, with a comment showing the math), so tests expect `report.js` to exit 0 unless they are testing an error. A test that runs the same book for a different year or without `--opening-balances` does not reconcile and does not assert the exit code. Tests write their workbooks to a temporary directory, so a test run never changes files in the repo. To refresh the committed example `tests/Full_Accounting_Test_Case.xlsx`, run `SAVE_ARTIFACT=1 node tests/run_integration_test.js`.

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
| Opening balances | TOML file validation, Setup-stage checks, and exact balance-sheet numbers with `--opening-balances` | `test_opening_balances.js` |
| 1099 | Required flag at / under the threshold, after refunds, 2025 vs 2026 | `test_1099_boundaries.js`, `test_1099_threshold.js` |
| 1099 | Missing vendor TIN/address flagged; `1099 Data` sheet amount | `run_integration_test.js`, `test_comprehensive_report.js` |
| Checker | Unknown category, uncategorized rows, non-zero exit | `run_integration_test.js`, `test_comprehensive_report.js` |
| Reports | `--pl-sub`, `--bs-sub`, `--vendor-sub`, `--customer-sub`, `--details` | `test_comprehensive_report.js`, `test_details_extended.js`, `test_pl_sub_display.js` |
| Reports | `--details` includes Ledger rows and doesn't change totals | `test_details_ledger.js` |
| CLI | `load_transactions.js` append / `--clear` / `--help`; `report.js` flags and `--save` | `test_arguments_coverage.js` |
| CLI | `--vendor-file`, `vendor.csv` next to the workbook, `--ignore-vendors`, `--all`, `--debug` | `test_vendor_file.js` |
| Reconciliation | End Balance required for Bank/CC sheets: all match, one missing, explicit 0, mismatch, mixed counts in the summary line | `test_reconciliation.js` |
| Report stamp | Version, git commit, run time, flags, tax year and workbook SHA-256/size on the console and in the `--save` Report Info sheet | `test_reconciliation.js` |
| Book file | Validation (every problem listed with a fix), registry and path lookup, default_flags / --year / workbook merge, the example files with `generate_excel.js --book` | `test_book_config.js` |
| Book file | Same reports as the Setup sheet (`--all`, `--1099`), tax IDs never on the console and present in the saved `1099 Data` sheet, private file handling, `--opening-balances` override | `test_book_equivalence.js` |
| Book file | `load_transactions.js --book`: lock check first, Lists sheet, dropdowns (Sub-Category by Category) point at Lists, a second import lands right under the first, `--refresh-lists`, legacy import unchanged | `test_book_loader.js` |
| Tax year check | A book row dated in another year is listed and fails the run (exit 1); reports still print; all-in-year summary | `test_tax_year_check.js` |
| Unbalanced ledger | Reports still print with the error, exit 1; `--checker` stops before the reports | `test_ledger_unbalanced.js` |
| Setup | No sheet configuration in Setup: default sheets are used and their header row is found | `test_fallback_sheets.js` |

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

## Opening Balances (`test_opening_balances.js`)

**Setup**: `Bank` (asset, linked to the Bank sheet, no StartBalance), `Mortgage` (unlinked liability), `Equipment` (unlinked asset), `Owner Investment` (equity). Bank sheet, 2026: −500 to `Mortgage`, −400 to `Equipment` (plus a 2025 row that is skipped). File (closing year 2025): Bank 1,000, Mortgage 3,433, Equipment 4,345, Owner Investment 343. Run with `--bs --year=2026`.

| Check | Math | Expected |
|---|---|---|
| Bank | 1,000 − 500 − 400 | 100.00 |
| Mortgage | 3,433 − 500 | 2,933.00 |
| Equipment | 4,345 + 400 | 4,745.00 |
| Retained Earnings (Opening Balance) | 1,000 + 4,345 − 3,433 − 343 | 1,569.00 |
| Balance sheet | 4,845 = 2,933 + 1,912 | `[OK] (A = L + E)` |

The test also covers each validation rule (duplicate key, `[acounts]` typo, missing or 3-digit year, bad `as_of_date`, empty `[accounts]`, string or 3-decimal value, nested table), the Setup-stage checks (unknown or P&L category, two keys for one category, wrong year, conflicting or equal StartBalance) and that a bad file exits 1 with every problem listed.

## Book Equivalence (`test_book_equivalence.js`)

**Setup**: one synthetic book (year 2026) built twice: as a workbook with Setup tables (header-row detection) and as a TOML book (`default_flags = ["--all"]`, `opening_balances`) with a workbook that has no Setup sheet, and again with a stale Setup sheet that must be ignored. It has income, expense, asset, liability, equity and transfer categories with sub-categories, a bank sheet linked to Checking, a card sheet (flipped) linked to Credit Card, a Ledger, an unlinked asset (Equipment), a 1099-NEC vendor, a customer, and opening balances. The run compares `--all` (and `--1099`) stdout and stderr, exit code included, after removing the stamp lines, the "Setup sheet ignored" line, the Setup-table warning, the payer-source line and the book-mode 1099 hint.

| Check | Math | Expected |
|---|---|---|
| Checking | 1,000 opening + (5,000 − 2,500 − 300 − 100 − 400) + 500 ledger | 3,200.00 |
| Credit Card | 200 + 150 charge − 100 bank autopay (the card-side Transfer payment is left out) | 250.00 |
| Equipment (unlinked asset) | 600 + 400 | 1,000.00 |
| Owner Investment | 1,000 + 500 | 1,500.00 |
| Net income | 5,000 − 2,500 − (300 + 150) | 2,050.00 |
| Retained Earnings (Opening Balance) | 1,000 + 600 − 200 − 1,000 | 400.00 |
| Balance sheet | 4,200 = 250 + 3,950 | `[OK] (A = L + E)` |

**Tax IDs** (fake: `00-0000000`, `00-0000001`): with a private file and complete vendor details, `--1099`, `--all`, `--debug --checker`, `--details` and `--save` runs never print either ID to stdout or stderr; `--save` writes both in `1099 Data` and in no other sheet. Without the private file the run reports "Missing: Tax ID (SSN/EIN)" and exits 1.

**Lock test** (`test_book_loader.js`): on Windows a PowerShell helper opens the workbook with sharing denied (what Excel does) and the test checks exit 1, the message, that the missing export was never looked at, and that the file is byte-identical. Elsewhere only the mocked-`fs` unit test of `lib/book/lock.js` runs (EBUSY, EPERM, EACCES, ENOENT, other errors).

## Known Gaps

- **`--save` exit code**: `--save` exits 1 whenever any warning was printed (even harmless ones such as "no vendor.xlsx found"), unlike runs without `--save`. `test_arguments_coverage.js` checks the saved file instead of the exit code for that step.
- **Word-only checks**: some older checks only look for a word in the output. New tests should assert the number on the same line.
- **Warnings don't change the exit code**: a `CRITICAL WARNING` (e.g. a transfer category on the wrong sheet) doesn't make `report.js` exit non-zero. Only errors and integrity issues do.

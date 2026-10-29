# Features and Test Coverage

Every feature listed here should have a test. Gaps are marked ⚠️ and tracked in `todo.md`.

## Commands

| Command | What it does |
|---|---|
| `node generate_excel.js [file]` | Creates a starting workbook: Setup, Bank Transactions, Credit Card Transactions, Ledger, Summary and VERSION sheets. |
| `node load_transactions.js <input> <bank\|cc> <workbook> [--clear]` | Imports a CSV or Excel export. Detects columns, reads amounts like `$1,234.56`, `(50.00)` and `1,000.00-`, and adds TOTAL/SUBTOTAL rows above the header. `--clear` deletes and recreates the target sheet first. |
| `node report.js <workbook> [flags]` | Builds and prints the reports (flags below). Accepts `.xlsx`, `.lnk` and `.url` paths. |
| `node scripts/batch_run.js "<glob>" [flags]` | Runs `report.js` on every matching workbook. |
| `node scripts/inspect.js <workbook>` | Prints a workbook's sheet configuration and headers. |

## `report.js` Flags

| Flag | Description | Test |
|---|---|---|
| `--year=YYYY` | Tax year to report. Rows dated in other years are skipped. Sets the 1099-NEC threshold ($600 before 2026, $2,000 from 2026). Default: `DEFAULT_TAX_YEAR` in `lib/accounting.js`. | `run_integration_test.js`, `test_loader_formats.js`, `test_1099_boundaries.js` |
| `--pl` | Profit & Loss summary. | `test_comprehensive_report.js`, `run_integration_test.js` |
| `--bs` | Balance sheet summary, with the A = L + E check. | `run_integration_test.js`, `test_transfers.js` |
| `--pl-sub` | Detailed P&L with sub-categories and a column per sheet. | `test_pl_sub_display.js`, `test_comprehensive_report.js` |
| `--bs-sub` | Detailed balance sheet. | `test_comprehensive_report.js` |
| `--vendor` | Vendor spending, with 1099 type and whether a 1099 is required. | `test_1099_threshold.js`, `test_1099_boundaries.js` |
| `--vendor-sub` | Detailed vendor spending. | `test_comprehensive_report.js` |
| `--customer` / `--customer-sub` | Customer income (summary / detailed). | `test_comprehensive_report.js` |
| `--1099`, `--1099=NEC`, `--1099=INT` | 1099 preparation: payer info, recipients over the threshold, missing details. | `run_integration_test.js`, `test_comprehensive_report.js` |
| `--details "Name"` | Every row for a category, vendor or customer, including Ledger rows. | `test_details_extended.js`, `test_details_ledger.js`, `test_arguments_coverage.js` |
| `--checker` | Data-integrity report: missing categories, unknown categories/vendors/customers, header problems. | `run_integration_test.js`, `test_comprehensive_report.js` |
| `--save` | Writes `report_<name>.xlsx` (one sheet per report, 1099 Data, Processing Log, Report Info) and a PDF next to the workbook. Does not modify the workbook. | `test_arguments_coverage.js`, `test_comprehensive_report.js` |
| `--all` | All reports. | `test_vendor_file.js` |
| `--vendor-file <path>` | Use a specific `vendor.xlsx` / `vendor.csv` for 1099 details. Without it, `vendor.xlsx` / `vendor.csv` next to the workbook (or in the current directory) is used. | `test_vendor_file.js` |
| `--opening-balances <path>` | Start the year from last year's closing balance-sheet balances in a TOML file (`year`, optional `as_of_date`, `[accounts]` keyed by Setup category). Strictly validated; any problem stops the run with exit 1. See `examples/opening-balances.example.toml`. | `test_opening_balances.js` |
| `--ignore-vendors` | Don't load `vendor.xlsx` / `vendor.csv`. | `test_vendor_file.js` |
| `--debug` | Verbose output. | `test_vendor_file.js` |
| `--help` | Usage. | `test_arguments_coverage.js` |

**Exit code:** `report.js` exits 1 on critical errors (unbalanced ledger or balance sheet, a ledger row without a date, incomplete 1099 vendor details, a row that fails to process, a missing or mismatched End Balance) or data-integrity issues. Otherwise 0.

## Reconciliation and Report Stamp

| Feature | Description | Test |
|---|---|---|
| End Balance required | Every Bank/CC sheet (and any sheet linked to a Balance Sheet account) needs the statement's ending balance in `Setup` (SheetInfo). The calculated balance must match it within 0.05. A missing End Balance prints `UNRECONCILED: "<sheet>" has no End Balance in Setup` and exits 1; an explicit `0` is valid (empty and 0 are different); a mismatch prints `END BALANCE MISMATCH` and exits 1. All problems are listed together. The Ledger sheet is exempt, and so are the default sheets used when Setup has no sheet configuration. | `test_reconciliation.js` |
| Reconciliation summary | One line on every run: `Reconciliation: 3 of 3 accounts match their End Balance` or `Reconciliation: 1 of 3 accounts match (1 mismatch, 1 missing End Balance)`. | `test_reconciliation.js` |
| Report stamp | What the report was made from: tool version, git commit (`+dirty` if the tool's checkout has uncommitted changes, `unknown` if git is unavailable), run time (ISO, local time zone), the command-line arguments as given, tax year, and the SHA-256 (first 12 hex digits) and size of every input file read (workbook, `--opening-balances` file, vendor file). Built once in `lib/report/stamp.js`; printed under `Report stamp:` at the top of every run, added to the `--save` Report Info sheet and printed on the PDF title page. | `test_reconciliation.js` |

## Accounting Rules

- **Polarity:** money into an asset is positive, money out is negative. Sheets with Flip Polarity = Yes are multiplied by −1 first.
- **Transfers:** rows categorized with a `Transfer` category are left out of the linked account's balance, so a credit-card payment recorded on both the bank and card sheets is counted once. See `docs/accounting-rules.md` Part 5. (`test_transfers.js`)
- **Statement balances:** each bank/card sheet is proven against its statement's ending balance (End Balance), see above. (`test_reconciliation.js`)
- **Ledger:** debits increase assets and expenses; credits increase income, liabilities and equity. Debits must equal credits and every row needs a date, or the run stops. (`test_ledger_strictness.js`, `test_ledger_unbalanced.js`)
- **Category validation:** P&L categories must be `Income` or `Expense`; balance-sheet categories must be `Asset`, `Liability` or `Equity`. Violations print `CRITICAL CONFIG ERROR`.
- **Asset polarity:** linked assets (bank sheets) use natural polarity. Unlinked assets (purchases) are inverted: spending increases the asset. (`test_asset_polarity.js`)
- **Opening balances:** with `--opening-balances`, an account linked to a sheet gets the file's amount as its sheet's starting balance; an unlinked account has it added after the asset polarity flip. Assets add to, and liabilities and equity subtract from, "Retained Earnings (Opening Balance)" so A = L + E holds. Positive = the account's normal balance. (`test_opening_balances.js`)
- **Detailed reports:** the Additions and Subtractions columns add magnitudes (`Math.abs`) so mixed signs don't cancel out.
- **1099 thresholds:** NEC/MISC use the year's threshold; INT uses $0 (all interest is listed). (`test_accounting_lib.js`, `test_1099_boundaries.js`)

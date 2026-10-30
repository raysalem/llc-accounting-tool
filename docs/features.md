# Features and Test Coverage

Every feature listed here should have a test. Gaps are marked ⚠️ and tracked in `todo.md`.

## Commands

| Command | What it does |
|---|---|
| `node generate_excel.js [file]` | Creates a starting workbook: Setup, Bank Transactions, Credit Card Transactions, Ledger, Summary and VERSION sheets. |
| `node generate_excel.js --book <id\|path> [file]` | Creates a starting workbook for a book file: one sheet per `[[sheets]]` entry, a Ledger, VERSION and the generated Lists sheet; no Setup sheet. Refuses to overwrite an existing file. |
| `node load_transactions.js <input> <bank\|cc> <workbook> [--clear]` | Imports a CSV or Excel export. Detects columns, reads amounts like `$1,234.56`, `(50.00)` and `1,000.00-`, and adds TOTAL/SUBTOTAL rows above the header. `--clear` deletes and recreates the target sheet first. |
| `node load_transactions.js --book <id\|path> <input> <sheet> [--clear]` | Same import, configured by a book file: checks the workbook is not open in Excel first, writes the rows, rewrites the generated Lists sheet and points the dropdowns at it. |
| `node load_transactions.js --book <id\|path> --refresh-lists` | Rewrites the Lists sheet from the book and the dropdowns on every book sheet (transactions and ledger). Also repairs dropdowns lost to a paste in Excel. Does not touch the rows. |
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
| `--book <id\|path>` | Read ALL configuration (company, sheets, categories, vendors, customers, default flags, default year) from a book file instead of the Setup sheet. The book names the workbook. `--year` overrides the book's year; `default_flags` are added to your flags. See the Book File section below and `docs/setup.md`. | `test_book_config.js`, `test_book_equivalence.js` |
| `--ignore-vendors` | Don't load `vendor.xlsx` / `vendor.csv`. | `test_vendor_file.js` |
| `--debug` | Verbose output. | `test_vendor_file.js` |
| `--help` | Usage. | `test_arguments_coverage.js` |

**Tax year check:** every run ends with `Tax year check: all N rows are dated YYYY`. Rows dated in another year are listed (sheet, row, date, description, amount) and left out of the reports; with `--book` they are an error (exit 1), since a book holds one tax year. (`test_tax_year_check.js`, `run_integration_test.js`)

**Ledger (book workbooks):** an `Entry ID` column comes first, a label for grouping the lines of one entry; the report ignores it. There is no Type column: the Category's `type` in the book decides P&L vs Balance Sheet. `blank_sheets` in the book adds empty tabs (e.g. `Worksheet`) that the report ignores.

**Exit code:** `report.js` exits 1 on critical errors (unbalanced ledger or balance sheet, a ledger row without a date, incomplete 1099 vendor details, a row that fails to process, a missing or mismatched End Balance) or data-integrity issues. Otherwise 0.

## Reconciliation and Report Stamp

| Feature | Description | Test |
|---|---|---|
| End Balance required | Every Bank/CC sheet (and any sheet linked to a Balance Sheet account) needs the statement's ending balance in `Setup` (SheetInfo). The calculated balance must match it within 0.05. A missing End Balance prints `UNRECONCILED: "<sheet>" has no End Balance in Setup` and exits 1; an explicit `0` is valid (empty and 0 are different); a mismatch prints `END BALANCE MISMATCH` and exits 1. All problems are listed together. The Ledger sheet is exempt, and so are the default sheets used when Setup has no sheet configuration. | `test_reconciliation.js` |
| Reconciliation summary | One line on every run: `Reconciliation: 3 of 3 accounts match their End Balance` or `Reconciliation: 1 of 3 accounts match (1 mismatch, 1 missing End Balance)`. | `test_reconciliation.js` |
| Report stamp | What the report was made from: tool version, git commit (`+dirty` if the tool's checkout has uncommitted changes, `unknown` if git is unavailable), run time (ISO, local time zone), the command-line arguments as given, tax year, and the SHA-256 (first 12 hex digits) and size of every input file read (workbook, `--opening-balances` file, vendor file). Built once in `lib/report/stamp.js`; printed under `Report stamp:` at the top of every run, added to the `--save` Report Info sheet and printed on the PDF title page. | `test_reconciliation.js` |

## Book File (`--book`)

One TOML file per book replaces the Setup sheet (schema in `docs/setup.md`). Without `--book` nothing changes.

| Feature | Description | Test |
|---|---|---|
| Strict validation | Unknown keys, bad types, a sheet missing from the workbook, a bad `link`, duplicate names, a vendor or customer named like a category, a missing `end_balance`, and TOML errors (duplicate keys) are all listed at once, each with a suggested fix (closest valid name). Exit 1 before the workbook is opened. | `test_book_config.js` |
| Finding a book | By path (`.toml`), by id in `books.toml` (current folder) or the file named by `LLC_BOOKS`; relative paths inside a book resolve against its folder. | `test_book_config.js` |
| Same reports as Setup | The book fills the same structures the Setup sheet does (categories, transfers, sheets, linked accounts, ledger, vendors, 1099, payer, opening balances): `--all` and `--1099` output is identical to the Setup version of the same book, apart from the stamp lines. | `test_book_equivalence.js` |
| Defaults | `default_flags` merge with the command line; `year` is the default `--year` (the command line wins); `opening_balances` feeds `--opening-balances` (the command line wins). | `test_book_config.js`, `test_book_equivalence.js` |
| Tax IDs | Kept in a separate private file, read only for 1099 reports, never printed to the console or put in an error message; only the saved `1099 Data` sheet has them. A missing one is reported as missing. | `test_book_equivalence.js`, `test_book_config.js` |
| Report stamp | Lists the book, the private file (name, hash and size only), the opening-balances file and the workbook. | `test_book_equivalence.js`, `test_book_config.js` |
| Lock check | `load_transactions.js --book` opens the workbook for update before reading the export; Excel holding it gives exit 1 and "close it in Excel". | `test_book_loader.js` |
| Lists sheet | Generated from the book (Categories, SubCategories, Vendors, Customers), marked as generated, protected; the dropdowns and Report Type formula point at it. Category, Vendor and Customer dropdowns list exactly the book's entries; the Sub-Category dropdown lists only the sub-categories of the row's Category (all of them while Category is empty). Dropdowns are range rules reaching 1000 rows past the data, so the next import still lands right under the data. | `test_book_loader.js` |
| Example files | `examples/example.book.toml` and `examples/example.private.toml` work with `generate_excel.js --book`. | `test_book_config.js` |

## Accounting Rules

- **Polarity:** money into an asset is positive, money out is negative. Sheets with Flip Polarity = Yes are multiplied by −1 first.
- **Transfers:** rows categorized with a `Transfer` category are left out of the linked account's balance, so a credit-card payment recorded on both the bank and card sheets is counted once. See `docs/accounting-rules.md` Part 5. (`test_transfers.js`)
- **Statement balances:** each bank/card sheet is proven against its statement's ending balance (End Balance), see above. (`test_reconciliation.js`)
- **Ledger:** debits increase assets and expenses; credits increase income, liabilities and equity. Every row needs a date, or the run stops. Debits must equal credits: an unbalanced ledger is an error (exit 1); the reports still print so the entries can be checked, and the imbalance is repeated in the final summary. With `--checker` the run stops before the reports instead. (`test_ledger_strictness.js`, `test_ledger_unbalanced.js`)
- **Category validation:** P&L categories must be `Income` or `Expense`; balance-sheet categories must be `Asset`, `Liability` or `Equity`. Violations print `CRITICAL CONFIG ERROR`.
- **Asset polarity:** linked assets (bank sheets) use natural polarity. Unlinked assets (purchases) are inverted: spending increases the asset. (`test_asset_polarity.js`)
- **Opening balances:** with `--opening-balances`, an account linked to a sheet gets the file's amount as its sheet's starting balance; an unlinked account has it added after the asset polarity flip. Assets add to, and liabilities and equity subtract from, "Retained Earnings (Opening Balance)" so A = L + E holds. Positive = the account's normal balance. (`test_opening_balances.js`)
- **Detailed reports:** the Additions and Subtractions columns add magnitudes (`Math.abs`) so mixed signs don't cancel out.
- **1099 thresholds:** NEC/MISC use the year's threshold; INT uses $0 (all interest is listed). (`test_accounting_lib.js`, `test_1099_boundaries.js`)

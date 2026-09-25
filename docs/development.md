# Development Guide

## Requirements

- **Node.js 20 or newer** (CI tests Node 20 and 22).
- **Microsoft Excel** (or another `.xlsx` editor) to view and edit workbooks.

Run `npm install` once. It installs the dependencies and turns on the pre-commit hook that blocks personal data (see `README.md`).

### Dependencies

| Package | Used for |
|---|---|
| `exceljs` | Reading and writing `.xlsx` workbooks. The only Excel library; don't add others. |
| `pdfkit` | The PDF copy of `--save` output. Optional at runtime: if it fails to load, no PDF is written. |
| `glob` | File matching in `scripts/batch_run.js`. |
| `eslint`, `@eslint/js`, `globals` (dev) | Linting (`npm run lint`). |

## How It Works

The Excel workbook is the database and the user interface. There is no server or other storage.

1. **Generate:** `generate_excel.js` creates a workbook with the Setup, transaction, Ledger, Summary and VERSION sheets.
2. **Import:** `load_transactions.js` maps a bank or card export's columns onto a transaction sheet (`lib/loader/`).
3. **Report:** `report.js` runs these phases in order, each in `lib/report/`, sharing state through the context object in `lib/report/context.js`:
   1. `setup.js`, `sheet-config.js`, `external-vendors.js`: read categories, vendors, customers, payer info and sheet configuration.
   2. `transactions.js` (with `sheet-columns.js`, `columns.js`, `linkage.js`): process each transaction sheet, apply polarity and the tax-year filter, total by category/vendor/customer, record integrity issues, and apply the sheet's net change to its linked account.
   3. `ledger.js`: apply manual double-entry rows; stop if unbalanced or undated.
   4. `wallets.js`: compare each account's calculated ending balance with the End Balance in Setup.
   5. `build-reports.js`: build P&L, balance sheet, vendor and customer rows.
   6. `print-statements.js`, `print-1099.js`, `diagnostics.js`: print reports and the final status; set the exit code.
   7. `summary.js` and `lib/output/`: `--save` output.

Shared helpers: `lib/accounting.js` (amounts, dates, tax year, 1099 thresholds), `lib/report/cells.js` (reading cell values), `lib/cli.js` (options), `lib/logger.js` (console capture).

## Standards

- **Lint and test before committing:** `npm run lint` and `npm test`. CI runs both.
- **Parse amounts with `parseAmount`** from `lib/accounting.js`, never `parseFloat`. It handles `$`, commas, `(x)`, `x-` and `CR`, and returns `NaN` for text so bad data is reported rather than treated as 0.
- **Don't swallow errors.** Row errors must be reported and fail the run; ESLint's `no-empty` rule blocks empty `catch` blocks.
- **No leftover debug logging.** Use the `showDebug` flag for tracing.
- **Excel/PDF parity:** if you add a column or calculation to the Excel output, add it to the PDF too.
- **Paths:** handle UNC paths (`\\server\share`) and spaces in paths.
- **Setup parsing:** changes must keep working with both formal Excel tables and header-row detection.
- **Asset polarity:** assets behave inversely to expenses relative to bank polarity (spending increases a purchased asset). New features must respect this.
- **Tests write to a temp directory**, never into the repo. See `docs/testing.md`.
- **Keep the docs in sync:** a new flag or behavior change updates `docs/features.md` (and `docs/setup.md` if the Setup sheet changes).

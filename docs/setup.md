# Book File Reference (`--book`)
*Configuration source: one TOML file per book. Start from `examples/example.book.toml`.*

A **book** is one LLC's configuration in a text file, instead of in the workbook's Setup sheet. With `--book`, `report.js` and `load_transactions.js` read the company info, sheets, categories, vendors, customers and defaults from that file, and **do not read the Setup sheet at all**. The workbook then holds only transactions, the Ledger and a generated `Lists` sheet. Without `--book`, everything works as before from the Setup sheet (the [legacy path](#setup-sheet-reference-legacy-path) below; it will be removed once every book has moved).

## Files

```
books.toml                  registry (optional): id -> path of the book file
acme.book.toml              everything the Setup sheet held, except tax IDs
acme.private.toml           tax IDs only (company and vendors). Never commit it; keep it away from AI agents.
Acme_2026.xlsx              transaction sheets + Ledger + generated "Lists" sheet (no Setup)
acme-2025-closing.toml      opening balances (optional, see README "Opening Balances")
```

`--book` takes an id or a path:

- `--book acme.book.toml` (anything ending in `.toml`, or an existing file) is a path.
- `--book acme` is an id, looked up in the registry: the file named by the environment variable `LLC_BOOKS`, else `./books.toml`. The registry is flat, one line per book, paths relative to the registry file:
  ```toml
  acme = "acme/acme.book.toml"
  ```
  The id must equal `id` in the book file.
- Relative paths **inside** a book file (`workbook`, `opening_balances`, `private`) resolve against the book file's folder.

## First LLC: step by step

1. Copy `examples/example.book.toml` and `examples/example.private.toml` next to where you keep your books and rename them (`acme.book.toml`, `acme.private.toml`).
2. Edit the book: your sheets, categories, vendors and customers. Put tax IDs only in the private file.
3. Create the workbook: `node generate_excel.js --book acme.book.toml` (makes the file named in `workbook`, with one sheet per `[[sheets]]` entry, a Ledger and the `Lists` sheet; it refuses to overwrite an existing file). An existing workbook works too, if its sheet names match.
4. Import statements: `node load_transactions.js --book acme.book.toml "chase_jan.csv" "Chase Checking"`.
5. Report: `node report.js --book acme.book.toml` (add `--1099` for 1099 output; add `--save` to write the report workbook).

A problem in the book stops the run with exit code 1 before the workbook is opened. **Every problem is listed at once, each with a suggested fix** (the closest valid key, type, sheet or category name).

## `<book>.toml` schema

Unknown keys and tables are errors, with the closest valid name suggested.

### `[book]`

| Key | Required | Meaning |
|---|---|---|
| `id` | yes | Short name of the book (the registry key). |
| `workbook` | yes | The Excel file. |
| `name` | no | Display name. |
| `year` | no | Default for `--year` (a 4-digit year). `--year` on the command line overrides it. |
| `opening_balances` | no | Opening-balances file (see README). `--opening-balances` on the command line overrides it. |
| `private` | no | The private file. Only read for `--1099` / `--all`. |
| `default_flags` | no | Report flags used on every run, e.g. `["--pl", "--bs"]`. Merged with the flags you type. Allowed: `--all --pl --bs --pl-sub --bs-sub --vendor --vendor-sub --customer --customer-sub --1099 --1099=NEC --1099=INT --checker --save --debug --ignore-vendors`. |
| `blank_sheets` | no | Empty tabs for your own use, e.g. `["Worksheet"]`, added after the ledger when the workbook is generated. The report ignores them. A name equal to a book sheet, `Lists` or `VERSION` is an error. |

### `[company]` (the payer, printed on 1099 output)

`name`, `address`, `city`, `state`, `zip`, `country`, `email`, `phone`, all optional strings. The company tax ID is **not** here (it is in the private file); a `tin` key in the book file is an error.

### `[[sheets]]` (one per worksheet to process)

| Key | Required | Meaning |
|---|---|---|
| `name` | yes | The exact worksheet name. It must exist in the workbook. |
| `type` | yes | `bank`, `cc`, `income`, `expense` or `ledger`. |
| `link` | yes (not for `ledger`) | The Balance Sheet category (asset, liability or equity) the sheet reconciles to. |
| `end_balance` | yes (not for `ledger`) | The statement's ending balance for the year. `0` is valid; it is required, not defaulted. The calculated balance must match it (see `docs/features.md`, End Balance required). |
| `start_balance` | no | Beginning balance (the opening-balances file is the other way). |
| `flip` | no | `true` for statements that show charges as positive. Default: `true` for `cc` and `expense`, `false` for `bank`, `income` and `ledger`. |
| `header_row` | no | Row number of the column headers. Without it the header row is found automatically. Sheets filled by `load_transactions.js` use row 3. |
| `short_name` | no | Abbreviation for report columns. |

A `ledger` sheet must be named `Ledger` or `General Ledger` and takes no `link` or balances. If the book has no `ledger` entry, the Ledger sheet is found and read as before.

### `[[categories]]`

| Key | Required | Meaning |
|---|---|---|
| `name` | yes | The category name. |
| `type` | yes | `asset`, `liability`, `equity` (Balance Sheet), `income`, `expense` (P&L) or `transfer`. The report (P&L, Balance Sheet, Transfer) follows from the type; there is no separate column. |
| `subcategories` | no | List of sub-category names. |
| `transfer_account` | no | Only for `transfer`: the Balance Sheet category the transfer belongs to (a `transfer` row on another sheet gets a warning). |
| `description` | no | A note for you. Also allowed on `[[sheets]]`, `[[vendors]]` and `[[customers]]`. The tool never prints or uses it. |

### `[[vendors]]`

| Key | Required | Meaning |
|---|---|---|
| `name` | yes | The exact text used in the Vendor column. |
| `form_1099` | no | `NEC`, `INT`, `MISC` or `none` (default `none`). |
| `business_name`, `recipient_name` | no | Legal business name; the recipient's name if different from `name`. |
| `address`, `city`, `state`, `zip`, `country`, `email`, `phone` | no | Printed on 1099 output. `address`, `city`, `state`, `zip` and the tax ID are required for a vendor over the 1099 threshold. |

### `[[customers]]`

`name` (required): the exact text used in the Customer column.

### Checks beyond keys and types

- Duplicate names (case-insensitive; vendors and customers also ignore extra spaces) in `[[sheets]]`, `[[categories]]`, a category's `subcategories`, `[[vendors]]`, `[[customers]]`.
- A vendor or customer with the same name as a category (the Setup sheet only warned; a book refuses).
- `link` must name a Balance Sheet category; linking to a P&L or Transfer category, or to a name that is not a category, is an error.
- `transfer_account` must name a Balance Sheet category.
- TOML errors, such as a duplicate key, are reported with their line and column.
- Sheets not in the workbook are listed with the closest worksheet name.
- A vendor and a customer with the same name, and sub-category names that equal a vendor or customer, still print the same warnings as with the Setup sheet.

## `<book>.private.toml` (tax IDs)

```toml
[company]
tin = "00-0000000"

[vendors."Acme Plumbing"]     # must match a [[vendors]] name in the book
tin = "00-0000001"
```

- Only `--1099` / `--all` read it. A named private file that does not exist is an error then; without a 1099 report it is not read.
- **A tax ID is never printed to the console** (not with `--debug` either), never put in an error message, and not in the `--save` Processing Log. It appears only in the saved `1099 Data` sheet. The console shows `Tax ID:  on file` for the payer, `on file` in the recipient table, or `MISSING` / `Missing: Tax ID (SSN/EIN)` when absent.
- The report stamp lists the private file by name, hash and size only.

## Other behavior with `--book`

- **Setup sheet ignored.** If the workbook still has a Setup sheet, one line is printed: `Setup sheet ignored: config comes from acme.book.toml`.
- **Workbook.** The book names it; do not also give a file name on the command line (that is an error, so a run can never mix two books).
- **`vendor.xlsx` / `vendor.csv`** are not searched for (the vendors are in the book); `--vendor-file <path>` still works.
- **Year.** `year` in the book is the default `--year`; `--year` on the command line wins.
- **Flags.** `default_flags` plus your flags (a flag given twice is the same as once).
- **Report stamp.** The inputs list the workbook, the book file, the private file (when read), the opening-balances file and any vendor file. `Flags` shows the command line as typed.

## `load_transactions.js --book`

```bash
node load_transactions.js --book acme "chase_jan.csv" "Chase Checking" [--clear]
node load_transactions.js --book acme --refresh-lists
```

- The book names the workbook and the sheets. The sheet argument must be a `[[sheets]]` name (a near miss is suggested); `cc` sheets use the card layout, the others the bank layout.
- **The lock is checked first**, before the export is read: the workbook is opened for update, and if Excel has it open the run stops with exit code 1 and `close it in Excel`, having changed nothing.
- The rows are written, then the **`Lists` sheet is rewritten** from the book: columns `Categories`, `SubCategories`, `Vendors`, `Customers` (and `Report Type (for Categories)`, which the Report Type formula uses). It carries a note that it is generated (edit the book, not the sheet) and is protected without a password (Review > Unprotect Sheet).
- The Category, Sub-Category, Vendor and Customer dropdowns and the Report Type formula point at `Lists!$A$2:$A$2000` and so on, never at Setup. The range is 2000 rows long, so new entries show up after `--refresh-lists` without touching the transaction sheets.
- `--refresh-lists` rewrites only `Lists` (after you edit categories, vendors or customers in the book).

# Setup Sheet Reference (legacy path)
*Configuration source without `--book`: the `Setup` tab of your workbook (start from `node generate_excel.js`). It will be removed once every book uses a book file.*

The "Setup" sheet is the brain of the application. It processes **5 distinct tables**. These tables usually reside side-by-side or stacked. The system detects them by their **Header Names**.

## 1. Category Mapping Table
*Defines how transactions are categorized and reported.*
- **Crucial Header**: `Category`
- **Columns**:
    | Header Alias | Purpose |
    | :--- | :--- |
    | `Category` | The main bucket name (e.g., "Office", "Travel"). |
    | `SubCategory` | (Optional) Granular detail (e.g., "Software", "Flights"). |
    | `AccountType` / `Type` | **Strictly Enforced**: <br> • If Report=`P&L`: Must be `Income` or `Expense`. <br> • If Report=`Balance Sheet`: Must be `Asset`, `Liability`, or `Equity`. |
    | `Report` / `PnL/BS` | `P&L` or `Balance Sheet`. Determines which report it hits. |

## 2. Vendor Configuration Table
*Central database for payee details and 1099 compliance.*
- **Crucial Header**: `Vendors` (or `Vendor`)
- **Columns**:
    | Header Alias | Purpose |
    | :--- | :--- |
    | `Vendors` | The EXACT string found in your bank/CC/ledger description. |
    | `BusinessName` | Legal business name for 1099. |
    | `SSN` / `EIN` / `TaxID` | **REQUIRED** for 1099 generation. |
    | `Address` | **REQUIRED** for 1099 generation. |
    | `1099Type` | `NEC`, `INT`, or `MISC`. |
    | `1099Required` | `YES` or `NO`. |

## 3. Customer Configuration Table
*Database for payer/income sources.*
- **Crucial Header**: `Customers` (or `Customer`)
- **Columns**:
    | Header Alias | Purpose |
    | :--- | :--- |
    | `Customers` | Name of the client/customer. |

## 4. Sheet Information Table
*Tells the tool which tabs to read and how to treat them.*
- **Crucial Header**: `SheetName` (or `SheetNameConfig`)
- **Columns**:
    | Header Alias | Purpose |
    | :--- | :--- |
    | `SheetName` | Exact name of the Excel tab (e.g., "Chase_1234"). |
    | `Sheet Type` / `Type` | `Bank` (asset), `CC` (liability), `Income`, `Expense`. Use the header `Sheet Type` when the Category table also has a `Type` column, so the two are not confused. |
    | `Flip Polarity` / `Flip` | `Yes` for statements that show charges as positive (most credit cards). All reporting treats money in as positive and money out as negative. |
    | `Header Row` / `Offset` | Row number of the column headers. Sheets filled by `load_transactions.js` use row 3 (rows 1-2 hold TOTAL and SUBTOTAL). |
    | `LinkAsset` / `Category` | For Asset/Liab sheets: The Balance Sheet Account name it reconciles to. |
    | `StartBalance` | Beginning balance ($) for the period. |
    | `EndBalance` | **Required** for every `Bank` or `CC` sheet and every sheet linked to a Balance Sheet account: the ending balance on the statement for the tax year. The calculated balance (start + sheet + ledger) must match it. Empty is an error (`UNRECONCILED`, exit 1); `0` is a valid balance. Not needed for the Ledger sheet. The template from `generate_excel.js` has an `End Balance` column set to 0; enter the real figure after loading a statement. |
    | `ShortName` | Abbreviated name for columns in detailed reports. |

## 5. Company Info (Payer) Table
*Your LLC's details, printed on 1099 output. A two-column key/value list.*
- **Crucial Header**: `Company Info` (or `Payer Info`). The value goes in the column to its right.
- **Keys read**: `Company Name` (or `Payer Name` / `Business Name` / `Name`), `TIN` (or `EIN` / `Tax ID`), `Address`, `City`, `State`, `Zip`, `Country`, `Email`, `Phone`.

## Header Detection
If the Setup sheet has formal Excel tables named `CompanyInfo`, `Categories`, `Vendor`, `Customer` and `SheetInfo`, they are used. Otherwise the tool scans the first 20 rows for the row with the most known header names, and matches columns by name. When two tables share a header name (such as `Type` or `Category`), the Category table uses the first occurrence and the Sheet table uses the last.


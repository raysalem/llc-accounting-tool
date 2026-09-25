# LLC Accounting Tool

A node.js-based suite of tools for managing LLC accounting using Excel as the primary interface. This tool automates the process of importing transactions, reconciling polarities between bank and credit card statements, and generating Profit & Loss (P&L) and Balance Sheet (BS) reports.

## Features

- **Automated Excel Template Generation**: Create a standardized accounting structure with Setup, Ledger, Bank, and Credit Card tabs.
- **Polarity Reconciliation**: Standardize inflow and outflow signs across different account types (e.g., flipping CC purchases from positive to negative).
- **Transaction Importing**: Scripts to batch import CSV/Excel data into moving parts.
- **Financial Reporting**: Generate P&L and Balance Sheet reports with categorized data and automated ledger balancing.
- **Data Integrity Checks**: Real-time validation of categories, vendors, and customers to prevent classification errors.

## Core Concepts & Features

### 1. Polarity Reconciliation
Bank statements and Credit Card statements often use different polarities (e.g., CC purchases appearing as positive numbers on the statement). To unify these, the tool uses a "Flip Polarity" setting in the **Setup** tab:
- **Flip Polarity = Yes**: Multiplies all transaction amounts by -1. Use this for CC statements where expenses are positive.
- **Flip Polarity = No**: Leaves amounts as is. Use this for bank statements where income is positive and expenses are negative.
- **Goal**: All internal reporting treats **Asset Inflow as Positive** and **Asset Outflow as Negative**.

### 2. Template Structure
- **Setup Tab**: The control center. 
    - **Categories**: Define categories, assign Report Type (`P&L`, `Balance Sheet`, or `Transfer`), and optional `Transfer Account` for validation.
    - **Vendors**: Manage 1099 status.
    - **Sheet Config**: Map sheets to accounts and toggle `Flip Polarity`.
- **Ledger Tab**: For manual double-entry adjustments (e.g., depreciation, owner investments, or adjustments).
- **Dynamic Column Mapping**: Transactions sheets no longer require a fixed layout. The tool detects "Date", "Amount", "Category", etc., based on the headers in the row specified by the "Header Row" offset in the `Setup` tab.
- **Ledger Integration**: Manual entries in the `Ledger` tab with categories matching account types (e.g., "Bank" or "CC") or sheet names are automatically incorporated into the calculated balances on the Balance Sheet.
- **Transaction Tabs**: (e.g., Bank Transactions, Credit Card Transactions) Where imported or manual line items live.
- **Header Detection**: The tool automatically scans the first 5 rows to identify where your data headers ("Date", "Amount", "Category") begin, skipping summary rows like "Total" or "Balance" that might appear at the top of exports.

## Technical Notes

### Excel "Visual Tables" vs. Formal Tables
To prevent persistent "Problem with content" (XML corruption) errors often caused by modifying formal Excel Tables (`ListObjects`) programmatically:
*   This tool now uses **Visual Tables**: It applies standard Blue header styling and **AutoFilters** to the data range.
*   It does **not** create formal Excel Table objects.
*   This ensures files remain healthy and open without repair warnings, while still providing sorting/filtering functionality.

### Data Clearing Logic
*   The `--clear` flag is now a **"Nuclear Option"**.
*   It **completely deletes** the target worksheet and recreates it from scratch.
*   This guarantees zero persistence of old data, hidden rows, or stale metadata.

### Account Number Detection
The tool attempts to auto-detect the account number from:
1.  **Filename**: Looks for patterns like `...- 81002.xlsx`.
2.  **File Content**: Scans the top rows of the source file for "Account Number: XXXXX".
If found, it populates the "Account Number" column for all rows and displays it in the top header row (Cell J1).

### Layout Changes
*   **Top Totals**: Transaction sheets now feature `TOTAL` (Sum) and `SUBTOTAL` (Filtered Sum) rows at the very top (Rows 1 & 2) for immediate visibility.
*   **Header Row**: The main data header now starts at **Row 3**.

### 2a. How to Use the Ledger
The **Ledger** tab is for manual double-entry accounting. It directly impacts your calculated Balance Sheet totals.

**Polarity Rules (Bank Accounts / Assets):**
- **Debit (Dr)**: **INCREASES** the account balance. Use this for Opening Balances, Deposits, or Capital Injections.
- **Credit (Cr)**: **DECREASES** the account balance. Use this for manual Fees, Withdrawals, or Transfers out.
- **Formula**: `Bank Balance = Total Bank Transactions + (Ledger Debits to Bank - Ledger Credits to Bank)`

**Example - Opening Balance:**
If your bank account started the year with $1,000, enter a row in the Ledger:
- **Date**: 01/01/202x
- **Description**: Opening Balance
- **Category**: [Account Name] (e.g., "Bank Transactions")
- **Debit**: 1000
- **Credit**: 0

### 3. Data Integrity Checker
The tool validates every transaction during the report generation process:
- **Data Integrity Checker**: Groups issues (uncategorized, illegal categories, unknown vendors/customers) by tab. Enforces date requirements on all entries.
- **Offset Verification**: Automatically warns if the row immediately following your configured header offset looks like a header itself.
- **Master Lists**: Cross-references against categories, vendors, and customers defined in the Setup tab.
- **Illegal Values**: Highlights entries using undefined categories or unknown vendors.
- `--1099`: Generate 1099-NEC/INT reports for enabled vendors. This functionality uses standard "Positive Money Paid" polarity (e.g. $50,000 paid is 50000) and filters for Net Expenses only, ignoring refunds.
- `--vendor`: Generates a spending report for all vendors, showing total net payments (Expenses as Positive) and 1099 status.
- **Missing Data**: Flags transactions that are missing a category assignment.
- **Reporting**: Issues are summarized in red at the bottom of the **Summary** tab and printed to the console.
- **Deep Dive**: Use the `--checker` flag for specific row numbers and descriptions of every error.

## Getting Started

### Prerequisites

- Node.js (v14 or higher)
- Microsoft Excel

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/raysalem/llc-accounting-tool.git
   cd llc-accounting-tool
   ```
2. Install dependencies:
   ```bash
   npm install
   ```

### Basic Usage

1. **Initialize a new Excel sheet**:
   ```bash
   node generate_excel.js
   ```
2. **Configure your accounts**:
   Open the generated `LLC_Accounting_Template.xlsx`, go to the **Setup** tab, and define your sheet names, polarity preferences, and **Header Row** offset (useful if your bank export has extra rows at the top).
3. **Import Transactions**:
   Use `load_transactions.js` to bring in your bank or CC data.
4. **Update Financials**:
   Run the report script to refresh the Summary and generate reports:
   ```bash
   node report.js My_Books.xlsx --pl --bs --year=2025
   ```

### Tax Year
Reports cover one calendar year. Rows dated in other years (transactions and ledger entries) are skipped, and the report says how many were skipped.
- `--year=YYYY` picks the year. Without it, the tool uses `DEFAULT_TAX_YEAR` in `lib/accounting.js` (currently **2026**). Update that constant once a year.
- The year also sets the 1099-NEC/MISC threshold: **$600** for payments before 2026, **$2,000** from 2026.

### Exit Codes
`report.js` exits with code **1** when the books have critical errors (for example an unbalanced balance sheet or ledger, or incomplete 1099 vendor details) or data-integrity issues (uncategorized rows, unknown categories, vendors or customers). Otherwise it exits 0, so scripts and CI can detect problems.

### Amount Formats
Both `load_transactions.js` and `report.js` read amounts such as `$1,234.56`, `(1,234.56)` (negative), `1,234.56-` and `100.00 CR`. Text that is not a number is reported and skipped, never treated as 0.

## Running the Tests

```bash
npm test
```

This runs every file in `tests/` (unit tests for `lib/accounting.js`, the end-to-end integration test and the feature tests) and exits non-zero if any fail. See `TESTING.md` for the test plan, the coverage map and the expected numbers.

## Continuous Integration

This project uses **GitHub Actions** to ensure code quality. On every push or pull request to the `main` branch, `npm test` runs on Node 20 and 22. The build fails if any test fails.

## Key Scripts

- `generate_excel.js`: Creates the initial boilerplate Excel structure.
- `report.js`: The main engine for calculating balances and generating reports.
- `load_transactions.js`: Handles importing data from external sources.
- `lib/accounting.js`: Shared amount/date parsing, the default tax year and 1099 thresholds.
- `inspect.js`: Consolidated utility for debugging and data validation.

## License

MIT

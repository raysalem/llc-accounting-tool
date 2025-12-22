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
- **Setup Tab**: The control center. Define your categories, report mappings (P&L vs Balance Sheet), vendor lists, and sheet configurations (Sheet Name, Type, Flip Polarity, and Header Row offset).
- **Ledger Tab**: For manual double-entry adjustments (e.g., depreciation, owner investments, or adjustments).
- **Dynamic Column Mapping**: Transactions sheets no longer require a fixed layout. The tool detects "Date", "Amount", "Category", etc., based on the headers in the row specified by the "Header Row" offset in the `Setup` tab.
- **Ledger Integration**: Manual entries in the `Ledger` tab with categories matching account types (e.g., "Bank" or "CC") or sheet names are automatically incorporated into the calculated balances on the Balance Sheet.
- **Transaction Tabs**: (e.g., Bank Transactions, Credit Card Transactions) Where imported or manual line items live.

### 3. Data Integrity Checker
The tool validates every transaction during the report generation process:
- **Data Integrity Checker**: Groups issues (uncategorized, illegal categories, unknown vendors/customers) by tab. Enforces date requirements on all entries.
- **Offset Verification**: Automatically warns if the row immediately following your configured header offset looks like a header itself.
- **Master Lists**: Cross-references against categories, vendors, and customers defined in the Setup tab.
- **Illegal Values**: Highlights entries using undefined categories or unknown vendors.
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
   Run the update script to refresh the Summary and generate reports:
   ```bash
   node update_financials.js
   ```

## Running the Integration Test

You can verify the entire workflow (template generation -> transaction loading -> report generation) by running the integration test:

```bash
node tests/run_integration_test.js
```

This script:
1.  Generates a fresh template.
2.  Loads data from `example_bank.csv` and `example_cc.csv`.
3.  Simulates transaction categorization.
4.  Prints a financial report and verifies the totals.
5.  Saves a complete Excel artifact to `tests/Full_Accounting_Test_Case.xlsx`.

## Continuous Integration

This project uses **GitHub Actions** to ensure code quality. On every push or pull request to the `main` branch, the integration test suite is automatically executed.

## Key Scripts

- `generate_excel.js`: Creates the initial boilerplate Excel structure.
- `update_financials.js`: The main engine for calculating balances and generating reports.
- `load_transactions.js`: Handles importing data from external sources.
- `inspect.js`: Consolidated utility for debugging and data validation.

## License

MIT

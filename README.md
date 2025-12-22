# LLC Accounting Tool

A node.js-based suite of tools for managing LLC accounting using Excel as the primary interface. This tool automates the process of importing transactions, reconciling polarities between bank and credit card statements, and generating Profit & Loss (P&L) and Balance Sheet (BS) reports.

## Features

- **Automated Excel Template Generation**: Create a standardized accounting structure with Setup, Ledger, Bank, and Credit Card tabs.
- **Polarity Reconciliation**: Easily toggle polarity flipping for credit card statements where purchases appear as positive numbers.
- **Transaction Importing**: Scripts to batch import CSV/Excel data into the accounting template.
- **Financial Reporting**: Generate P&L and Balance Sheet reports directly from your transaction data.
- **Data Integrity Checks**: Verify categorized transactions and ensure no "illegal operations" (e.g., transfers between unlinked accounts) occur.

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

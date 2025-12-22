# LLC Accounting Template Project Context

## Overview
This project generates a standardized Excel template for LLC accounting. It handles transaction tracking, reconciliation, and automated reporting.

## Artifacts
- **Generator Script**: `generate_excel.js`
- **Report Updater**: `update_financials.js`
- **Importer Script**: `load_transactions.js`
- **Example Generator**: `generate_example_data.js`

## Reconciling Different Polarities
Bank statements and Credit Card statements often use different polarities (e.g., CC purchases appearing as positive numbers).

To reconcile these:
1.  Go to the **Setup** tab.
2.  In the **Sheet Configuration** section (Columns I, J, K):
    - Set **Flip Polarity?** to **Yes** for your Credit Card sheets if purchases are imported as positive numbers.
    - Set it to **No** for your Bank sheets if income is positive and expenses are negative.
3.  The `update_financials.js` script will automatically normalize all data to: **Asset Inflow = Positive** and **Asset Outflow = Negative** for your P&L and Balance Sheet reports.

## Template Structure

### 1. Setup Tab
- **Sheet Configuration** (Cols I/J/K):
    - `Sheet Name`: Name of the tab.
    - `Account Type`: `Bank` or `CC` (determines column mapping).
    - `Flip Polarity?`: Set to `Yes` to multiply all values by -1 (used to unify CC charges).

### 2. Usage Guide
1.  **Configure**: Define your sheets and polarity in Setup.
2.  **Transactions**: Enter data in the Bank or CC tabs.
3.  **Run Reports**: Close Excel and run `node update_financials.js`.
    - Use `--print-only` to see balances in the console.
    - Use `--pl` or `--bs` to filter outputs.

# Integration Test Overview

This directory contains the integration test suite for the LLC Accounting Tool. The test verifies the entire pipeline: template creation, data loading, categorization, and financial reporting.

## Test Components

- `example_bank.csv`: Sample bank transactions including income (Salary/Consulting) and expenses (Rent).
- `example_cc.csv`: Sample credit card purchases (Coffee/Supplies/Cloud Services).
- `run_integration_test.js`: The main test runner that automates the workflow.

## Test Scenario Description

This integration test simulates a typical month of LLC activity:
1.  **Revenue & Expenses**: Imports a bank CSV with a salary deposit, consulting income, and a rent payment.
2.  **Credit Card Spending**: Imports a CC CSV with office supplies and travel expenses.
3.  **Manual Adjustments**: Adds an owner investment and an audit adjustment via the Ledger.
4.  **Integrity Stress Test**: Intentionally adds one "illegal" category and one unknown vendor to verify that the checker correctly identifies and reports them.

## Expected Outcome

When running the integration test, the following values are calculated and verified:

### 1. Bank Balance (Calculated)
- **Input**: `6000.00` (CSV) + `1000.00` (Ledger Debit) + `150.00` (Integrity Rows)
- **Expected Total**: `7150.00`

### 2. Credit Card Balance (Calculated)
- **Input**: `180.50` (CSV charges)
- **Polarity Flip**: Applied (Expenses shown as negative impacts)
- **Expected Total**: `-180.50`

### 3. Ledger Impact
- **Owner Investment**: `1000.00` Debit to `Bank Transactions` (Increases Asset)
- **Audit Adjustment**: `50.00` Debit to `Office` (Increases Expense)

### 4. Net Income (P&L)
- **Revenue**: `7500.00` (Sales)
- **Expenses**: `-165.00` (CSV Office) - `15.50` (CSV Travel) - `1500.00` (CSV Rent) - `50.00` (Ledger Office)
- **Expected Net Income**: `5769.50`

## Test Artifacts
The test generates `tests/Full_Accounting_Test_Case.xlsx` which contains the complete setup, categorized transactions, and manual ledger entries.

## How to Run
From the project root:
```bash
node tests/run_integration_test.js
```

### Detailed Integrity Reporting
When running the financial update, you can use the `--checker` flag to see exact row numbers and details for any data integrity issues:
```bash
node update_financials.js --checker
```

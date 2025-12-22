# Integration Test Overview

This directory contains the integration test suite for the LLC Accounting Tool. The test verifies the entire pipeline: template creation, data loading, categorization, and financial reporting.

## Test Components

- `example_bank.csv`: Sample bank transactions including income (Salary/Consulting) and expenses (Rent).
- `example_cc.csv`: Sample credit card purchases (Coffee/Supplies/Cloud Services).
- `run_integration_test.js`: The main test runner that automates the workflow.

## Expected Outcome

When running the integration test, the following values are calculated and verified:

### 1. Bank Balance (Calculated from Transactions)
- **Input**: `5000.00` (Salary) - `1500.00` (Rent) + `2500.00` (Consulting)
- **Expected Total**: `6000.00`

### 2. Credit Card Balance (Calculated from Transactions)
- **Input**: `15.50` (Coffee) + `120.00` (Supplies) + `45.00` (Cloud)
- **Polarity Flip**: Applied (converts positive charges to negative impacts)
- **Expected Total**: `-180.50`

### 3. Ledger Impact
- **Owner Investment**: `1000.00` Debit to `Checking Account` (Increases Asset)
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

# Integration Test Overview

This directory contains the integration test suite for the LLC Accounting Tool. The test verifies the entire pipeline: template creation, data loading, categorization, and financial reporting.

## Test Components

- `example_bank.csv`: Sample bank transactions including income (Salary/Consulting) and expenses (Rent).
- `example_cc.csv`: Sample credit card purchases (Coffee/Supplies/Cloud Services).
- `run_integration_test.js`: The main test runner that automates the workflow.

## Expected Outcome

When running the integration test, the following values are calculated and verified:

### 1. Bank Balance (Calculated)
- **Input**: `5000.00` (Salary) - `1500.00` (Rent) + `2500.00` (Consulting)
- **Expected Total**: `6000.00`

### 2. Credit Card Balance (Calculated)
- **Input**: `15.50` (Coffee) + `120.00` (Supplies) + `45.00` (Cloud)
- **Polarity Flip**: Applied (converts positive charges to negative impacts)
- **Expected Total**: `-180.50`

### 3. Net Income (P&L)
- **Revenue**: `7500.00` (Salary + Client Payment)
- **Expenses**: `-1680.50` (Rent + Office Supplies + Travel)
- **Expected Net Income**: `5819.50`

## How to Run
From the project root:
```bash
node tests/run_integration_test.js
```

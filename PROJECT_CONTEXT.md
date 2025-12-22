# AI Project Context: LLC Accounting Tool

This document is a technical reference for the Antigravity agent. For user-facing features, installation, and usage, see `README.md`.

## System Architecture

The tool is a collection of Node.js scripts that interface with Excel (`.xlsx`) via the `exceljs` library. It avoids a complex database, treating the Excel workbook as the "Single Source of Truth" and the primary UI.

### Logic Flow
1.  **Generation**: `generate_excel.js` creates a workbook with strictly defined named ranges and column mappings.
2.  **Ingestion**: `load_transactions.js` performs fuzzy header matching to map external CSV/Excel data into the internal template columns.
3.  **Processing**: `update_financials.js` aggregates data.
    -   **Dynamic Mapping**: Identifies column indices by reading the header row at `config.offset`.
    -   **Normalization**: Flips polarities based on "Flip Polarity?" config in the Setup sheet.
    -   **Ledger Integration**: If a Ledger category name matches a Sheet Name or its `Account Type` (Setup Col 2), the impact is added to that account's calculated balance.
    -   **Validation**: Performs a multi-pass scan for uncategorized rows, rows missing dates, or "illegal" entries.
    -   **Reporting**: Writes calculated totals to the Summary sheet and console.

## Internal Data Mappings

### Sheet Configuration (Master Config)
Read from `Setup` sheet, Columns I-L:
- `I`: Sheet Name
- `J`: Account Type (`Bank` | `CC`)
- `K`: Flip Polarity (`Yes`/`No`)
- `L`: Header Row Offset (Transactions start at `offset + 1`)

### Transaction Column Indices
Initially hardcoded, but now resolved dynamically by `update_financials.js`. It looks for case-insensitive matches for "Date", "Description", "Amount", "Category", "Sub-Category", "Vendor", and "Customer" in the specified header row.

## Integrity Logic
The checker maintains Sets of `validCategories`, `validVendors`, and `validCustomers` from the Setup sheet. Any transaction using a value not in these sets is pushed into an `illegalCategories/Vendors/Customers` array along with its row and sheet context for the `--checker` report.

## CI/CD
- **GitHub Actions**: Configured in `.github/workflows/test.yml` to run `node tests/run_integration_test.js` on every push.

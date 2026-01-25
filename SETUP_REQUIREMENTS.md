# Excel Setup Requirements (Golden Logic)
*Configuration source: `LLC_Accounting_Template.xlsx` > `Setup` Tab*

The "Setup" sheet is the brain of the application. It processes **4 Distinct Tables**. These tables usually reside side-by-side or stacked. The system detects them by their **Header Names**.

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
    | `Type` | `Asset` (Bank), `Liability` (CC), `Income`, `Expense`. |
    | `LinkAsset` / `Category` | For Asset/Liab sheets: The Balance Sheet Account name it reconciles to. |
    | `StartBalance` | Beginning balance ($) for the period. |
    | `EndBalance` | Ending balance ($) for validation. |
    | `ShortName` | Abbreviated name for columns in detailed reports. |

---
*Generated: 2026-01-24*

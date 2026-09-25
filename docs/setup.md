# Setup Sheet Reference
*Configuration source: the `Setup` tab of your workbook (start from `node generate_excel.js`).*

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
    | `EndBalance` | Ending balance ($) for validation. |
    | `ShortName` | Abbreviated name for columns in detailed reports. |

## 5. Company Info (Payer) Table
*Your LLC's details, printed on 1099 output. A two-column key/value list.*
- **Crucial Header**: `Company Info` (or `Payer Info`). The value goes in the column to its right.
- **Keys read**: `Company Name` (or `Payer Name` / `Business Name` / `Name`), `TIN` (or `EIN` / `Tax ID`), `Address`, `City`, `State`, `Zip`, `Country`, `Email`, `Phone`.

## Header Detection
If the Setup sheet has formal Excel tables named `CompanyInfo`, `Categories`, `Vendor`, `Customer` and `SheetInfo`, they are used. Otherwise the tool scans the first 20 rows for the row with the most known header names, and matches columns by name. When two tables share a header name (such as `Type` or `Category`), the Category table uses the first occurrence and the Sheet table uses the last.


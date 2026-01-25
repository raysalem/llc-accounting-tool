# Future Roadmap & TODOs (2026 Season)

Derived from the v2.1.0 Post-Mortem and CPA Review.

### 1. Maintain Strict Reconciliation (Zero Tolerance)
*   **Context:** The shift to "Hard Stops" (`process.exit(1)`) on missing dates and the "End Balance Verification" loop proved critical for data integrity.
*   **TODO:** Ensure future versions **never** loosen these constraints. The system must refuse to generate a report if the Ledger is unbalanced or if Ledger + Sheet != Bank End Balance.

### 2. Implement 1099 Credit Card Exclusion
*   **Context:** Currently, 1099 totals are aggregated by Vendor Name across ALL payment methods. The IRS requires excluding payments made via Credit Card (or Third-Party Processors like PayPal/Upwork), as the processor files a 1099-K.
*   **Current Risk:** Safe Harbor Over-Reporting (annoying for vendors, but legal).
*   **TODO:** In `report.js`, modify the 1099 aggregation logic to check the Source Sheet type.
    *   `If (Sheet.Type == 'CC')` -> Skip adding to 1099-NEC stats.
    *   `If (Sheet.Type == 'Bank')` -> Include in 1099-NEC stats.

### 3. Architectural Refactoring (The Monolith)
*   **Context:** `report.js` has grown to 3,000+ lines. It contains file reading, parsing, accounting logic, reconciliation, and reporting all in one file.
*   **Risk:** Modifying business logic (like the 1099 fix above) carries a high risk of breaking unrelated monitoring or parsing logic.
*   **TODO:** Break the application into modules:
    *   `src/parser.js`: Excel/CSV reading.
    *   `src/ledger.js`: Double-entry accounting engine.
    *   `src/reconcile.js`: Validation and balancing logic.
    *   `src/report.js`: Output generation.

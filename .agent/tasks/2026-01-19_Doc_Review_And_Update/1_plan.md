# Plan: Documentation & Impact Review

## User Requirement
"When doing changes, review documentation and code base, to review impact and update accordingly, dont make changes in isolation."

## Context
I have made significant changes to:
1.  **Vendor Polarity**: Flipped calculation (Positive = Spending).
2.  **1099 Logic**: Updates to positive threshold checking and object-based map storage.
3.  **Setup Logic**: Aggressive header normalization.

## Impact Analysis
- **Codebase**:
    - `report.js`: Logic updated.
    - `PROJECT_CONTEXT.md`: Defines logic flow. Needs update?
    - `README.md`: Describes flags and behavior. Needs update?
    - `LLC_Accounting_Template.xlsx` (Conceptual): Logic changed interpretation of values.
- **Documentation**:
    - Does `PROJECT_CONTEXT.md` describe "Account Polarity"? Yes.
    - Does it mention "Vendor Report Polarity"? Currently output conventions might be stale.
    - Does `README.md` explain the new 1099 filtering logic (Net Payments)?

## Tasks
1.  **Review `PROJECT_CONTEXT.md`**:
    - Check "Application Logic" -> "Polarity" section.
    - Update to reflect that Vendor Reports now show **Positive** values for Spending.
    - Update 1099 section to mention "Net Payments Only" and "Payer Info Normalization".
2.  **Review `README.md`**:
    - Ensure `--vendor` report description matches new reality (Positive values).
    - Ensure `--1099` section mentions "Generated CSV uses standard Positive Polarity".
3.  **Review `.cursorrules`**: I already added the 4-phase workflow. Add this new rule? "Review Docs & Impact".

## Plan
1.  Read `PROJECT_CONTEXT.md` and `README.md`.
2.  Update `PROJECT_CONTEXT.md` with new Vendor Polarity and 1099 details.
3.  Update `README.md` to reflect positive 1099 values and robust setup.
4.  Add "Doc Review Rule" to `.cursorrules`.

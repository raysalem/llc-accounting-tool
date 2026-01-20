# Findings: BS/PL Classification Issue

## Root Cause Identified
The diagnostic tool revealed the problem:

### Duplicate Category Entries
The user's Setup sheet has **DUPLICATE** entries for "39th":
```
39th                          Asset          Balance Sheet  ← First entry (wins)
39th                          Expense        P&L            ← Duplicate
39th                          Expense        P&L            ← Duplicate
39th                          Income         P&L            ← Duplicate
```

The code processes Setup sheet rows sequentially and stores categories in a Map:
```javascript
uniqueCategories.set(lower, {
    report,
    accountType: typeVal,
    ...
});
```

When there are duplicates, **the first entry wins** because subsequent `.set()` calls overwrite the previous value. Since the first "39th" entry is marked as "Balance Sheet", that's what the code uses.

## Similar Issues
The diagnostic also shows duplicate "Granada" entries (all P&L, so no visible issue):
```
granada                       Expense        P&L
granada                       Expense        P&L
Granada                       Expense        P&L  (case difference!)
Granada                       Expense        P&L
```

And duplicate "taxes" and "repair" entries.

## User's Setup Sheet Problem
The user needs to:
1. **Remove duplicate category entries** from Setup sheet
2. **Keep only ONE entry per category**
3. **Ensure that entry has the correct Report type**

For "39th":
- Delete the "Asset / Balance Sheet" entry
- Keep one of the "P&L" entries (or create a single correct one)

## Code Improvement Opportunity
We identified this in the code review as item #10 (LOW PRIORITY):
> **Duplicate Detection - MISSING**: No check for duplicate categories, vendors, or customers in Setup

We should add a warning when duplicates are detected.

## Immediate Solution
User must fix their Setup sheet by removing duplicate "39th" entries and keeping only the P&L version.

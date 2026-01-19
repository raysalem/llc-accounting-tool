# Plan: Debug 1099 Report Failure

## Problem
The user reports that `node report.js ... --vendor --1099` "did not work".
Possible causes:
1.  **Threshold**: Payments might be under $600 (threshold check).
2.  **Config**: The "1099" column in Setup might not be detected securely or values ('NEC', 'INT') aren't matching due to whitespace/formatting.
3.  **Data Flow**: The accumulation logic for `vendor1099Stats` might be flawed or reset.
4.  **CLI Parsing**: Maybe multiple flags (`--vendor --1099`) are conflicting in `specificFilter` logic (though code allows overlap).
5.  **Output**: It prints "(No 1099 ...)" which might look like "not working" if the user expects data they *know* is there.

## Approach
1.  **Enhance Debugging**:
    - Add explicit logging in `report.js` when a vendor is marked as 1099-eligible.
    - Log *why* a vendor is skipped in the report (e.g., "Total $450 < $600").
    - Ensure the '1099' column search in Setup is robust (it was fuzzy/fallback, need to verify).
2.  **Rule Update**:
    - Add the requested "Reflection Rules" update to `.cursorrules`.

## Steps
1.  **Add Rule**: Update `.cursorrules` with the specific reflection questions.
2.  **Modify `report.js`**:
    - Add debug log: "Identified 1099 Vendor: [Name] (Type: [Type])".
    - Update `print1099` to show:
        - If list is empty: "No enabled vendors found."
        - If all filtered out: "Found N vendors, but none met $600 threshold."
3.  **Verify**: Run `run_integration_test.js` again to see the new debug output.

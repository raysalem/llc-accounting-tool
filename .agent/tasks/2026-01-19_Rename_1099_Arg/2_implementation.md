# Implementation: Rename 1099 Argument

## Details
1.  **Code Changes**:
    - Modified `report.js` to accept `--1099` as a primary flag.
    - Updated the `--help` output to show `--1099`.
    - Kept `--1099-nec` as a hidden supported alias for backward compatibility (though the help text highlights the new flag).
2.  **Documentation**:
    - Updated `README.md` to list `--1099` as the flag for generating "1099-NEC/INT reports".

## Robustness
- **Backward Compatibility**: Existing scripts using `--1099-nec` will still work because the check is `args.includes('--1099') || args.includes('--1099-nec')`.
- **Input Sanitization**: The flag check is boolean and resilient.
- **Strict Typing**: No new internal types introduced, but existing boolean flags are maintained.

## Considerations
- **User Intent**: The user explicitly stated "arg shoudl be 1099". This implies the report scope has broadened (NEC + INT), so the flag name must be generic.
- **Reporting Output**: The output logic (printing "1099-NEC REPORT" vs "1099-INT REPORT") remains separate sections based on the vendor's assignment in the Setup tab.

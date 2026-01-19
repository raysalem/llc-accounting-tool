# Plan: Rename 1099 Argument

## Problem
The current command-line argument `--1099-nec` is too specific, as the reporting functionality has been expanded to support both 1099-NEC and 1099-INT forms. The user requested generating a generic `--1099` argument that covers both cases.

## Approach
1.  **Refactor `report.js`**:
    - Update argument parsing to primarily look for `--1099`.
    - Deprecate or keep `--1099-nec` as a hidden alias (or remove it if strictly requested, but alias is safer for backward comp, though strictly user asked for "arg should be 1099"). I will prioritize the new name but support the old one as an alias for now, or just strictly switch if the user implies replacement. Given "arg shoudl be 1099 ... since report is for both", I will make `--1099` the primary canonical flag.
    - Update the `--help` menu text.
2.  **Update Documentation**:
    - Update `README.md` to reference `--1099` instead of `--1099-nec`.
3.  **Update Tests**:
    - Verify if any tests (like `test_arguments_coverage.js`) reference `--1099-nec` and update them to `--1099`.

## Steps
1.  Modify `report.js`:
    - Change flag detection logic.
    - Update Help string.
2.  Modify `README.md`:
    - Replace `--1099-nec` references.
3.  Modify `tests/test_arguments_coverage.js`:
    - Update test cases to use `--1099`.
4.  Verify by running `node tests/run_integration_test.js`.

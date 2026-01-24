# LLC Accounting Tool - Feature Set & Test Matrix

This document defines the official feature set of the functionality. Every feature listed here must be supported and tested.

## 1. Core Processing Features
- **File Input Support**: Accepts `.xlsx`, `.lnk`, and `.url` file paths.
- **Dynamic Ledger Balancing**: Calculates Asset/Liability totals based on double-entry principles.
- **Wallet Reconciliation**: Verifies that (Start + SheetChange - LedgerImpact) == Ending Balance.
- **Strict Configuration**: Requires "Setup" sheet with headers on Row 1 (Category, SubCategory, Report, Vendor, SheetName).

## 2. CLI Flags & Reporting
| Flag | Description | Test Coverage |
| :--- | :--- | :--- |
| `--help` | Displays usage instructions and exits. | ✅ `manual` |
| `--save` | Writes changes (Summary tab, formatting) back to the Excel file. | ✅ `test_comprehensive_report.js` |
| `--checker` | Runs data integrity checks (missing categories, illegal vendors). | ✅ `test_comprehensive_report.js` |
| `--debug` | Prints verbose log output for troubleshooting. | ✅ `manual` |
| `--pl` | Prints summary Profit & Loss report. | ✅ `test_comprehensive_report.js` |
| `--bs` | Prints summary Balance Sheet report. | ✅ `test_comprehensive_report.js` |
| `--pl-sub` | Prints detailed P&L with sub-categories and per-sheet breakdown. | ✅ `test_pl_sub_display.js` |
| `--bs-sub` | Prints detailed Balance Sheet with sub-categories and per-sheet breakdown. | ⚠️ **Needs Test** (Regression Check) |
| `--vendor` | Prints vendor spending summary. | ✅ `test_comprehensive_report.js` |
| `--vendor-sub` | Prints detailed vendor spending with breakdown. | ⚠️ **Needs Test** |
| `--customer` | Prints customer income summary. | ✅ `test_comprehensive_report.js` |
| `--customer-sub` | Prints detailed customer income with breakdown. | ⚠️ **Needs Test** |
| `--1099` | Generates 1099-NEC and 1099-INT reports based on Vendor configuration. | ✅ `test_1099_threshold.js` |
| `--1099=NEC` | Generates only 1099-NEC reports. | ✅ `test_1099_threshold.js` |
| `--1099=INT` | Generates only 1099-INT reports. | ✅ `test_1099_threshold.js` |
| `--details "Cat"`| Prints transaction-level details for a specific category. | ✅ `test_details_flag.js` |
| `--ignore-vendors`| Skips loading external vendor database files. | ⚠️ **Needs Test** |
| `--vendor-file` | Loads a custom vendor file path. | ⚠️ **Needs Test** |

## 3. Logic & Rules (Anti-Slop)
- **Signed vs Magnitude Reporting**:
  - In detailed reports (`*-sub`), "Additions" and "Subtractions" columns must sum the **Magnitude** (`Math.abs`) of values.
  - *Reason*: Preventing net cancellation when summing mixed-sign buckets (e.g. Ledger Debit [Neg] vs Payment [Pos]).
- **1099 Thresholds**:
  - Default NEC threshold is $600 (defined as `THRESHOLD_1099_NEC`).
  - INT threshold is $0.
- **Strict Headers**:
  - `Setup` sheet must be readable at Row 1.
  - Missing headers trigger a CRITICAL ERROR.

## 4. Test Gaps Checklist
- [ ] Create test for `--bs-sub` output format (verifying Positive Magnitude logic).
- [ ] Create test for `--vendor-sub` and `--customer-sub`.
- [ ] Create test for `--vendor-file` argument parsing.

---
*Updated: 2026-01-23*

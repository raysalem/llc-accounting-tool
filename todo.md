# TODO

Notes from an outside review of the repo (September 2026), updated after the first round of fixes.

## Done in this round

- [x] **Personal data removed from git history.** The bank statement CSV, 25+ `*.txt` debug/output files, the `.agent/tasks/` notes, `debug_headers.js`, `run_all_2025.bat` and ~12 one-off scripts with NAS paths were removed from every commit. Business names, the NAS address and the street address were replaced throughout. History was force-pushed.
- [x] **CI can fail now.** The integration test asserts real numbers and exits 1 on failure. `npm test` runs every test file. CI runs `npm ci && npm test` on Node 20 and 22 with `actions/*@v4`.
- [x] **All tests pass.** Causes that were fixed:
  - Template header mismatch: the Setup sheet-config column was named `Account Type`, which `report.js` read as the category type. It is now `Sheet Type`. The header row is now 3 to match `load_transactions.js`.
  - Integration ledger entries were one-sided; they are now balanced double entry.
  - The comprehensive test read the 1099 amount from the wrong column.
  - `--save` writes `report_<name>.xlsx`; the arguments test now checks that file.
  - Crash (`toUpperCase` of undefined) when a sheet had no short name.
  - `--pl-sub` hid the "(No Sub-Cat)" line when a category mixed rows with and without sub-categories, so the sub-lines did not add up to the total.
  - Removed `test_ledger_resilience.js`: it contradicted `test_ledger_strictness.js` and the rulebook (a ledger row with no date is a hard stop).
  - Folded `test_details_flag.js` into `test_arguments_coverage.js`; it only passed if another test's temp file happened to exist.
- [x] **One amount parser** (`lib/accounting.js` → `parseAmount`) used by both scripts. It handles `$`, commas, `(1,234.56)`, `1,234.56-` and `CR`. Previously `(1,234.56)` became +1234.56 in `report.js` and `"1,234.56"` became 1 in `load_transactions.js`.
- [x] **No more automatic version bump.** `report.js` no longer rewrites `package.json` on every `--save`.
- [x] **Tax year.** `DEFAULT_TAX_YEAR = 2026` in `lib/accounting.js`, which `--year=YYYY` overrides. Out-of-year transaction and ledger rows are skipped and counted. The 1099-NEC/MISC threshold is $600 before 2026 and $2,000 from 2026. The hardcoded `600`s are gone.
- [x] **Exit codes.** `report.js` exits 1 on critical errors or data-integrity issues. Before, it exited 0 unless `--save` was used.
- [x] **Time-zone-safe dates** in `load_transactions.js` (`parseDateUTC`), so `12/31/2025` can't slip into the next or previous year.

## Follow-up on GitHub (only you can do these)

- [ ] Rewriting history does not clear GitHub's caches. If the repo was ever public, assume the old data was seen. To purge cached views of old commits, contact GitHub Support and name the removed files, or delete and re-create the repo.
- [ ] Check for forks or other clones. They still contain the old history.
- [ ] Any local clone made before the rewrite must be re-cloned, not pulled. Pulling would merge the old history back in.
- [ ] Consider making the repo private. It is built around your real books.

## Test plan review

What the suite covers well now: the end-to-end numbers (P&L, balance sheet, A = L + E), the integrity checker, the 1099 threshold and missing vendor details, the tax-year filter, the CLI flags and the parsing helpers.

Gaps, in priority order:

- [x] **Transfers and credit-card payments.** `tests/test_transfers.js` automates the three `TESTING.md` scenarios with exact balances. The logic was already correct: the payment is counted once, refunds reduce the expense and the liability, and a transfer category on the wrong sheet raises a warning.
- [x] **Year-boundary test.** `tests/test_loader_formats.js` loads Dec 31 / Jan 1 rows in ISO and US formats under UTC, Los Angeles and Tokyo time zones. Against the old loader it fails: a Jan 1, 2026 charge landed in 2025 on a machine set to Tokyo time.
- [x] **End-to-end amount formats.** The same test loads `$1,234.56`, `(50.00)`, `1,000.00-` and `25`. The old loader turned `(50.00)` into 0 and `1,000.00-` into +1,000.
- [x] **Ledger stops.** `test_ledger_strictness.js` covers a missing date; `test_ledger_unbalanced.js` covers an unbalanced ledger (and that a balanced one passes).
- [x] **1099 boundaries.** `tests/test_1099_boundaries.js` covers $600.00 and $599.99 in 2025, a $700 payment with a $150 refund, $600 in 2026, and $2,000.00 and $1,999.99 in 2026.
- [x] **Tests write to `os.tmpdir()`.** A full `npm test` leaves no files in the repo.
- [x] **`tests/Full_Accounting_Test_Case.xlsx`** is only refreshed with `SAVE_ARTIFACT=1`.
- [x] **One test plan.** `TESTING.md` now has a coverage map (scenario → test file) and the expected numbers. `tests/EXPECTED_OUTCOME.md` was merged into it.
- [ ] **Most older feature tests build sample books that don't balance**, so they have to ignore `report.js`'s exit code: `test_1099_threshold`, `test_details_extended`, `test_pl_sub_display` and `test_arguments_coverage`. Give them balanced fixtures so they can assert exit 0, which would catch new errors.
- [ ] **Many older checks look for a word in the console output** (e.g. `includes('TestVendor')`). Where possible, assert the number on the same line (see `valueFor()` in `test_transfers.js`), or read the saved `report_*.xlsx`.

## Code

- [x] **Split `updateFinancials()`** into modules under `lib/report/` and `lib/output/`, and `load_transactions.js` into `lib/loader/`. Output was verified byte-for-byte identical on 102 report runs and 7 loader scenarios.
- [ ] **Unit-test the calculation phases directly** (e.g. `buildReports`, `applySheetLinkage`) now that they are separate functions.
- [ ] **Row processing is still one large closure** in `lib/report/transactions.js` (~370 lines) and `lib/report/ledger.js`. Splitting them further means rewriting them, not just moving code; do it with tests in place.
- [ ] **Shared state:** phases share a mutable `ctx` object. Over time, have each phase return its results instead of writing into `ctx`.
- [ ] **Dead code found during the split:** the 1099 section builds CSV text that is never written (`lib/report/print-1099.js`); `--save` fills a Summary sheet in the input workbook but never saves it (`lib/report/summary.js`); the end-balance check at the end of `lib/report/linkage.js` computes a difference it never uses. Either finish or remove each one. The "(Run with --save to update the Excel file)" hint is misleading: `--save` writes a separate report file.
- [ ] **Remove dead code and stale comments.** Examples: "REDUNDANT BLOCK REMOVED", "The previous logic here (Lines 713-732)", "TODO: Replace this with table-based reading once tables are created".
- [ ] **Setup parsing picks columns by header name.** Duplicate names across tables (`Type`, `Category`) are resolved by "first" vs "last" occurrence, which is how the template bug happened. Prefer the formal Excel tables (`CompanyInfo`, `Categories`, `Vendor`, `Customer`, `SheetInfo`) that the code already looks for, and make the template create them.
- [ ] **The fallback sheet configs** (used when Setup has none) assume header row 1, but `load_transactions.js` writes the header in row 3.
- [ ] **The integrity checker ignores rows it can't read.** Rows with unparseable dates are skipped with a warning only under `--checker`. Consider always counting them as issues.
- [ ] **Decide whether a `CRITICAL WARNING` should fail the run.** For example, a transfer category used on the wrong sheet prints a `CRITICAL WARNING` but `report.js` still exits 0.
- [ ] **Add ESLint and a formatter** and run them in CI. `.cursorrules` bans magic numbers and implicit conversions, but nothing enforces it.
- [ ] **Silent `catch (e) { }` blocks** (e.g. `generate_excel.js:102`, several in `report.js`) hide real failures. Log them at least under `--debug`.
- [ ] **The CSV parser in `load_transactions.js`** is a regex. It doesn't unescape `""` inside quoted fields and breaks on newlines inside quotes. Use a small CSV library, or document the limits. (`.cursorrules` currently forbids `csv-parser`; revisit that rule.)
- [ ] **`glob` is only used by `scripts/batch_run.js`**, and `pdfkit` is required at runtime. Check both are still needed and list them in `DEPENDENCIES.md`.
- [ ] **1099-INT threshold** is `0` (reports all interest). The IRS threshold is generally $10. Decide which you want and document it.
- [ ] **Remember to update `DEFAULT_TAX_YEAR`** each January (or derive it from the current date minus one year during filing season).

## Repo hygiene

- [ ] **Too many overlapping top-level docs:** `PROJECT_CONTEXT`, `FEATURE_SET`, `DEV_CHECKLIST`, `FUTURE_TASKS`, `SETUP_REQUIREMENTS`, `DEPENDENCIES`, `TESTING` and `ACCOUNTING_RULEBOOK`. Merge them into `README.md`, `docs/setup.md`, `docs/accounting-rules.md` and this file.
- [x] **One-off scripts moved to `scripts/`.** `fix_garbage.js` was deleted; it patched text in the old `report.js`.
- [ ] **`scripts/monitor_booking.js`** checks a campsite-booking website and is unrelated to this tool. Consider moving it to its own repo.
- [ ] **Keep real data out of the repo going forward:** keep `.gitignore` covering `*.txt`, `*.xlsx` and `*.csv` outside `tests/` and `examples/`. Add a pre-commit hook that blocks bank-export file names and your NAS paths.
- [ ] **Use descriptive commit messages** (not "minor" or "large chang set") so the history of an accounting tool can be audited.
- [ ] **Version numbering:** now that the patch number no longer auto-increments, bump the version by hand on releases. Consider resetting to a meaningful version (e.g. 3.1.0).

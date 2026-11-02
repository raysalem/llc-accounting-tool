# FAQ and How-To

Short answers to common questions, with the steps to follow. The full rules are in
[accounting-rules.md](accounting-rules.md) and [setup.md](setup.md).

## How-to: set starting (opening) balances

Use this when a new year starts and some balance-sheet accounts already have values: a mortgage, a loan,
equipment, money in the bank, the owner's investment. Profit & Loss accounts never carry over; they
start at zero every year.

**1. Get last year's closing numbers.** Use the balance sheet of the year that closed: last year's report,
or the return (Form 1065 Schedule L, end-of-year column), or the lender's year-end statement for a loan.
The tool does not read tax returns. You copy the numbers in yourself.

**2. Make sure each account is a balance-sheet category with the right type** in the book file
(`[[categories]]`). The type decides where the number goes:

| What it is | `type` | Positive number means |
|---|---|---|
| Bank account, equipment, property | `asset` | the LLC owns it |
| Mortgage, loan, credit card | `liability` | the LLC owes it |
| Owner investment / contributions | `equity` | the owner put it in |

**Check the loan types carefully.** A loan typed `asset` is not caught as an error: the equation still
says `[OK]`, but the loan shows up under Assets and Retained Earnings is too high by twice the loan. See the FAQ below.

**3. Write the opening-balances file** next to the book file, named after the year that closed,
e.g. `acme-2025-closing.toml`:

```toml
year = 2025                 # the year that closed (one before the book's year)
as_of_date = 2025-12-31     # optional: 2025-12-31 or 2026-01-01

[accounts]                  # keys are category names from the book
"Checking" = 1000.00
"Mortgage" = 250000.00
"Owner Investment" = 5000
```

Numbers only (no quotes, no `$` or commas), at most 2 decimals, always positive for the account's normal balance.

**4. Point the book at it.** In the book file under `[book]`:

```toml
opening_balances = "acme-2025-closing.toml"   # relative to the book file's folder
```

(For a one-off run without editing the book: `node report.js --book acme --bs --opening-balances other.toml`.
The command line wins over the book.)

**5. Fix the bank and card End Balances.** A sheet linked to an account with an opening balance starts from
that amount, so its `end_balance` (the statement's year-end balance) must include it. If it doesn't, the checker
reports `END BALANCE MISMATCH`.

**6. Check it.** Run `node report.js --book acme --bs`, or open the LLC on the local site. Expected:
- `Opening balances: N accounts from acme-2025-closing.toml (closing year 2025)` at the top,
- every account in the **Opening** column under the right heading (loans under Liabilities),
- a line **Retained Earnings (Opening Balance)** that makes the two sides equal,
- `Equation Status: [OK] (A = L + E)`.

## FAQ

**What is "Retained Earnings (Opening Balance)"?**
The difference between what the LLC owned and what it owed plus the equity you listed. It keeps
Assets = Liabilities + Equity true without you typing every past year's profit. If it looks wrong, an account
is missing, has the wrong type, or has the wrong amount.

**The equation is OK, but the loan shows under Assets.**
The category is typed `asset`. Change it to `type = "liability"` in the book file. Retained Earnings drops by twice
the loan amount, which is the correct value.

**"Sheet X has StartBalance … but the file says …. Remove one of them."**
A bank or card sheet has `start_balance` in the book and the same account is in the opening-balances file. Keep one:
the opening-balances file is the better place, so set the sheet's `start_balance = 0`.

**"File year 2024 must be the year before --year=2026."**
`year` in the file is the year that closed. For the 2026 books it is `2025`.

**"Account … does not match any Setup category."**
The name is misspelled or the category doesn't exist yet. Names are not case-sensitive, and spaces, `_` and `-`
are interchangeable. Add the category to the book first, or fix the spelling.

**"… is a P&L category. Only Balance Sheet accounts carry over."**
Income and expense accounts start at zero each year; remove them from the file.

**"must be a number, got a string; remove the quotes."**
Write `250000.00`, not `"250,000.00"`.

**My loan has two parts (e.g. a mortgage and a solar loan).**
Give each its own liability category and its own line in the file. Their sum should match the total on last
year's balance sheet, within rounding.

**How do I record loan payments during the year?**
Split each payment: the principal part goes to the loan's category (it lowers what is owed); the interest part goes
to an interest expense category (it's a P&L expense). The lender's year-end statement (Form 1098) gives the interest total to check against.

**Can I change the opening balances later?**
Yes. Edit the file and save; the local site recomputes, and with Before/after on you can see what moved.
A report already exported keeps its stamp (file hashes) so you can tell which version it used.

**What is the Version ID?**
The version of the tool that made a report: `<version>-<commit>-<code hash>`, e.g. `3.0.142-557496f-8c7bbd`.
It is printed at the top of every run, in the Report Info sheet, in the footer of every Excel sheet and PDF page,
and at the bottom of the site page (the site's own and the one of the report shown). Two reports with the same
Version ID were made by exactly the same code. If the site shows a red "code changed" banner, stop `site.js` and start it again.

**Does the opening-balances file change my workbook?**
No. It is read on every run; nothing is written to Excel unless you run a report with `--save`.

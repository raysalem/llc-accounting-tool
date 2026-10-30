// Checks each account's calculated ending balance against the End Balance in Setup.
// Every bank/card sheet (or sheet linked to a Balance Sheet account) must have an
// End Balance: a missing one is an error, an explicit 0 is a valid balance.

const path = require('path');

const END_BALANCE_TOLERANCE = 0.05; // Largest difference still counted as a match

// True when the sheet is an account with a statement: not the ledger, and a Bank/CC sheet
// or one linked to a Balance Sheet account. The default sheets (used when Setup has no sheet
// configuration) have no End Balance column to fill in, so they are exempt.
function needsEndBalance(config) {
    if (config.type === 'ledger' || config.isDefault) return false;
    // Same type words sheet-config.js accepts for bank and card sheets
    const t = config.type || '';
    const isAccountType = t === 'bank' || t === 'cc' || ['asset', 'credit', 'liability'].some(w => t.includes(w));
    return isAccountType || Boolean(config.linkedAccount);
}

async function validateWallets(ctx) {
    const {
        catStats, processedSheetTotals, sheetConfigs, showChecker, uniqueCategories,
    } = ctx;

    // Now that Ledger entries are integrated into 'catStats', we can accurately check Wallet Balances.

    if (showChecker) console.log(`\n--- WALLET RECONCILIATION ---`);
    if (showChecker) console.log(`Sheet Name`.padEnd(30) + `Start`.padStart(10) + `Sheet`.padStart(10) + `  Ledger`.padStart(10) + `  End (Calc)`.padStart(12) + `  End (Exp)`.padStart(12) + `  Diff`.padStart(10));

    // Counted for the one-line summary printed by diagnostics.js
    const reconciliation = { total: 0, mismatches: [], missing: [] };
    ctx.reconciliation = reconciliation;

    sheetConfigs.forEach(config => {
        if (config.type === 'ledger') return;
        const required = needsEndBalance(config);
        if (required) reconciliation.total++;

        // Re-Calculate End Balance using same logic, but now catStats is full
        let calculatedEnd = 0;
        const sheetTotal = processedSheetTotals[config.name] || 0;

        let effectiveLink = config.linkedAccount;
        if (!effectiveLink && config.cat) {
            const jitMatch = uniqueCategories.get(config.cat.toLowerCase());
            if (jitMatch) effectiveLink = jitMatch.displayName;
        }

        let ledgerImpact = 0;

        // Determine correct "Sheet Change" display polarity
        const sheetChangeVal = (config.type.toLowerCase() === 'expense') ? -sheetTotal : sheetTotal;

        if (effectiveLink && catStats[effectiveLink]) {
            // Global Balance (Start + Sheet + Ledger)
            calculatedEnd = catStats[effectiveLink].total;

            ledgerImpact = calculatedEnd - (config.startBalance || 0) - sheetChangeVal;

        } else {
            // Unlinked: Just Start + Sheet
            calculatedEnd = (config.startBalance || 0) + sheetChangeVal;
        }

        // A value that is not a finite number (NaN from an unreadable cell) counts as missing;
        // otherwise NaN compares false against the tolerance and passes as a match.
        const exp = Number.isFinite(config.endBalance) ? config.endBalance : null;
        if ((exp === null || exp === undefined) && required) {
            reconciliation.missing.push(config.name);
            ctx.hasErrors = true;
            const where = ctx.book ? `end_balance in ${path.basename(ctx.book.file)}` : 'End Balance in Setup (SheetInfo)';
            console.error(`  [!] UNRECONCILED: "${config.name}" has no ${where}. Enter the statement's ending balance for the tax year.`);
        }
        if (exp !== null && exp !== undefined) {
            const diff = Math.abs(calculatedEnd - exp);

            if (showChecker) {
                console.log(
                    `${config.shortName.substring(0, 29).padEnd(30)}` +
                    `${(config.startBalance || 0).toFixed(2).padStart(10)}` +
                    `${sheetChangeVal.toFixed(2).padStart(10)}` +
                    `${ledgerImpact.toFixed(2).padStart(10)}` +
                    `${calculatedEnd.toFixed(2).padStart(12)}` +
                    `${exp.toFixed(2).padStart(12)}` +
                    `${diff > END_BALANCE_TOLERANCE ? ('!' + diff.toFixed(2)).padStart(10) : ''.padStart(10)}`
                );
            }

            if (diff > END_BALANCE_TOLERANCE) {
                ctx.hasErrors = true;
                if (required) reconciliation.mismatches.push(config.name);
                console.warn(`  [!] END BALANCE MISMATCH: "${config.name}" (Calc: ${calculatedEnd.toFixed(2)} vs Exp: ${exp.toFixed(2)})`);
            }
        }
    });
}

module.exports = { validateWallets };

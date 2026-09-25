// Checks each account's calculated ending balance against the End Balance in Setup.

async function validateWallets(ctx) {
    const {
        catStats, processedSheetTotals, sheetConfigs, showChecker, uniqueCategories,
    } = ctx;

    // --- 3b. Post-Ledger Wallet Validation ---
    // Now that Ledger entries are integrated into 'catStats', we can accurately check Wallet Balances.

    if (showChecker) console.log(`\n--- WALLET RECONCILIATION ---`);
    if (showChecker) console.log(`Sheet Name`.padEnd(30) + `Start`.padStart(10) + `Sheet`.padStart(10) + `  Ledger`.padStart(10) + `  End (Calc)`.padStart(12) + `  End (Exp)`.padStart(12) + `  Diff`.padStart(10));

    sheetConfigs.forEach(config => {
        if (config.type === 'ledger') return;

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

        const exp = config.endBalance;
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
                    `${diff > 0.05 ? ('!' + diff.toFixed(2)).padStart(10) : ''.padStart(10)}`
                );
            }

            if (diff > 0.05) {
                ctx.hasErrors = true;
                console.warn(`  [!] END BALANCE MISMATCH: "${config.name}" (Calc: ${calculatedEnd.toFixed(2)} vs Exp: ${exp.toFixed(2)})`);
            }
        }
    });
}

module.exports = { validateWallets };

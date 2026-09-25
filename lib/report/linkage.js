// Applies a processed sheet's net change to the balance-sheet account it is
// linked to (from Setup, or found from the sheet's category), excluding
// transfer rows so card payments are not counted twice.

function applySheetLinkage(ctx, config, { sheetTotal, sheetAdds, sheetSubs, excludedFromLinkage, excludedAdds, excludedSubs }) {
    const { catStats, uniqueCategories, showDebug } = ctx;

    // JIT Linkage Fallback: If no linked account from Step 1, try again using raw cat column
    // This catches cases where categories might have fully loaded or matched differently
    let effectiveLink = config.linkedAccount;
    if (!effectiveLink && config.cat && config.type !== 'ledger') {
        const jitMatch = uniqueCategories.get(config.cat.toLowerCase());
        if (jitMatch) effectiveLink = jitMatch.displayName;
    }

    if (effectiveLink) {
        const linkName = effectiveLink;
        const lConf = uniqueCategories.get(linkName.toLowerCase());
        const aType = (lConf && lConf.accountType) ? lConf.accountType.toString().toLowerCase() : '';
        const isAsset = aType.includes('asset') || aType.includes('bank') || aType.includes('cash');
        const isLiability = aType.includes('liability') || aType.includes('credit') || aType.includes('cc') || aType.includes('payable') || aType.includes('loan');

        if (!catStats[linkName]) catStats[linkName] = { total: 0, subCats: {}, sheets: {} };
        if (!catStats[linkName].sheets) catStats[linkName].sheets = {};

        const previous = catStats[linkName].total;
        // Polarity Logic: Assets increase with positive flow (Income). Liabilities/Equity decrease.
        // Balance Sheet Items:
        // Asset: Balance += sheetTotal
        // Liability/Equity: Balance -= sheetTotal
        const effectiveSheetTotal = sheetTotal - excludedFromLinkage;

        if (isAsset) {
            catStats[linkName].total += effectiveSheetTotal;
        } else {
            catStats[linkName].total -= effectiveSheetTotal;
        }

        let linkageMsg = `[Linkage Logic] Linked "${config.name}" to ${isAsset ? 'Asset' : 'Liability'} "${linkName}".`;
        if (config.startBalance && Math.abs(config.startBalance) > 0.001) {
            catStats[linkName].total += config.startBalance;
            linkageMsg = `[Linkage Logic] Applied Starting Balance (${config.startBalance.toFixed(2)}) and ${config.shortName} Total (${effectiveSheetTotal.toFixed(2)}) to ${isAsset ? 'Asset' : 'Liability'} "${linkName}".`;
        } else {
            linkageMsg = `[Linkage Logic] Applied ${config.shortName} Total (${effectiveSheetTotal.toFixed(2)}) to ${isAsset ? 'Asset' : 'Liability'} "${linkName}".`;
        }

        if (!catStats[linkName].sheets[config.name]) {
            catStats[linkName].sheets[config.name] = { add: 0, sub: 0, total: 0 };
        }
        const sStat = catStats[linkName].sheets[config.name];
        const sourceIsLiability = config.type.toLowerCase().includes('expense') || config.type.toLowerCase().includes('cc');

        const effSheetAdds = sheetAdds - excludedAdds;
        const effSheetSubs = sheetSubs - excludedSubs;

        if (isAsset) {
            // Asset Target: Direct Mapping (Inflow=Add, Outflow=Sub)
            sStat.total += effectiveSheetTotal;
            sStat.add += effSheetAdds;
            sStat.sub += effSheetSubs;
        } else {
            // Liability Target (Positive Balance = Debt)
            if (sourceIsLiability) {
                // Source is the Liability itself (CC Sheet)
                // Expense (Negative Outflow) -> Increases Debt (Add)
                // Payment (Positive Inflow) -> Decreases Debt (Sub)
                sStat.total -= effectiveSheetTotal; // Invert to make Exp positive
                sStat.add += (-effSheetSubs); // Map Negs to Add
                sStat.sub += effSheetAdds;    // Map Pos to Sub
            } else {
                // Source is External Asset (Bank Sheet) paying the Liability
                // Payment (Negative Outflow) -> Decreases Debt (Sub)
                // Refund (Positive Inflow) -> Increases Debt (Add)? (Rare, but logical)
                sStat.total += effectiveSheetTotal; // Direct (Neg reduces Debt)
                sStat.add += effSheetAdds;    // Pos adds to Debt
                sStat.sub += effSheetSubs;    // Neg subtracts from Debt
            }
        }

        console.log(`${linkageMsg} Balance: ${previous.toFixed(2)} -> ${catStats[linkName].total.toFixed(2)}`);
    } else if (config.type !== 'ledger') {
        // Only warn if it's not a Ledger (Ledgers are manual)
        if (showDebug) console.log(`[Linkage Logic] Sheet "${config.name}" (Type: ${config.type}) has NO LINKED ACCOUNT. Total (${sheetTotal.toFixed(2)}) NOT applied to any Balance Sheet asset.`);
    }

    // Check End Balance if configured (strict check against null)
    // Moved here to ensure 'effectiveLink' and 'catStats' are fully computed
    // if (config.endBalance !== null) {
    // Use the "Global" Ledger-adjusted balance for this account if linked
    let calculatedEnd = 0;
    let usedGlobal = false;

    if (effectiveLink && catStats[effectiveLink]) {
        // If linked, use the total accounting balance (includes Ledger + Sheet + Start)
        // Note: Assets are Positive, Liabilities are Negative in catStats.total logic?
        // Actually in catStats, Liability Balances are currently calculated as Positive numbers 
        // (Expenses=Increase, Payments=Decrease).
        calculatedEnd = catStats[effectiveLink].total;
        usedGlobal = true;
    } else {

        // Fallback to sheet-only calculation (Start + Change)
        // For Expense sheets, validationChange should be Positive (Increase) to match Ending Balance magnitude.
        const sheetChangeVal = (config.type.toLowerCase() === 'expense') ? -sheetTotal : sheetTotal;
        calculatedEnd = (config.startBalance || 0) + sheetChangeVal;
    }

    // Determine "Sheet Change" (The part coming from THIS sheet)
    // For Expense sheets, internal SheetTotal is Positive (Expenses). We invert it to match Balance Impact (Negative).
    // This is used for REPORTING display, and differentiating Ledger impact.
    const sheetChangeVal = (config.type.toLowerCase() === 'expense') ? -sheetTotal : sheetTotal;

    let finalCalcEnd = calculatedEnd; // Use the one determined above (Global or Fallback)
    let ledgerAdj = 0;

    if (usedGlobal) {
        // catStats.total includes: Start + SheetChange + Ledger
        // So Ledger = Total - Start - SheetChange
        ledgerAdj = finalCalcEnd - (config.startBalance || 0) - sheetChangeVal;
    }

    const diff = Math.abs(finalCalcEnd - config.endBalance);

    // Verification Log
}

module.exports = { applySheetLinkage };

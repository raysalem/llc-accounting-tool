// Prints the requested reports to the console: P&L, balance sheet, vendor,
// customer, 1099 preparation (lib/report/print-1099.js) and --details.
const { get1099Threshold } = require('../accounting');
const { NORM_VEND } = require('./cells');
const { printDetailedTable } = require('./format-table');
const { print1099Report } = require('./print-1099');

async function printStatements(ctx) {
    const {
        catStats, detailsRows, illegalVendors, netIncome, reports, sheetConfigs, showAll, showBSSub,
        showCustomer, showCustomerSub, showDetails, showPLSub, showVendor, showVendorSub,
        targetDetailsFilter, taxYear, vendor1099Map,
    } = ctx;

    // --- 5. Console Output ---
    // --- 5. Console Output ---

    // Collect all sheet names for columns (Order: Configured Sheets...)
    const distinctSheets = new Set();
    const sheetNameMap = {}; // Map full name -> short name
    sheetConfigs.forEach(s => {
        distinctSheets.add(s.name);
        sheetNameMap[s.name] = s.shortName;
    });
    // Ensure Ledger is included if it was processed
    if (Object.keys(catStats).some(c => catStats[c].sheets && catStats[c].sheets[ctx.ledgerConfig.name])) {
        distinctSheets.add(ctx.ledgerConfig.name);
        sheetNameMap[ctx.ledgerConfig.name] = ctx.ledgerConfig.shortName || ctx.ledgerConfig.name;
    }
    const reportSheetList = Array.from(distinctSheets);

    // PL report (Always display)
    if (true) {
        if (showPLSub) {
            printDetailedTable('PROFIT & LOSS (Detailed)', reports.pl, reportSheetList, sheetNameMap);
        } else {
            console.log(`\n--- PROFIT & LOSS ---`);
            if (!reports.pl.length) console.log('(No Data)');
            else {
                const max = Math.max(...reports.pl.map(r => r.label.length), 10);
                reports.pl.forEach(r => {
                    console.log(`${r.label.padEnd(max + 5)} : ${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
                });
            }
        }
        console.log(`\n=== NET INCOME: ${netIncome.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ===\n`);
    }
    // BS report (Always display)
    if (true) {
        console.log(`\n--- BALANCE SHEET ---`);

        if (showBSSub) {
            printDetailedTable('ASSETS (Detailed)', reports.assets, reportSheetList, sheetNameMap, "Category", true);
            console.log(`TOTAL ASSETS:`.padEnd(30) + reports.assets.reduce((a, b) => a + b.value, 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));

            printDetailedTable('LIABILITIES (Detailed)', reports.liabilities, reportSheetList, sheetNameMap, "Category", true);
            console.log(`TOTAL LIABILITIES:`.padEnd(30) + reports.liabilities.reduce((a, b) => a + b.value, 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));

            printDetailedTable('EQUITY (Detailed)', reports.equity, reportSheetList, sheetNameMap, "Category", true);
            console.log(`TOTAL EQUITY:`.padEnd(30) + reports.equity.reduce((a, b) => a + b.value, 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));
        } else {
            console.log(`\n[ASSETS]`);
            if (!reports.assets.length) console.log('(No Assets)');
            else {
                const max = Math.max(...reports.assets.map(r => r.label.length), 10);
                reports.assets.forEach(r => {
                    const open = (r.opening || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15);
                    const end = r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15);
                    console.log(`${r.label.padEnd(max + 5)} : ${open} -> ${end}`);
                });
            }
            const totalAssets = reports.assets.reduce((a, b) => a + b.value, 0);
            console.log(`TOTAL ASSETS:`.padEnd(30) + totalAssets.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));

            console.log(`\n[LIABILITIES]`);
            if (!reports.liabilities.length) console.log('(No Liabilities)');
            else {
                const max = Math.max(...reports.liabilities.map(r => r.label.length), 10);
                reports.liabilities.forEach(r => {
                    const open = (r.opening || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15);
                    const end = r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15);
                    console.log(`${r.label.padEnd(max + 5)} : ${open} -> ${end}`);
                });
            }
            const totalLiabilities = reports.liabilities.reduce((a, b) => a + b.value, 0);
            console.log(`TOTAL LIABILITIES:`.padEnd(30) + totalLiabilities.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));

            console.log(`\n[EQUITY]`);
            if (!reports.equity.length) console.log('(No Equity)');
            else {
                const max = Math.max(...reports.equity.map(r => r.label.length), 10);
                reports.equity.forEach(r => {
                    const open = (r.opening || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15);
                    const end = r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15);
                    console.log(`${r.label.padEnd(max + 5)} : ${open} -> ${end}`);
                });
            }
            const totalEquity = reports.equity.reduce((a, b) => a + b.value, 0);
            console.log(`TOTAL EQUITY:`.padEnd(30) + totalEquity.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));
        }
        const totalAssets = reports.assets.reduce((a, b) => a + b.value, 0);
        const totalLiabilities = reports.liabilities.reduce((a, b) => a + b.value, 0);
        const totalEquity = reports.equity.reduce((a, b) => a + b.value, 0);

        console.log(`\n` + `=`.repeat(55));
        console.log(`TOTAL LIABILITIES + EQUITY:`.padEnd(30) + (totalLiabilities + totalEquity).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(23));
        console.log(`=`.repeat(55));

        const eqDiff = Math.abs(totalAssets - (totalLiabilities + totalEquity));
        if (eqDiff < 0.01) {
            console.log(`Equation Status:             [OK] (A = L + E)`);
        } else {
            console.warn(`Equation Status:             [FAIL] Mismatch of ${eqDiff.toFixed(2)}`);
            console.warn(`[!] The Balance Sheet does not balance. Total Assets must equal total Liabilities + Equity.`);
            ctx.hasErrors = true;
        }
        console.log('');
    }

    if (showAll || showVendor) {
        // printSection('VENDOR SPENDING', reports.vendors); <-- Replacing with Detailed Table
        console.log(`\n--- VENDOR SPENDING ---`);
        if (reports.vendors.length === 0) console.log('(No Data)');
        else {
            const h = `Vendor`.padEnd(30) +
                `Total`.padStart(15) +
                `  1099 Type`.padEnd(12) +
                `Required`.padEnd(10);
            console.log(h);
            console.log('-'.repeat(h.length));

            reports.vendors.forEach(r => {
                const info = vendor1099Map.get(NORM_VEND(r.label)) || { type: '', req: '' };

                // Determine if vendor actually qualifies for 1099 reporting
                let displayReq = '';
                if (info.type) {
                    const threshold = get1099Threshold(info.type, taxYear);
                    const meetsThreshold = r.value > 0 && r.value >= threshold;
                    displayReq = meetsThreshold ? 'YES' : '';
                }

                console.log(
                    `${r.label.substring(0, 29).padEnd(30)}` +
                    `${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}` +
                    `  ${(info.type || '').padEnd(12)}` +
                    `${displayReq.padEnd(10)}`
                );
            });
        }

        // Explicitly warn about unknown vendors if requested
        if (illegalVendors.length > 0) {
            console.log(`\n[!] WARNING: ${illegalVendors.length} transactions have unknown Vendors.`);
            const uniqueUnknown = Array.from(new Set(illegalVendors.map(i => i.value)));
            console.log(`    Unknown Vendors: ${uniqueUnknown.join(', ')}`);
        }
    }
    print1099Report(ctx);

    // Customer report (already handles sub)
    if (showAll || showCustomer || showCustomerSub) {
        if (showCustomerSub) {
            printDetailedTable('CUSTOMER INCOME (Detailed)', reports.customers, reportSheetList, sheetNameMap, "Customer");
        } else {
            // Standard Customer Summary
            console.log(`\n--- CUSTOMER INCOME ---`);
            if (reports.customers.length === 0) console.log('(No Data)');
            else {
                const h = `Customer`.padEnd(30) + `Total`.padStart(15);
                console.log(h);
                console.log('-'.repeat(h.length));
                reports.customers.forEach(r => {
                    console.log(
                        `${r.label.substring(0, 29).padEnd(30)}` +
                        `${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`
                    );
                });
            }
        }
    }

    // Vendor sub report
    if (showAll || showVendorSub) {
        if (showVendorSub) {
            printDetailedTable('VENDOR SPENDING (Detailed)', reports.vendors, reportSheetList, sheetNameMap, "Vendor");
        } else {
            // Already shown in main "Vendor Spending" block if showVendor is on, 
            // but if showVendorSub is ON and showVendor is OFF (edge case), we show detailed.
            // If both, we might duplicate? 
            // The main vendor block shows 1099 info. This shows stats.
            // Let's assume if showVendor is on, we don't need this basic block unless sub is requested.
            // But the original code printed it.
        }
    }

    if (showDetails) {
        console.log(`\n--- DETAILS: "${targetDetailsFilter}" ---`);
        if (detailsRows.length === 0) {
            console.log('(No matching transactions found)');
        } else {
            console.log(`Date`.padEnd(12) + `Description`.padEnd(35) + `Sub-Cat`.padEnd(20) + `Amount`.padStart(12) + `  Source`);
            console.log(`-`.repeat(85));
            let total = 0;
            detailsRows.sort((a, b) => a.date.localeCompare(b.date));
            detailsRows.forEach(r => {
                total += r.amount;
                console.log(
                    `${r.date.padEnd(12)}${r.desc.substring(0, 34).padEnd(35)}${r.subCat.substring(0, 19).padEnd(20)}${r.amount.toFixed(2).padStart(12)}  ${r.sheet} (Row ${r.row})`
                );
            });
            console.log(`-`.repeat(85));
            console.log(`TOTAL`.padEnd(67) + total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(12));
        }
    }

    if (Math.abs(ctx.ledgerValidationTotal) > 0.01) {
        console.warn(`\n[!] CRITICAL WARNING: Ledger does not sum to zero!`);
        console.warn(`    Net Mismatch (Dr - Cr): ${ctx.ledgerValidationTotal.toFixed(2)}`);
        console.warn(`    Double-entry accounting requires Debits to equal Credits. Please check your Ledger entries.`);
        ctx.hasErrors = true;
    }

    Object.assign(ctx, { reportSheetList, sheetNameMap });
}

module.exports = { printStatements };

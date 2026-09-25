// Prints wallet reconciliation, the final status, data-integrity issues and
// configuration warnings, and sets a non-zero exit code when there are problems.

async function printDiagnostics(ctx) {
    const {
        customerReportUsage, duplicateCategories, illegalCategories, illegalCustomers,
        illegalSubCategories, illegalVendors, offsetWarnings, showChecker, taxYear,
        uncategorizedDetails, vendorReportUsage, walletCheckResults,
    } = ctx;

    // --- Wallet Reconciliation Report ---
    if (walletCheckResults.length > 0) {
        console.log('\n--- WALLET RECONCILIATION ---');
        console.log(`Sheet Name`.padEnd(25) + `Start`.padStart(12) + `Sheet Change`.padStart(15) + `Ledger Adj`.padStart(12) + `Calc End`.padStart(12) + `Exp End`.padStart(12) + `Diff`.padStart(12));
        console.log('-'.repeat(100));
        walletCheckResults.forEach(w => {
            if (!w.passed) ctx.hasErrors = true;
            const status = w.passed ? '' : ' [MISMATCH]';
            console.log(
                `${w.sheet.padEnd(25)}` +
                `${w.start.toFixed(2).padStart(12)}` +
                `${w.change.toFixed(2).padStart(15)}` +
                `${(w.ledgerPart || 0).toFixed(2).padStart(12)}` +
                `${w.calcEnd.toFixed(2).padStart(12)}` +
                `${w.expected.toFixed(2).padStart(12)}` +
                `${w.diff.toFixed(2).padStart(12)}` +
                `${status}`
            );
        });
    }



    if (ctx.skippedOutOfYear > 0) {
        console.log(`\n[Tax Year] Skipped ${ctx.skippedOutOfYear} transaction row(s) dated outside ${taxYear}. Use --year=YYYY to report on a different year.`);
    }

    // --- Final Status ---
    const totalIssues = (global.globalWarningCount || 0);
    if (!ctx.hasErrors && totalIssues === 0) {
        console.log(`\n✅ [ALL SYSTEMS GO] Financial reports are internally consistent and wallets reconcile (0 Warnings/Errors).`);
    } else {
        console.log(`\n❌ [CHECKS FAILED] System detected ${ctx.hasErrors ? 'CRITICAL ERRORS' : totalIssues + ' warnings'}. Please review logs above.`);
    }

    const hasIssues = uncategorizedDetails.length > 0 || illegalCategories.length > 0 || illegalVendors.length > 0 || illegalCustomers.length > 0 || illegalSubCategories.length > 0;
    // Non-zero exit when the books have errors or integrity issues, so scripts and CI can detect it.
    if (ctx.hasErrors || hasIssues) process.exitCode = 1;
    if (hasIssues) {
        console.log('\n--- DATA INTEGRITY ISSUES ---');
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet),
            ...illegalSubCategories.map(x => x.sheet)
        ]);
        issueSheetsFound.forEach(s => {
            console.log(`\n>> Tab: ${s.toUpperCase()}`);
            const uncat = uncategorizedDetails.filter(x => x.sheet === s);
            if (uncat.length) console.log(`  [!] ${uncat.length} rows missing category`);
            const cats = new Set(illegalCategories.filter(x => x.sheet === s).map(x => x.value));
            if (cats.size) console.log(`  [!] Illegal Categories: ${Array.from(cats).join(', ')}`);
            const vends = new Set(illegalVendors.filter(x => x.sheet === s).map(x => x.value));
            if (vends.size) console.log(`  [!] Unknown Vendors: ${Array.from(vends).join(', ')}`);
            const custs = new Set(illegalCustomers.filter(x => x.sheet === s).map(x => x.value));
            if (custs.size) console.log(`  [!] Unknown Customers: ${Array.from(custs).join(', ')}`);
            const subCats = new Set(illegalSubCategories.filter(x => x.sheet === s).map(x => x.value));
            if (subCats.size) console.log(`  [!] Illegal Sub-Categories: ${Array.from(subCats).join(', ')}`);

            if (showChecker) {
                uncat.forEach(x => console.log(`      - [${x.date}] Row ${x.row}: MISSING CATEGORY ("${x.desc}")`));
                illegalCategories.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL CATEGORY "${x.value}"`));
                illegalVendors.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN VENDOR "${x.value}"`));
                illegalCustomers.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN CUSTOMER "${x.value}"`));
                illegalSubCategories.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL SUB-CATEGORY "${x.value}" in category "${x.category}"`));
            }
        });

        if (illegalVendors.length > 0) {
            console.warn('\n[!] WARNING: Unknown Vendors Detected!');
            console.warn('The following vendors are not in your Setup sheet override list or vendor.xlsx:');
            const allUniqueVendors = Array.from(new Set(illegalVendors.map(x => x.value))).sort();
            allUniqueVendors.forEach(v => console.warn(` - "${v}"`));

            console.warn('\nPlease add these vendors to "Setup" or "vendor.xlsx" to proceed.');
            // Removed process.exit(1) to allow full report generation and multi-issue diagnostic
        }
    }

    // Output compliance errors if any
    if (global.complianceErrors && global.complianceErrors.length > 0) {
        console.error('\n--- COMPLIANCE ERRORS ---');
        global.complianceErrors.forEach(err => console.error(err));
    }

    if (duplicateCategories.length) {
        console.log('\n--- CATEGORY REPORT TYPE CONFLICTS ---');
        console.log('[!] The following categories have CONFLICTING Report types in your Setup sheet.');
        console.log('[!] Multiple rows with the same category but DIFFERENT subcategories is VALID.');
        console.log('[!] But all rows for a category must have the SAME Report type (P&L or BS).\n');
        duplicateCategories.forEach(d => {
            console.log(`  "${d.name}" (Row ${d.row}): Trying to set Report="${d.newReport}", but earlier row set it to "${d.existingReport}"`);
        });
        console.log('\nTo fix:');
        console.log('1. Open the Excel file and go to the Setup tab');
        console.log('2. Find all rows for the conflicting category');
        console.log('3. Ensure ALL rows have the SAME value in the Report column (either all "P&L" or all "BS")');
    }

    // Check for vendors/customers used in both P&L and BS
    const mixedVendors = [];
    const mixedCustomers = [];

    vendorReportUsage.forEach((reportMap, vendor) => {
        if (reportMap.size > 1) {
            mixedVendors.push({ name: vendor, reportMap });
        }
    });

    customerReportUsage.forEach((reportMap, customer) => {
        if (reportMap.size > 1) {
            mixedCustomers.push({ name: customer, reportMap });
        }
    });

    if (mixedVendors.length > 0 || mixedCustomers.length > 0) {
        console.log('\n--- VENDOR/CUSTOMER REPORT TYPE CONFLICTS ---');
        console.log('[!] The following vendors/customers are used in BOTH P&L and Balance Sheet categories.');
        console.log('[!] This usually indicates a data entry error.');
        console.log('[!] Vendors/customers should typically only appear in one report type.\n');

        if (mixedVendors.length > 0) {
            console.log('Vendors with mixed usage:');
            mixedVendors.forEach(v => {
                const reportTypes = Array.from(v.reportMap.keys());
                console.log(`\n  "${v.name}" appears in: ${reportTypes.join(' and ')}`);
                // Show example rows for each report type
                v.reportMap.forEach((examples, reportType) => {
                    console.log(`    ${reportType} examples:`);
                    examples.forEach(ex => {
                        console.log(`      - [${ex.date}] ${ex.sheet} Row ${ex.row}: Category="${ex.category}"`);
                    });
                });
            });
        }

        if (mixedCustomers.length > 0) {
            if (mixedVendors.length > 0) console.log('');
            console.log('Customers with mixed usage:');
            mixedCustomers.forEach(c => {
                const reportTypes = Array.from(c.reportMap.keys());
                console.log(`\n  "${c.name}" appears in: ${reportTypes.join(' and ')}`);
                // Show example rows for each report type
                c.reportMap.forEach((examples, reportType) => {
                    console.log(`    ${reportType} examples:`);
                    examples.forEach(ex => {
                        console.log(`      - [${ex.date}] ${ex.sheet} Row ${ex.row}: Category="${ex.category}"`);
                    });
                });
            });
        }

        console.log('\nTo fix:');
        console.log('1. Review your transactions for these vendors/customers');
        console.log('2. Ensure they are categorized consistently (all P&L or all BS)');
        console.log('3. If legitimately needed in both, consider using different vendor/customer names');
    }

    if (offsetWarnings.length) {
        console.log('\n--- OFFSET WARNINGS ---');
        offsetWarnings.forEach(w => console.log(`[!] Sheet "${w.sheet}" Row ${w.row} looks like a header (Found: ${w.matches.join(', ')}). Adjust Setup tab offset.`));
    }

    Object.assign(ctx, { hasIssues });
}

module.exports = { printDiagnostics };

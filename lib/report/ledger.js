// Processes the Ledger (manual double-entry) sheet and stops the run if it is
// unbalanced or has rows without a date.
const { parseAmount, getYear } = require('../accounting');
const { NORM_VEND, getVal } = require('./cells');
const { HEADERS, findCol } = require('./columns');

async function processLedger(ctx) {
    const {
        catStats, customerReportUsage, customerStats, detailsRows,
        illegalCategories, illegalCustomers, illegalSubCategories, illegalVendors, ledgerSheet,
        processedSheetTotals, sheetConfigs, showChecker, showDebug, showDetails, targetDetailsFilter, taxYear,
        uniqueCategories, validCategories, validCustomers, validSubCategories, validVendors,
        vendor1099Map, vendor1099Stats, vendorReportUsage, vendorStats,
    } = ctx;

    // Find Ledger configuration from Setup
    // Find Ledger configuration from Setup, or fallback to default
    ctx.ledgerConfig = sheetConfigs.find(c => c.name.toLowerCase() === 'general ledger' || c.name.toLowerCase() === 'ledger');

    if (!ctx.ledgerConfig) {
        if (showChecker) console.log(`[Info] No explicit "Ledger" config found in Setup. Using default configuration with auto-detected header.`);

        // Auto-detect header row
        let detectedOffset = 3;
        ledgerSheet.eachRow((row, r) => {
            if (r > 10) return; // Only scan first 10 rows
            const values = row.values.map(v => (v ? v.toString().toLowerCase() : ''));
            const hasDate = values.some(v => v.includes('date'));
            const hasCat = values.some(v => v.includes('category'));
            if (hasDate && hasCat) {
                detectedOffset = r;
            }
        });

        ctx.ledgerConfig = {
            name: ledgerSheet.name,
            type: 'ledger',
            flip: false,
            offset: detectedOffset,
            startBalance: 0,
            endBalance: null,
            linkedAccount: null
        };
        sheetConfigs.push(ctx.ledgerConfig);
    }

    const ledgerHeaderRow = ctx.ledgerConfig.offset || 3;

    // Dynamic Mapping for Ledger
    const ledgerMap = { date: null, desc: null, category: null, subCat: null, vendor: null, customer: null, dr: null, cr: null };
    const ledgerHeader = ledgerSheet.getRow(ledgerHeaderRow);

    ledgerHeader.eachCell((cell, colNumber) => {
        const val = getVal(cell);

        if (findCol(val, HEADERS.DATE)) ledgerMap.date = colNumber;
        else if (findCol(val, HEADERS.DESC)) ledgerMap.desc = colNumber;
        else if (findCol(val, HEADERS.SUBCAT)) ledgerMap.subCat = colNumber;
        else if (findCol(val, HEADERS.CATEGORY)) ledgerMap.category = colNumber;
        else if (findCol(val, HEADERS.VENDOR)) ledgerMap.vendor = colNumber;
        else if (findCol(val, HEADERS.CUSTOMER)) ledgerMap.customer = colNumber;
        else if (findCol(val, HEADERS.DEBIT)) ledgerMap.dr = colNumber;
        else if (findCol(val, HEADERS.CREDIT)) ledgerMap.cr = colNumber;
    });

    ctx.ledgerValidationTotal = 0; // Raw Dr - Cr, must be 0
    let ledgerDebitTotal = 0;
    let ledgerCreditTotal = 0;
    let ledgerRows = 0;

    ledgerSheet.eachRow((row, r) => {
        try {
            if (r <= ledgerHeaderRow) return;
            const rawDate = ledgerMap.date ? getVal(row.getCell(ledgerMap.date)) : '';
            const rawDesc = ledgerMap.desc ? getVal(row.getCell(ledgerMap.desc)) : '';
            const cat = ledgerMap.category ? getVal(row.getCell(ledgerMap.category)) : '';

            // Ledger SubCat support
            const subCatVal = ledgerMap.subCat ? getVal(row.getCell(ledgerMap.subCat)) : '';

            const dr = (ledgerMap.dr && row.getCell(ledgerMap.dr).value) ? (parseAmount(getVal(row.getCell(ledgerMap.dr))) || 0) : 0;
            const cr = (ledgerMap.cr && row.getCell(ledgerMap.cr).value) ? (parseAmount(getVal(row.getCell(ledgerMap.cr))) || 0) : 0;

            const vendorVal = ledgerMap.vendor ? getVal(row.getCell(ledgerMap.vendor)) : '';
            const customerVal = ledgerMap.customer ? getVal(row.getCell(ledgerMap.customer)) : '';

            if (!rawDate) {
                if (cat || dr || cr || rawDesc) {
                    console.error(`\n[CRITICAL ERROR] Ledger Row ${r}: Missing Date!`);
                    console.error(`In the General Ledger, every transaction row must have an explicit Date. (Values: ${rawDesc} | ${cat})`);
                    console.error(`Please fix the Ledger sheet and try again.`);
                    process.exit(1);
                }
                return; // Truly empty row
            }

            // Skip ledger entries outside the selected tax year
            const ledgerYear = getYear(rawDate);
            if (ledgerYear !== null && ledgerYear !== taxYear) {
                ctx.skippedOutOfYear++;
                return;
            }

            ledgerDebitTotal += dr;
            ledgerCreditTotal += cr;
            // Accumulate validation total (Dr should equal Cr, so Dr - Cr should be 0 across all rows)
            ctx.ledgerValidationTotal += (dr - cr);

            // Vendor Validation
            if (vendorVal && vendorVal.toString().trim()) {
                const vStr = vendorVal.toString().trim();
                const vLower = NORM_VEND(vStr);
                // Using 'General Ledger' or dynamic name? ledgerConfig.name is better but 'Ledger' is hardcoded here in context
                if (!validVendors.has(vLower)) illegalVendors.push({ value: vStr, sheet: 'Ledger', row: r, date: (rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A')) });
            }

            if (cat) {
                const catStr = cat.toString().trim();
                const catLower = catStr.toLowerCase();
                const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');

                if (!validCategories.has(catLower)) illegalCategories.push({ value: catStr, sheet: 'Ledger', row: r, date: displayDate });

                const conf = uniqueCategories.get(catLower);

                const displayCat = conf?.displayName || catStr;
                const isAsset = conf && conf.accountType && conf.accountType.toLowerCase().includes('asset');
                const impact = isAsset ? (dr - cr) : (cr - dr);

                if (impact !== 0 || catStr) {
                    ledgerRows++;
                    processedSheetTotals[ctx.ledgerConfig.name] = (processedSheetTotals[ctx.ledgerConfig.name] || 0) + impact;
                }

                if (!catStats[displayCat]) catStats[displayCat] = { total: 0, subCats: {}, sheets: {} };
                if (!catStats[displayCat].sheets) catStats[displayCat].sheets = {};
                catStats[displayCat].total += impact;
                // Track Sheet contribution (using canonical name)
                if (!catStats[displayCat].sheets[ctx.ledgerConfig.name]) catStats[displayCat].sheets[ctx.ledgerConfig.name] = { add: 0, sub: 0, total: 0 };
                const lSheet = catStats[displayCat].sheets[ctx.ledgerConfig.name];
                if (impact >= 0) lSheet.add += impact; else lSheet.sub += impact;
                lSheet.total += impact;

                // Ledger SubCat aggregation
                const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';

                // Validate subcategory if one is provided
                if (subCatVal && sName !== '(No Sub-Cat)') {
                    const sLower = sName.toLowerCase();
                    if (!validSubCategories.has(sLower)) {
                        illegalSubCategories.push({
                            value: sName,
                            category: displayCat,
                            sheet: 'Ledger',
                            row: r,
                            date: displayDate
                        });
                    }
                }

                catStats[displayCat].subCats[sName] = (catStats[displayCat].subCats[sName] || 0) + impact;

                // Capture Details (same matching as transaction sheets: category, vendor or customer)
                const detailsMatch = showDetails && (
                    catLower === targetDetailsFilter || displayCat.toLowerCase() === targetDetailsFilter ||
                    (vendorVal && NORM_VEND(vendorVal.toString()) === targetDetailsFilter) ||
                    (customerVal && NORM_VEND(customerVal.toString()) === targetDetailsFilter));
                if (detailsMatch) {
                    detailsRows.push({
                        date: displayDate,
                        desc: rawDesc,
                        subCat: sName,
                        amount: impact,
                        sheet: 'Ledger',
                        row: r
                    });
                }

                // Vendor / Customer Stats from Ledger
                if (vendorVal && vendorVal.toString().trim()) {
                    const vStr = vendorVal.toString().trim();
                    const vLower = NORM_VEND(vStr);
                    if (!validVendors.has(vLower)) illegalVendors.push({ value: vStr, sheet: 'Ledger', row: r, date: displayDate });

                    const displayVendor = validVendors.get(vLower) || vStr;
                    // Vendor: Net Debit (Expense)
                    // Vendor Stats Structure: { total, add, sub, sheets }
                    // Vendor Stats Structure: { total, add, sub, sheets }
                    const impactVal = (dr - cr);
                    if (!vendorStats[displayVendor]) vendorStats[displayVendor] = { total: 0, add: 0, sub: 0, sheets: {} };
                    if (!vendorStats[displayVendor].sheets) vendorStats[displayVendor].sheets = {};

                    if (impactVal >= 0) vendorStats[displayVendor].add += impactVal; else vendorStats[displayVendor].sub += impactVal;
                    vendorStats[displayVendor].total += impactVal;

                    // Ensure sheets exists before accessing
                    if (!vendorStats[displayVendor].sheets) vendorStats[displayVendor].sheets = {};
                    if (!vendorStats[displayVendor].sheets[ctx.ledgerConfig.name]) vendorStats[displayVendor].sheets[ctx.ledgerConfig.name] = { add: 0, sub: 0, total: 0 };
                    const vSheet = vendorStats[displayVendor].sheets[ctx.ledgerConfig.name];
                    if (impactVal >= 0) vSheet.add += impactVal; else vSheet.sub += impactVal;
                    vSheet.total += impactVal;

                    // Track which report type this vendor is used in
                    const catConf = uniqueCategories.get(catLower);
                    if (catConf && catConf.report) {
                        if (!vendorReportUsage.has(displayVendor)) {
                            vendorReportUsage.set(displayVendor, new Map());
                        }
                        const reportMap = vendorReportUsage.get(displayVendor);
                        if (!reportMap.has(catConf.report)) {
                            reportMap.set(catConf.report, []);
                        }
                        // Store first 2 examples per report type
                        if (reportMap.get(catConf.report).length < 2) {
                            reportMap.get(catConf.report).push({
                                sheet: 'Ledger',
                                row: r,
                                category: displayCat,
                                date: displayDate
                            });
                        }
                    }

                    const is1099 = vendor1099Map.get(vLower);
                    if (is1099) {
                        const t = (is1099.type || 'NEC');
                        if (!vendor1099Stats[t]) vendor1099Stats[t] = {};
                        if (!vendor1099Stats[t][displayVendor]) {
                            vendor1099Stats[t][displayVendor] = 0;
                        }
                        vendor1099Stats[t][displayVendor] += impactVal;
                    }
                    // Track Vendor Sub-Categories in Ledger
                    if (!vendorStats[displayVendor].subCats) vendorStats[displayVendor].subCats = {};
                    vendorStats[displayVendor].subCats[sName] = (vendorStats[displayVendor].subCats[sName] || 0) + impactVal;
                }

                if (customerVal && customerVal.toString().trim()) {
                    const cStr = customerVal.toString().trim();
                    const cLower = cStr.toLowerCase();
                    if (!validCustomers.has(cLower)) illegalCustomers.push({ value: cStr, sheet: 'Ledger', row: r, date: displayDate });

                    const displayCustomer = validCustomers.get(cLower) || cStr;
                    // Customer: Net Credit (Income)
                    const custImpact = (cr - dr);
                    if (!customerStats[displayCustomer]) customerStats[displayCustomer] = { total: 0, add: 0, sub: 0, sheets: {} };
                    if (!customerStats[displayCustomer].sheets) customerStats[displayCustomer].sheets = {};

                    if (custImpact >= 0) customerStats[displayCustomer].add += custImpact; else customerStats[displayCustomer].sub += custImpact;
                    customerStats[displayCustomer].total += custImpact;

                    if (!customerStats[displayCustomer].sheets[ctx.ledgerConfig.name]) customerStats[displayCustomer].sheets[ctx.ledgerConfig.name] = { add: 0, sub: 0, total: 0 };
                    const cSheet = customerStats[displayCustomer].sheets[ctx.ledgerConfig.name];
                    if (custImpact >= 0) cSheet.add += custImpact; else cSheet.sub += custImpact;
                    cSheet.total += custImpact;

                    // Track Customer Sub-Categories in Ledger
                    if (!customerStats[displayCustomer].subCats) customerStats[displayCustomer].subCats = {};
                    const cSub = sName || '(No Sub-Cat)';
                    customerStats[displayCustomer].subCats[cSub] = (customerStats[displayCustomer].subCats[cSub] || 0) + custImpact;

                    // Track which report type this customer is used in
                    const catConf = uniqueCategories.get(catLower);
                    if (catConf && catConf.report) {
                        if (!customerReportUsage.has(displayCustomer)) {
                            customerReportUsage.set(displayCustomer, new Map());
                        }
                        const reportMap = customerReportUsage.get(displayCustomer);
                        if (!reportMap.has(catConf.report)) {
                            reportMap.set(catConf.report, []);
                        }
                        // Store first 2 examples per report type
                        if (reportMap.get(catConf.report).length < 2) {
                            reportMap.get(catConf.report).push({
                                sheet: 'Ledger',
                                row: r,
                                category: displayCat,
                                date: displayDate
                            });
                        }
                    }
                }

                // (Integration block removed - handled via standard catStats logic)
            }
        } catch (ledgerRowError) {
            // Always report: a row that fails part-way would leave totals incomplete.
            console.error(`[ERROR] Ledger Row ${r}: ${ledgerRowError.message}`);
            if (showDebug) console.error(ledgerRowError.stack);
            ctx.hasErrors = true;
        }
    });

    const isBalanced = Math.abs(ctx.ledgerValidationTotal) < 0.01;
    const ledgerStatus = isBalanced ? `Balanced (Volume: ${ledgerDebitTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })})` : `UNBALANCED by ${ctx.ledgerValidationTotal.toFixed(2)}`;
    if (showDebug) console.log(`[Sheet Stats] "${ctx.ledgerConfig.name}": Processed ${ledgerRows} rows with data. Status: ${ledgerStatus}.`);

    if (!isBalanced) {
        console.error(`\n[CRITICAL ERROR] Ledger is UNBALANCED!`);
        console.error(`Total Debits ($${ledgerDebitTotal.toFixed(2)}) minus total Credits ($${ledgerCreditTotal.toFixed(2)}) = ${ctx.ledgerValidationTotal.toFixed(2)} (should be 0.00).`);
        console.error(`Please fix the Ledger sheet and try again.`);
        process.exit(1);
    }
}

module.exports = { processLedger };

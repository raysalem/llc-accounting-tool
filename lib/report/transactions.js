// Processes every configured transaction sheet: maps columns, applies polarity and
// the tax-year filter, accumulates category/vendor/customer totals and records
// data-integrity issues, then applies each sheet total to its linked account.
const { getYear, parseAmount } = require('../accounting');
const { NORM_VEND, getVal } = require('./cells');
const { HEAD_MATCH } = require('./columns');
const { detectSheetColumns } = require('./sheet-columns');
const { applySheetLinkage } = require('./linkage');

async function processTransactionSheets(ctx) {
    const {
        catStats, customerReportUsage, customerStats, detailsRows, illegalCategories,
        illegalCustomers, illegalSubCategories, illegalVendors, offsetWarnings,
        processedSheetTotals, sheetConfigs, showChecker, showDebug, showDetails,
        targetDetailsFilter, taxYear, uncategorizedDetails, uniqueCategories, validCategories,
        validCustomers, validSubCategories, validVendors, vendor1099Map, vendor1099Stats,
        vendorReportUsage, vendorStats, workbook,
    } = ctx;

    const sheetTotalMap = new Map();
    for (const config of sheetConfigs) {
        let sheetTotal = 0;
        let sheetAdds = 0;
        let sheetSubs = 0;
        let sheet = workbook.getWorksheet(config.name);
        if (!sheet) {
            sheet = workbook.worksheets.find(s => s.name.trim().toLowerCase() === config.name.trim().toLowerCase());
        }
        if (!sheet) {
            if (showChecker) console.log(`Sheet "${config.name}" NOT FOUND`);
            continue;
        }

        // --- Skip Ledger in main loop (handled separately in Step 3) ---
        if (config.name.toLowerCase() === 'ledger' || config.name.toLowerCase() === 'general ledger' || config.type.toLowerCase() === 'ledger') {
            continue;
        }

        const tStr = config.type.toLowerCase();
        const isCC = tStr.includes('cc') || tStr.includes('card') || tStr.includes('credit') || tStr.includes('amex');
        const pType = isCC ? 'cc' : 'bank';

        const map = detectSheetColumns(sheet, config, isCC, { showDebug, showChecker });

        let processedRows = 0;
        let excludedFromLinkage = 0;
        let excludedAdds = 0;
        let excludedSubs = 0;

        if (showDebug && sheet.name.includes('AX CC')) {
            console.log(`[DEBUG AX CC] Sheet Found. Config Offset: ${config.offset}. Total Rows in Sheet: ${sheet.rowCount}`);
        }

        sheet.eachRow((row, r) => {
            try {
                if (r <= config.offset) {
                    // if (sheet.name.includes('AX CC')) console.log(`[AX CC] Skipping Row ${r} (<= Offset ${config.offset})`);
                    return;
                }

                const vendorVal = map.vendor ? getVal(row.getCell(map.vendor)) : '';
                const customerVal = map.customer ? getVal(row.getCell(map.customer)) : '';
                const categoryVal = map.category ? getVal(row.getCell(map.category)) : '';
                const subCatVal = map.subCat ? getVal(row.getCell(map.subCat)) : '';
                let amount = map.amount ? getVal(row.getCell(map.amount)) : 0;

                // --- Initialize stats structures if needed ---
                // Category stats (for PL/BS) with per-sheet breakdown
                // --- REDUNDANT BLOCK REMOVED ---
                // The previous logic here (Lines 713-732) was using catLower as key, which conflicts with later logic using DisplayName.
                // We strip this block and rely on the consolidated logic below.
                // -------------------------------

                // Vendor / Customer stats tracking (consolidated below)


                // Customer stats with per-sheet breakdown (existing logic retained below)

                if (typeof amount !== 'number') {
                    // Sanitize string (remove $, commas, etc)
                    const amtStr = amount.toString().trim();
                    if (amtStr === '' || amtStr === '-') {
                        amount = 0;
                    } else {
                        const parsed = parseAmount(amtStr);
                        if (isNaN(parsed)) {
                            if (showChecker && amtStr.length > 0) {
                                console.warn(`[WARNING] Sheet "${sheet.name}" Row ${r}: Could not parse amount "${amount}". Skipping.`);
                            }
                            return; // Skip this row safely
                        }
                        amount = parsed;
                    }
                }

                function excelDateToJS(serial) {
                    if (typeof serial !== 'number') return serial;
                    // Excel date offset: Jan 1, 1900.
                    // Note: Excel incorrectly treats 1900 as a leap year, so we subtract 2.
                    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
                    return date;
                }

                const rawDateInput = map.date ? getVal(row.getCell(map.date)) : '';
                let dateObj = rawDateInput;
                if (typeof rawDateInput === 'number' && rawDateInput > 10000) {
                    dateObj = excelDateToJS(rawDateInput);
                }

                const rawDesc = map.desc ? getVal(row.getCell(map.desc)).toString() : '';
                const matchDesc = rawDesc.toLowerCase();

                const displayDate = dateObj instanceof Date ? dateObj.toISOString().split('T')[0] :
                    (dateObj && typeof dateObj === 'string' ? dateObj : 'N/A');



                // Offset check
                if (r === config.offset + 1) {
                    const rowValues = row.values.map(v => (v ? v.toString().toLowerCase() : ''));
                    const rowText = rowValues.join(' ');
                    if (HEAD_MATCH(rowText)) {
                        offsetWarnings.push({ sheet: sheet.name, row: r, matches: ['Header Signature Detected'] });
                    }
                }

                // --- Robust Skipping Logic ---
                const hasDesc = rawDesc.trim().length > 0;
                const hasAmount = Math.abs(amount) > 0.0001;

                // 1. If completely empty, skip silently
                if (!rawDateInput && !hasDesc && !hasAmount) return;

                // 2. Strict Date Check
                if (displayDate === 'N/A' || !rawDateInput) {
                    // Only warn if there is other data
                    if ((hasDesc || hasAmount) && showChecker) {
                        console.warn(`[WARNING] Sheet "${sheet.name}" Row ${r}: Skipped due to invalid/missing Date. (Desc: "${rawDesc}", Amt: ${amount})`);
                    }
                    return;
                }

                // 3. Skip invalid amounts (already handled above, but double check)
                if (isNaN(amount)) return;

                // 4. Skip rows outside the selected tax year
                const rowYear = getYear(dateObj);
                if (rowYear !== null && rowYear !== taxYear) {
                    ctx.skippedOutOfYear++;
                    return;
                }

                processedRows++;

                if (config.flip) amount *= -1;

                // Accumulate Sheet Total (Net Flow)
                // Note: Net Flow affects the Asset/Liability Balance linked to this sheet.
                // However, we must EXCLUDE "Transfer" rows from this impact if they are meant to be ignored.
                if (amount >= 0) sheetAdds += amount; else sheetSubs += amount;
                sheetTotal += amount;
                if (pType === 'cc') ctx.ccTotal += amount; else ctx.bankTotal += amount;

                // Track global processed total for this sheet
                processedSheetTotals[config.name] = (processedSheetTotals[config.name] || 0) + amount;

                if (!categoryVal && Math.abs(amount) > 0.01) {
                    if (pType === 'cc') ctx.uncategorizedCC++; else ctx.uncategorizedBank++;
                    uncategorizedDetails.push({ sheet: config.shortName || sheet.name, row: r, date: displayDate, desc: rawDesc });
                }

                // Define catLower for use in vendor/customer tracking
                const catLower = categoryVal ? categoryVal.toString().trim().toLowerCase() : null;

                // Check for EXCLUSION (Transfer)
                // If a row is categorized as 'Transfer', it should NOT impact the Linked Liability Account Balance?
                // Logic: 
                // CC Sheet: Payment (+100). Category "Transfer". 
                // If we include +100 in sheetTotal, it reduces Liability.
                // If we WANT to ignore it, we must subtract 100 from sheetTotal before applying to Link.
                if (catLower) {
                    const cData = uniqueCategories.get(catLower);
                    if (cData) {
                        if (cData.report === 'Transfer') {
                            excludedFromLinkage += amount;
                            if (amount >= 0) excludedAdds += amount; else excludedSubs += amount;

                            if (showDebug) console.log(`[Transfer Logic] Excluding ${amount.toFixed(2)} from Linkage. Cat: "${cData.displayName}"`);
                        } else if (cData.displayName.includes('Transfer')) {
                            // Debugging why it missed?
                            if (showDebug) console.log(`[Transfer Logic Debug] Found "${cData.displayName}" but Report Type is "${cData.report}" (Expected 'Transfer'). Amount: ${amount}`);
                        }
                    } else if (catLower.includes('transfer')) {
                        if (showDebug) console.log(`[Transfer Logic Debug] Category "${catLower}" NOT FOUND in Config. Falling back to default handling.`);
                    }

                    if (cData && cData.report === 'Transfer') {
                        // Already handled above, removed dup
                        // Proceed to validation
                        if (cData.transferAccount) {
                            const targetAccount = cData.transferAccount.toLowerCase();
                            const currentLink = (config.linkedAccount || '').toLowerCase();
                            const currentName = (config.name || '').toLowerCase();
                            // We expect 'Transfer AX CC' to be used ONLY in the sheet representing that account
                            if (!currentLink.includes(targetAccount) && !currentName.includes(targetAccount)) {
                                console.warn(`[!] CRITICAL WARNING (Row ${r}): Category "${cData.displayName}" expects Transfer Account "${cData.transferAccount}", but used in Sheet "${config.name}" (Linked: ${config.linkedAccount}). This may be incorrect.`);
                            }
                        }
                    }
                }

                const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';

                if (categoryVal) {
                    const catStr = categoryVal.toString().trim();

                    if (!validCategories.has(catLower)) {
                        illegalCategories.push({ value: catStr, sheet: sheet.name, row: r, date: displayDate });
                    }

                    // Use Display Name for stats if available, else usage case
                    const displayCat = uniqueCategories.get(catLower)?.displayName || catStr;

                    if (!catStats[displayCat]) catStats[displayCat] = { total: 0, subCats: {}, sheets: {} };

                    // Asset Polarity Logic: 
                    // If target is an ASSET, we want Spending (Bank Withdrawals) to be POSITIVE (Increase in Asset).
                    // If source is CC (Liability), Spending is already Positive.
                    // If source is Bank (Asset), Spending is Negative. So we Flip Bank sources for Assets.
                    let catAmount = amount;
                    const catData = uniqueCategories.get(catLower);
                    if (catData && catData.report === 'Balance Sheet') {
                        const type = (catData.accountType || '').toLowerCase();
                        if (type.includes('asset')) {
                            // If config.flip is FALSE (Bank), then Spending is Negative. Invert it to Positive.
                            // If config.flip is TRUE (CC), then Spending is Positive. Keep it Positive.
                            if (!config.flip) {
                                catAmount = amount * -1;
                            }
                        }
                    }

                    catStats[displayCat].total += catAmount;

                    if (catAmount >= 0) {
                        if (!catStats[displayCat].add) catStats[displayCat].add = 0;
                        catStats[displayCat].add = (catStats[displayCat].add || 0) + catAmount;
                    } else {
                        if (!catStats[displayCat].sub) catStats[displayCat].sub = 0;
                        catStats[displayCat].sub = (catStats[displayCat].sub || 0) + catAmount;
                    }

                    // Track Sheets for Display Category
                    if (!catStats[displayCat].sheets) catStats[displayCat].sheets = {};
                    if (!catStats[displayCat].sheets[config.name]) {
                        catStats[displayCat].sheets[config.name] = { add: 0, sub: 0, total: 0 };
                    }
                    const cDS = catStats[displayCat].sheets[config.name];
                    if (catAmount >= 0) cDS.add += catAmount; else cDS.sub += catAmount;
                    cDS.total += catAmount;

                    // Validate subcategory if one is provided
                    if (subCatVal && sName !== '(No Sub-Cat)') {
                        const sLower = sName.toLowerCase();
                        if (!validSubCategories.has(sLower)) {
                            illegalSubCategories.push({
                                value: sName,
                                category: displayCat,
                                sheet: config.shortName || sheet.name,
                                row: r,
                                date: displayDate
                            });
                        }
                    }

                    catStats[displayCat].subCats[sName] = (catStats[displayCat].subCats[sName] || 0) + catAmount;

                    // Capture Details

                }

                if (vendorVal && vendorVal.toString().trim()) {
                    const vStr = vendorVal.toString().trim();
                    const vLower = NORM_VEND(vStr);
                    if (!validVendors.has(vLower)) illegalVendors.push({ value: vStr, sheet: sheet.name, row: r, date: displayDate });

                    const displayVendor = validVendors.get(vLower) || vStr;
                    // Standardize: Expenses are Positive. Income is Negative.
                    const vendAmount = amount * -1;

                    if (!vendorStats[displayVendor]) {
                        vendorStats[displayVendor] = { total: 0, add: 0, sub: 0, sheets: {} };
                    }
                    if (!vendorStats[displayVendor].sheets) vendorStats[displayVendor].sheets = {};

                    if (vendAmount >= 0) {
                        vendorStats[displayVendor].add += vendAmount;
                    } else {
                        vendorStats[displayVendor].sub += vendAmount;
                    }
                    vendorStats[displayVendor].total += vendAmount;

                    if (!vendorStats[displayVendor].sheets[config.name]) {
                        vendorStats[displayVendor].sheets[config.name] = { add: 0, sub: 0, total: 0 };
                    }
                    const vSheet = vendorStats[displayVendor].sheets[config.name];
                    if (vendAmount >= 0) vSheet.add += vendAmount; else vSheet.sub += vendAmount;
                    vSheet.total += vendAmount;


                    // Track which report type this vendor is used in
                    if (catLower) {
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
                            const dCatLabel = catConf.displayName || categoryVal.toString().trim();
                            if (reportMap.get(catConf.report).length < 2) {
                                reportMap.get(catConf.report).push({
                                    sheet: sheet.name,
                                    row: r,
                                    category: dCatLabel,
                                    date: displayDate
                                });
                            }
                        }
                    }

                    const is1099 = vendor1099Map.get(vLower);
                    if (is1099) {
                        const t = (is1099.type || 'NEC');
                        if (!vendor1099Stats[t]) vendor1099Stats[t] = {};
                        if (!vendor1099Stats[t][displayVendor]) {
                            vendor1099Stats[t][displayVendor] = 0;
                        }
                        vendor1099Stats[t][displayVendor] += vendAmount;
                    }

                    // Track Vendor Sub-Categories
                    if (!vendorStats[displayVendor].subCats) vendorStats[displayVendor].subCats = {};
                    vendorStats[displayVendor].subCats[sName] = (vendorStats[displayVendor].subCats[sName] || 0) + vendAmount;
                }
                if (customerVal && customerVal.toString().trim()) {
                    const cStr = customerVal.toString().trim();
                    const cLower = NORM_VEND(cStr);
                    if (!validCustomers.has(cLower)) illegalCustomers.push({ value: cStr, sheet: sheet.name, row: r, date: displayDate });

                    const displayCustomer = validCustomers.get(cLower) || cStr;

                    // Existing customer stats logic (kept unchanged)
                    if (!customerStats[displayCustomer]) {
                        customerStats[displayCustomer] = { add: 0, sub: 0, total: 0, sheets: {} };
                    }
                    if (amount >= 0) {
                        customerStats[displayCustomer].add += amount;
                    } else {
                        customerStats[displayCustomer].sub += amount;
                    }
                    customerStats[displayCustomer].total += amount;
                    if (!customerStats[displayCustomer].sheets[config.name]) {
                        customerStats[displayCustomer].sheets[config.name] = { add: 0, sub: 0, total: 0 };
                    }
                    const cSheetStat = customerStats[displayCustomer].sheets[config.name];
                    if (amount >= 0) cSheetStat.add += amount; else cSheetStat.sub += amount;
                    cSheetStat.total += amount;

                    // Track Customer Sub-Categories
                    if (!customerStats[displayCustomer].subCats) customerStats[displayCustomer].subCats = {};
                    const cSub = sName || '(No Sub-Cat)';
                    customerStats[displayCustomer].subCats[cSub] = (customerStats[displayCustomer].subCats[cSub] || 0) + amount;

                    // Track which report type this customer is used in
                    if (catLower) {
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
                            const displayCat = catConf.displayName || categoryVal.toString().trim();
                            if (reportMap.get(catConf.report).length < 2) {
                                reportMap.get(catConf.report).push({
                                    sheet: sheet.name,
                                    row: r,
                                    category: displayCat,
                                    date: displayDate
                                });
                            }
                        }
                    }
                }
                // Capture Details (Generic Filter)
                if (showDetails) {
                    const cNorm = catLower;
                    const cDisp = cNorm ? (uniqueCategories.get(cNorm)?.displayName || cNorm) : null;
                    const vNorm = vendorVal ? NORM_VEND(vendorVal.toString()) : null;
                    const custNorm = customerVal ? NORM_VEND(customerVal.toString()) : null;

                    const match = (cNorm === targetDetailsFilter) || (cDisp && cDisp.toLowerCase() === targetDetailsFilter) ||
                        (vNorm === targetDetailsFilter) || (custNorm === targetDetailsFilter);

                    if (match) {
                        detailsRows.push({
                            date: displayDate,
                            desc: rawDesc,
                            subCat: subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)',
                            amount: amount,
                            sheet: sheet.name,
                            row: r
                        });
                    }
                }
            } catch (rowError) {
                if (showChecker) console.error(`[ERROR] Sheet "${sheet.name}" Row ${r}: Crash detected. ${rowError.message}`);
            }
        });

        if (showDebug) console.log(`[Sheet Stats] "${config.shortName}": Processed ${processedRows} rows. Total Change: ${sheetTotal.toFixed(2)}`);
        sheetTotalMap.set(config.name, sheetTotal);



        applySheetLinkage(ctx, config, { sheetTotal, sheetAdds, sheetSubs, excludedFromLinkage, excludedAdds, excludedSubs });
    }

}

module.exports = { processTransactionSheets };

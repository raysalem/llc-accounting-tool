const ExcelJS = require('exceljs');
const fs = require('fs');

async function updateFinancials() {
    const args = process.argv.slice(2);
    const printOnly = args.includes('--print-only');
    const showPL = args.includes('--pl');
    const showBS = args.includes('--bs');
    const showVendor = args.includes('--vendor');
    const showCustomer = args.includes('--customer');
    const showPLSub = args.includes('--pl-sub');
    const showChecker = args.includes('--checker');

    const specificFilter = showPL || showBS || showVendor || showCustomer || showPLSub || showChecker;
    const showAll = printOnly && !specificFilter;

    let filename = args.find(a => !a.startsWith('--')) || 'LLC_Accounting_Template.xlsx';
    if (!fs.existsSync(filename)) {
        console.error(`Error: File '${filename}' not found.`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    try {
        if (showChecker) console.log(`Loading workbook: ${filename}...`);
        await workbook.xlsx.readFile(filename);
    } catch (e) {
        console.error(`Error reading file: ${e.message}`);
        return;
    }

    const setupSheet = workbook.getWorksheet('Setup');
    const ledgerSheet = workbook.getWorksheet('Ledger');
    let summarySheet = workbook.getWorksheet('Summary');

    if (!setupSheet || !ledgerSheet) {
        console.error('Error: Mandatory sheets (Setup or Ledger) missing.');
        return;
    }

    // --- State ---
    const validCategories = new Set();
    const validVendors = new Set();
    const validCustomers = new Set();
    const uniqueCategories = new Map();
    const sheetConfigs = [];

    const catStats = {};
    const vendorStats = {};
    const customerStats = {};
    let bankTotal = 0;
    let ccTotal = 0;
    let uncategorizedBank = 0;
    let uncategorizedCC = 0;

    const illegalCategories = [];
    const illegalVendors = [];
    const illegalCustomers = [];
    const uncategorizedDetails = [];
    const offsetWarnings = [];

    // --- Helper ---
    function getVal(cell) {
        if (!cell) return '';
        let v = cell.value;
        if (v && typeof v === 'object') {
            if (v instanceof Date) return v;
            if (v.result !== undefined) {
                if (v.result === null || v.result === undefined) return '';
                return v.result;
            }
            if (v.richText) return v.richText.map(t => t.text).join('').trim();
            if (v.text) return v.text.toString().trim();
            if (v.formula) return '';
            return v.toString().trim();
        }
        return (v === null || v === undefined) ? '' : v.toString().trim();
    }

    // --- 1. Read Setup ---
    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const catName = getVal(row.getCell(1));
        const accountType = getVal(row.getCell(2)); // Account Type column
        const report = getVal(row.getCell(4));
        if (catName) {
            const trimmed = catName.toString().trim();
            validCategories.add(trimmed);
            uniqueCategories.set(trimmed, { report, accountType });
        }
        const vendor = getVal(row.getCell(6));
        if (vendor) validVendors.add(vendor.toString().trim());
        const customer = getVal(row.getCell(7));
        if (customer) validCustomers.add(customer.toString().trim());

        const confSheetName = getVal(row.getCell(9));
        const confType = getVal(row.getCell(10));
        const confFlip = getVal(row.getCell(11));
        const confOffset = getVal(row.getCell(12));

        if (confSheetName && confType) {
            sheetConfigs.push({
                name: confSheetName.toString().trim(),
                type: confType.toString().trim(),
                flip: !!(confFlip && confFlip.toString().toLowerCase().includes('y')),
                offset: parseInt(confOffset) || 0
            });
        }
    });

    if (sheetConfigs.length === 0) {
        sheetConfigs.push({ name: 'Bank Transactions', type: 'Bank', flip: false, offset: 1 });
        sheetConfigs.push({ name: 'Credit Card Transactions', type: 'CC', flip: true, offset: 1 });
    }

    // --- 2. Process Transaction Sheets ---
    const bankMapDefault = { date: 1, desc: 2, amount: 3, category: 4, subCat: 5, vendor: 7, customer: 8 };
    const ccMapDefault = { date: 1, desc: 3, amount: 4, category: 5, subCat: 6, vendor: 8, customer: 9 };

    for (const config of sheetConfigs) {
        let sheet = workbook.getWorksheet(config.name);
        if (!sheet) {
            sheet = workbook.worksheets.find(s => s.name.trim().toLowerCase() === config.name.trim().toLowerCase());
        }
        if (!sheet) {
            if (showChecker) console.log(`Sheet "${config.name}" NOT FOUND`);
            continue;
        }

        const tStr = config.type.toLowerCase();
        const isCC = tStr.includes('cc') || tStr.includes('card') || tStr.includes('credit') || tStr.includes('amex');
        const pType = isCC ? 'cc' : 'bank';

        // Dynamic Map detection
        const headerRowIndex = config.offset || 1;
        const headerRow = sheet.getRow(headerRowIndex);
        const map = isCC ? { ...ccMapDefault } : { ...bankMapDefault };

        headerRow.eachCell((cell, colNumber) => {
            const val = getVal(cell).toString().toLowerCase();
            if (val === 'date') map.date = colNumber;
            else if (val === 'description' || val === 'desc') map.desc = colNumber;
            else if (val === 'amount') map.amount = colNumber;
            else if (val === 'sub-category' || val === 'subcat') map.subCat = colNumber;
            else if (val === 'category' || val === 'cat') map.category = colNumber;
            else if (val === 'vendor' || val === 'vend') map.vendor = colNumber;
            else if (val === 'customer' || val === 'cust') map.customer = colNumber;
        });

        if (showChecker) {
            console.log(`\nProcessing "${sheet.name}":`);
            console.log(`  Header Row: ${headerRowIndex}`);
            console.log(`  Mapping: ${JSON.stringify(map)}`);
        }

        sheet.eachRow((row, r) => {
            if (r <= config.offset) return;

            const vendorVal = getVal(row.getCell(map.vendor));
            const customerVal = getVal(row.getCell(map.customer));
            const categoryVal = getVal(row.getCell(map.category));
            const subCatVal = getVal(row.getCell(map.subCat));
            let amount = getVal(row.getCell(map.amount));
            if (typeof amount !== 'number') amount = parseFloat(amount) || 0;

            const rawDate = getVal(row.getCell(map.date));
            const rawDesc = getVal(row.getCell(map.desc)).toString().toLowerCase();

            // Offset check
            if (r === config.offset + 1) {
                const rowValues = row.values.map(v => (v ? v.toString().toLowerCase() : ''));
                const matches = ['date', 'amount', 'category', 'description'].filter(t => rowValues.some(rv => rv.includes(t)));
                if (matches.length >= 3) {
                    offsetWarnings.push({ sheet: sheet.name, row: r, matches });
                }
            }

            if (!rawDate && !rawDesc && !amount) return;
            if (rawDesc.includes('total') || rawDesc.includes('balance') || rawDesc.includes('sum')) return;
            if (!rawDate) return;

            if (config.flip) amount *= -1;
            if (pType === 'cc') ccTotal += amount; else bankTotal += amount;

            const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');

            if (!categoryVal && Math.abs(amount) > 0.01) {
                if (pType === 'cc') uncategorizedCC++; else uncategorizedBank++;
                uncategorizedDetails.push({ sheet: sheet.name, row: r, date: displayDate, desc: rawDesc });
            } else if (categoryVal) {
                const catStr = categoryVal.toString().trim();
                if (!validCategories.has(catStr)) {
                    illegalCategories.push({ value: catStr, sheet: sheet.name, row: r, date: displayDate });
                }
                if (!catStats[catStr]) catStats[catStr] = { total: 0, subCats: {} };
                catStats[catStr].total += amount;
                const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';
                catStats[catStr].subCats[sName] = (catStats[catStr].subCats[sName] || 0) + amount;
            }

            if (vendorVal) {
                const vStr = vendorVal.toString().trim();
                if (!validVendors.has(vStr)) illegalVendors.push({ value: vStr, sheet: sheet.name, row: r, date: displayDate });
                vendorStats[vStr] = (vendorStats[vStr] || 0) + amount;
            }
            if (customerVal) {
                const cStr = customerVal.toString().trim();
                if (!validCustomers.has(cStr)) illegalCustomers.push({ value: cStr, sheet: sheet.name, row: r, date: displayDate });
                customerStats[cStr] = (customerStats[cStr] || 0) + amount;
            }
        });
    }

    // --- 3. Process Ledger ---
    ledgerSheet.eachRow((row, r) => {
        if (r === 1) return;
        const rawDate = getVal(row.getCell(1));
        const rawDesc = getVal(row.getCell(2));
        const cat = getVal(row.getCell(3));
        const dr = parseFloat(getVal(row.getCell(4))) || 0;
        const cr = parseFloat(getVal(row.getCell(5))) || 0;

        // Skip truly empty rows or rows without dates (user requirement)
        if (!rawDate && !cat && !rawDesc && !dr && !cr) return;

        if (!rawDate) {
            if (cat || dr || cr) {
                if (showChecker) console.log(`Ledger Row ${r}: SKIPPED (Missing Date). Rows must have dates.`);
            }
            return;
        }

        if (cat) {
            const catStr = cat.toString().trim();
            const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');

            if (!validCategories.has(catStr)) illegalCategories.push({ value: catStr, sheet: 'Ledger', row: r, date: displayDate });

            const conf = uniqueCategories.get(catStr);
            const impact = (conf && conf.report === 'P&L') ? (cr - dr) : (dr - cr);
            if (!catStats[catStr]) catStats[catStr] = { total: 0, subCats: {} };
            catStats[catStr].total += impact;

            // Integration: Update calculated bank/cc balances from Ledger entries
            const targetAccountType = conf ? (conf.accountType || '').toLowerCase() : '';
            const matchingConfig = sheetConfigs.find(s => {
                const sName = s.name.toLowerCase();
                const sType = s.type.toLowerCase();
                const cName = catStr.toLowerCase();

                return sName === cName ||
                    sType === cName ||
                    (targetAccountType && sType === targetAccountType) ||
                    (targetAccountType && sName === targetAccountType);
            });

            if (matchingConfig) {
                const isMatchingCC = matchingConfig.type.toLowerCase().includes('cc') || matchingConfig.type.toLowerCase().includes('credit');
                // Debiting an asset increase balance. Crediting an asset decreases balance.
                if (isMatchingCC) ccTotal += (dr - cr); else bankTotal += (dr - cr);
                if (showChecker) console.log(`Ledger Row ${r} [${displayDate}]: Applied ${dr - cr} impact to "${matchingConfig.name}" balance.`);
            }
        }
    });

    // --- 4. Prepare Reports ---
    const reports = { pl: [], bs: [] };
    const pnlNames = Array.from(uniqueCategories.keys()).filter(n => uniqueCategories.get(n).report === 'P&L').sort();
    reports.pl = pnlNames.map(n => ({ label: n, value: catStats[n] ? catStats[n].total : 0 }));
    const netIncome = reports.pl.reduce((a, b) => a + b.value, 0);

    reports.bs = [{ label: '** Bank Balance (Calculated)', value: bankTotal }, { label: '** CC Balance (Calculated)', value: ccTotal }];

    // --- 5. Console Output ---
    function printSection(title, rows) {
        console.log(`\n--- ${title} ---`);
        if (!rows.length) { console.log('(No Data)'); return; }
        const max = Math.max(...rows.map(r => r.label.length), 10);
        rows.forEach(r => console.log(`${r.label.padEnd(max + 5)} : ${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`));
    }

    if (showAll || showPL) { printSection('PROFIT & LOSS', reports.pl); console.log(`\n=== NET INCOME: ${netIncome.toFixed(2)} ===\n`); }
    if (showAll || showBS) printSection('BALANCE SHEET', reports.bs);

    const hasIssues = uncategorizedDetails.length > 0 || illegalCategories.length > 0 || illegalVendors.length > 0 || illegalCustomers.length > 0;
    if (hasIssues) {
        console.log('\n--- DATA INTEGRITY ISSUES ---');
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet)
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

            if (showChecker) {
                uncat.forEach(x => console.log(`      - [${x.date}] Row ${x.row}: MISSING CATEGORY ("${x.desc}")`));
                illegalCategories.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL CATEGORY "${x.value}"`));
                illegalVendors.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN VENDOR "${x.value}"`));
                illegalCustomers.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN CUSTOMER "${x.value}"`));
            }
        });
    }

    if (offsetWarnings.length) {
        console.log('\n--- OFFSET WARNINGS ---');
        offsetWarnings.forEach(w => console.log(`[!] Sheet "${w.sheet}" Row ${w.row} looks like a header (Found: ${w.matches.join(', ')}). Adjust Setup tab offset.`));
    }

    if (printOnly) return;

    // --- 6. Summary Sheet Update ---
    if (summarySheet) workbook.removeWorksheet(summarySheet.id);
    summarySheet = workbook.addWorksheet('Summary');
    summarySheet.getCell('A1').value = `Financial Summary (${new Date().toLocaleString()})`;
    summarySheet.getCell('A1').font = { size: 14, bold: true };

    let summaryRow = 3;
    summarySheet.getCell(`A${summaryRow}`).value = 'Profit & Loss';
    summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
    reports.pl.forEach(r => { summarySheet.getCell(`A${summaryRow}`).value = r.label; summarySheet.getCell(`B${summaryRow}`).value = r.value; summaryRow++; });
    summarySheet.getCell(`A${summaryRow}`).value = 'NET INCOME'; summarySheet.getCell(`B${summaryRow}`).value = netIncome;
    summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow += 3;

    summarySheet.getCell(`A${summaryRow}`).value = 'Balance Sheet';
    summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
    reports.bs.forEach(r => { summarySheet.getCell(`A${summaryRow}`).value = r.label; summarySheet.getCell(`B${summaryRow}`).value = r.value; summaryRow++; });

    if (hasIssues) {
        summaryRow += 3;
        summarySheet.getCell(`A${summaryRow}`).value = 'Data Integrity Check';
        summarySheet.getCell(`A${summaryRow}`).font = { bold: true, color: { argb: 'FFFF0000' } }; summaryRow++;
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet)
        ]);
        issueSheetsFound.forEach(s => {
            summarySheet.getCell(`A${summaryRow}`).value = `Tab: ${s.toUpperCase()}`;
            summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
            const uncat = uncategorizedDetails.filter(x => x.sheet === s).length;
            if (uncat) { summarySheet.getCell(`A${summaryRow}`).value = '  Uncategorized Rows'; summarySheet.getCell(`B${summaryRow}`).value = uncat; summaryRow++; }
            const cats = Array.from(new Set(illegalCategories.filter(x => x.sheet === s).map(x => x.value))).join(', ');
            if (cats) { summarySheet.getCell(`A${summaryRow}`).value = '  Illegal Categories'; summarySheet.getCell(`B${summaryRow}`).value = cats; summaryRow++; }
            const vends = Array.from(new Set(illegalVendors.filter(x => x.sheet === s).map(x => x.value))).join(', ');
            if (vends) { summarySheet.getCell(`A${summaryRow}`).value = '  Unknown Vendors'; summarySheet.getCell(`B${summaryRow}`).value = vends; summaryRow++; }
            summaryRow++;
        });
    }

    try {
        await workbook.xlsx.writeFile(filename);
        if (showChecker) console.log(`\nSuccessfully updated financials in ${filename}`);
    } catch (e) {
        console.error('Error saving file:', e.message);
    }
}

updateFinancials().catch(console.error);

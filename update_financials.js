const ExcelJS = require('exceljs');
const fs = require('fs');

async function updateFinancials() {
    // --- Argument Parsing ---
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
    if (!filename.endsWith('.xlsx')) filename += '.xlsx';

    if (!fs.existsSync(filename)) {
        console.error(`Error: File '${filename}' not found.`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(filename);
    } catch (error) {
        if (error.code === 'EBUSY') {
            console.error(`\nERROR: The file '${filename}' is currently OPEN in Excel.`);
            return;
        }
        throw error;
    }

    const setupSheet = workbook.getWorksheet('Setup');
    const ledgerSheet = workbook.getWorksheet('Ledger');
    let summarySheet = workbook.getWorksheet('Summary');

    if (!setupSheet || !ledgerSheet) {
        console.error('Error: Required sheets (Setup, Ledger) missing.');
        return;
    }

    // --- 1. Read Setup Data & Configuration ---
    const validCategories = new Set();
    const validVendors = new Set();
    const validCustomers = new Set();
    const uniqueCategories = new Map();
    const sheetConfigs = [];

    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const catName = row.getCell(1).value;
        const report = row.getCell(4).value;
        if (catName) {
            const trimmed = catName.toString().trim();
            validCategories.add(trimmed);
            uniqueCategories.set(trimmed, { report });
        }

        const vendor = row.getCell(6).value;
        if (vendor) validVendors.add(vendor.toString().trim());

        const customer = row.getCell(7).value;
        if (customer) validCustomers.add(customer.toString().trim());

        const confSheetName = row.getCell(9).value;
        const confType = row.getCell(10).value;
        const confFlip = row.getCell(11).value;
        const confOffset = row.getCell(12).value;

        if (confSheetName && confType) {
            sheetConfigs.push({
                name: confSheetName.toString().trim(),
                type: confType.toString().trim(),
                flip: !!(confFlip && confFlip.toString().toLowerCase().includes('y')),
                offset: parseInt(confOffset) || 1
            });
        }
    });

    if (sheetConfigs.length === 0) {
        sheetConfigs.push({ name: 'Bank Transactions', type: 'Bank', flip: false, offset: 1 });
        sheetConfigs.push({ name: 'Credit Card Transactions', type: 'CC', flip: false, offset: 1 });
    }

    // --- 2. Aggregate Transactions ---
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

    const bankMap = { date: 1, desc: 2, amount: 3, category: 4, subCat: 5, vendor: 7, customer: 8 };
    const ccMap = { date: 1, desc: 3, amount: 4, category: 5, subCat: 6, vendor: 8, customer: 9 };

    function processLine(row, processingType, flip) {
        const isCC = (processingType === 'cc');
        const map = isCC ? ccMap : bankMap;

        const vendor = row.getCell(map.vendor).value;
        const customer = row.getCell(map.customer).value;
        const category = row.getCell(map.category).value;
        const subCat = row.getCell(map.subCat).value;
        let amount = row.getCell(map.amount).value || 0;

        if (amount && typeof amount === 'object' && amount.result !== undefined) amount = amount.result;
        if (typeof amount !== 'number') amount = parseFloat(amount) || 0;

        const rawDate = row.getCell(map.date).value;
        const rawDesc = row.getCell(map.desc).value ? row.getCell(map.desc).value.toString().toLowerCase() : '';

        if (!rawDate && !rawDesc && !amount) return;
        if (rawDesc.includes('total') || rawDesc.includes('balance') || rawDesc.includes('sum')) return;
        if (!rawDate) return;

        if (flip) amount = amount * -1;

        if (isCC) ccTotal += amount;
        else bankTotal += amount;

        // --- Integrity Checks ---
        const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');

        if (!category && Math.abs(amount) > 0.01) {
            if (isCC) uncategorizedCC++; else uncategorizedBank++;
            uncategorizedDetails.push({ sheet: processingType, row: row.number, date: displayDate, desc: rawDesc });
        } else if (category) {
            const catStr = category.toString().trim();
            if (!validCategories.has(catStr)) {
                illegalCategories.push({ value: catStr, sheet: processingType, row: row.number, date: displayDate });
            }

            if (!catStats[catStr]) catStats[catStr] = { total: 0, subCats: {} };
            catStats[catStr].total += amount;

            const subName = subCat ? subCat.toString().trim() : '(No Sub-Cat)';
            catStats[catStr].subCats[subName] = (catStats[catStr].subCats[subName] || 0) + amount;
        }

        if (vendor) {
            const vStr = vendor.toString().trim();
            if (!validVendors.has(vStr)) illegalVendors.push({ value: vStr, sheet: processingType, row: row.number, date: displayDate });
            vendorStats[vStr] = (vendorStats[vStr] || 0) + amount;
        }

        if (customer) {
            const cStr = customer.toString().trim();
            if (!validCustomers.has(cStr)) illegalCustomers.push({ value: cStr, sheet: processingType, row: row.number, date: displayDate });
            customerStats[cStr] = (customerStats[cStr] || 0) + amount;
        }
    }

    for (const config of sheetConfigs) {
        const sheet = workbook.getWorksheet(config.name);
        if (sheet) {
            let pType = 'bank';
            const tStr = config.type.toLowerCase();
            const isCC = tStr.includes('cc') || tStr.includes('card') || tStr.includes('credit') || tStr.includes('amex');
            if (isCC) pType = 'cc';

            sheet.eachRow((row, r) => { if (r > config.offset) processLine(row, pType, config.flip); });
        }
    }

    // --- 3. Process Ledger ---
    ledgerSheet.eachRow((row, r) => {
        if (r === 1) return;
        const rawDate = row.getCell(1).value;
        const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');
        const cat = row.getCell(3).value;
        const dr = row.getCell(4).value || 0;
        const cr = row.getCell(5).value || 0;
        if (cat) {
            const catStr = cat.toString().trim();
            if (!validCategories.has(catStr)) illegalCategories.push({ value: catStr, sheet: 'Ledger', row: r, date: displayDate });

            const config = uniqueCategories.get(catStr);
            const impact = (config && config.report === 'P&L') ? (cr - dr) : (dr - cr);
            if (!catStats[catStr]) catStats[catStr] = { total: 0, subCats: {} };
            catStats[catStr].total += impact;
        }
    });

    // --- 4. Reports ---
    const reports = {};
    const pnlNames = Array.from(uniqueCategories.keys()).filter(n => uniqueCategories.get(n).report === 'P&L').sort();
    reports.pl = pnlNames.map(n => ({ label: n, value: catStats[n] ? catStats[n].total : 0 }));
    let netIncome = reports.pl.reduce((a, b) => a + b.value, 0);

    const bsNames = Array.from(uniqueCategories.keys()).filter(n => uniqueCategories.get(n).report === 'Balance Sheet').sort();
    reports.bs = [{ label: '** Bank Balance (Calculated)', value: bankTotal }, { label: '** CC Balance (Calculated)', value: ccTotal }];

    const exclusions = [...sheetConfigs.map(s => s.name.toLowerCase()), 'checking account', 'savings account', 'credit card'];
    bsNames.forEach(n => {
        if (exclusions.some(ex => n.toLowerCase().includes(ex))) return;
        const amt = catStats[n] ? catStats[n].total : 0;
        if (Math.abs(amt) > 0.01) reports.bs.push({ label: n, value: amt });
    });

    function printSection(title, rows) {
        console.log(`\n--- ${title} ---`);
        if (!rows.length) { console.log('(No Data)'); return; }
        const max = Math.max(...rows.map(r => r.label.length), 10);
        rows.forEach(r => console.log(`${r.label.padEnd(max + 5)} : ${r.value.toFixed(2).padStart(10)}`));
    }

    if (showAll || showPL) { printSection('PROFIT & LOSS', reports.pl); console.log(`\n=== NET INCOME: ${netIncome.toFixed(2)} ===\n`); }
    if (showAll || showBS) printSection('BALANCE SHEET', reports.bs);

    // --- Print Integrity Issues to Console ---
    const hasIssues = uncategorizedBank > 0 || uncategorizedCC > 0 || illegalCategories.length > 0 || illegalVendors.length > 0 || illegalCustomers.length > 0;
    if (hasIssues) {
        console.log('\n--- DATA INTEGRITY ISSUES ---');

        const allIssueSheets = new Set([
            ...uncategorizedDetails.map(d => d.sheet),
            ...illegalCategories.map(c => c.sheet),
            ...illegalVendors.map(v => v.sheet),
            ...illegalCustomers.map(c => c.sheet)
        ]);

        allIssueSheets.forEach(sheetName => {
            const displaySheet = sheetName.toUpperCase();
            console.log(`\n>> Tab: ${displaySheet}`);

            const sheetUncat = uncategorizedDetails.filter(d => d.sheet === sheetName);
            if (sheetUncat.length > 0) console.log(`  [!] ${sheetUncat.length} rows missing category`);

            const sheetIllegalCats = illegalCategories.filter(d => d.sheet === sheetName);
            if (sheetIllegalCats.length > 0) {
                const uniqueCats = new Set(sheetIllegalCats.map(c => c.value));
                console.log(`  [!] Illegal Categories: ${Array.from(uniqueCats).join(', ')}`);
            }

            const sheetIllegalVendors = illegalVendors.filter(d => d.sheet === sheetName);
            if (sheetIllegalVendors.length > 0) {
                const uniqueVends = new Set(sheetIllegalVendors.map(v => v.value));
                console.log(`  [!] Unknown Vendors: ${Array.from(uniqueVends).join(', ')}`);
            }

            const sheetIllegalCustomers = illegalCustomers.filter(d => d.sheet === sheetName);
            if (sheetIllegalCustomers.length > 0) {
                const uniqueCusts = new Set(sheetIllegalCustomers.map(c => c.value));
                console.log(`  [!] Unknown Customers: ${Array.from(uniqueCusts).join(', ')}`);
            }

            if (showChecker) {
                sheetUncat.forEach(d => console.log(`      - [${d.date}] Row ${d.row}: MISSING CATEGORY ("${d.desc}")`));
                sheetIllegalCats.forEach(d => console.log(`      - [${d.date}] Row ${d.row}: ILLEGAL CATEGORY "${d.value}"`));
                sheetIllegalVendors.forEach(d => console.log(`      - [${d.date}] Row ${d.row}: UNKNOWN VENDOR "${d.value}"`));
                sheetIllegalCustomers.forEach(d => console.log(`      - [${d.date}] Row ${d.row}: UNKNOWN CUSTOMER "${d.value}"`));
            }
        });
    }

    if (printOnly) return;

    if (summarySheet) workbook.removeWorksheet(summarySheet.id);
    summarySheet = workbook.addWorksheet('Summary');
    summarySheet.getCell('A1').value = 'Financial Summary (' + new Date().toLocaleString() + ')';
    summarySheet.getCell('A1').font = { size: 14, bold: true };

    let row = 3;
    summarySheet.getCell(`A${row}`).value = 'Profit & Loss';
    summarySheet.getCell(`A${row}`).font = { bold: true }; row++;
    reports.pl.forEach(r => { summarySheet.getCell(`A${row}`).value = r.label; summarySheet.getCell(`B${row}`).value = r.value; row++; });
    summarySheet.getCell(`A${row}`).value = 'NET INCOME'; summarySheet.getCell(`B${row}`).value = netIncome;
    summarySheet.getCell(`A${row}`).font = { bold: true }; row += 3;

    summarySheet.getCell(`A${row}`).value = 'Balance Sheet';
    summarySheet.getCell(`A${row}`).font = { bold: true }; row++;
    reports.bs.forEach(r => { summarySheet.getCell(`A${row}`).value = r.label; summarySheet.getCell(`B${row}`).value = r.value; row++; });

    // --- Integrity Report in Excel (Per Tab) ---
    if (hasIssues) {
        row += 3;
        summarySheet.getCell(`A${row}`).value = 'Data Integrity Check (By Tab)';
        summarySheet.getCell(`A${row}`).font = { bold: true, color: { argb: 'FFFF0000' } }; row++;

        const allIssueSheets = new Set([
            ...uncategorizedDetails.map(d => d.sheet),
            ...illegalCategories.map(c => c.sheet),
            ...illegalVendors.map(v => v.sheet),
            ...illegalCustomers.map(c => c.sheet)
        ]);

        allIssueSheets.forEach(sheetName => {
            summarySheet.getCell(`A${row}`).value = `Tab: ${sheetName.toUpperCase()}`;
            summarySheet.getCell(`A${row}`).font = { bold: true };
            row++;

            const uncat = uncategorizedDetails.filter(d => d.sheet === sheetName).length;
            if (uncat > 0) { summarySheet.getCell(`A${row}`).value = '  Uncategorized Rows'; summarySheet.getCell(`B${row}`).value = uncat; row++; }

            const cats = Array.from(new Set(illegalCategories.filter(d => d.sheet === sheetName).map(c => c.value))).join(', ');
            if (cats) { summarySheet.getCell(`A${row}`).value = '  Illegal Categories'; summarySheet.getCell(`B${row}`).value = cats; row++; }

            const vends = Array.from(new Set(illegalVendors.filter(d => d.sheet === sheetName).map(v => v.value))).join(', ');
            if (vends) { summarySheet.getCell(`A${row}`).value = '  Unknown Vendors'; summarySheet.getCell(`B${row}`).value = vends; row++; }

            const custs = Array.from(new Set(illegalCustomers.filter(d => d.sheet === sheetName).map(c => c.value))).join(', ');
            if (custs) { summarySheet.getCell(`A${row}`).value = '  Unknown Customers'; summarySheet.getCell(`B${row}`).value = custs; row++; }

            row++; // spacer
        });
    }

    try {
        await workbook.xlsx.writeFile(filename);
        console.log(`\nSuccessfully updated financials in ${filename}`);
    } catch (e) {
        console.error('Error saving file:', e.message);
    }
}
updateFinancials();

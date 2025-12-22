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

    const specificFilter = showPL || showBS || showVendor || showCustomer || showPLSub;
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
    const uniqueCategories = new Map();
    const sheetConfigs = [];

    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const catName = row.getCell(1).value;
        const report = row.getCell(4).value;
        if (catName) uniqueCategories.set(catName.toString().trim(), { report });

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

        // --- Smart Skip (Prevents Double Counting Totals) ---
        // 1. Skip strictly empty rows
        if (!rawDate && !rawDesc && !amount) return;
        // 2. Skip rows that look like manual Totals/Balances
        if (rawDesc.includes('total') || rawDesc.includes('balance') || rawDesc.includes('sum')) return;
        // 3. Skip rows without a date (Transactions should always have a date)
        if (!rawDate) return;

        if (flip) amount = amount * -1;

        if (isCC) ccTotal += amount;
        else bankTotal += amount;

        if (!category && Math.abs(amount) > 0.01) {
            if (isCC) uncategorizedCC++; else uncategorizedBank++;
        }

        if (category) {
            const catStr = category.toString().trim();
            if (!catStats[catStr]) catStats[catStr] = { total: 0, subCats: {} };
            catStats[catStr].total += amount;

            const subName = subCat ? subCat.toString().trim() : '(No Sub-Cat)';
            catStats[catStr].subCats[subName] = (catStats[catStr].subCats[subName] || 0) + amount;
        }

        if (vendor) {
            const vStr = vendor.toString().trim();
            vendorStats[vStr] = (vendorStats[vStr] || 0) + amount;
        }

        if (customer) {
            const cStr = customer.toString().trim();
            customerStats[cStr] = (customerStats[cStr] || 0) + amount;
        }
    }

    for (const config of sheetConfigs) {
        const sheet = workbook.getWorksheet(config.name);
        if (sheet) {
            let pType = 'bank';
            const tStr = config.type.toLowerCase();
            const tokens = tStr.split(/[^a-z0-9]+/);

            const isCC = tokens.includes('cc') || tokens.includes('card') || tokens.includes('credit') || tokens.includes('amex') || tokens.includes('visa') || tokens.includes('liability');
            const isBank = tokens.includes('checking') || tokens.includes('savings') || tokens.includes('debit');

            if (isBank) pType = 'bank';
            else if (isCC) pType = 'cc';
            else if (tStr.includes('bank')) pType = 'bank';

            sheet.eachRow((row, r) => { if (r > config.offset) processLine(row, pType, config.flip); });
        }
    }

    // --- 3. Process Ledger ---
    ledgerSheet.eachRow((row, r) => {
        if (r === 1) return;
        const cat = row.getCell(3).value;
        const dr = row.getCell(4).value || 0;
        const cr = row.getCell(5).value || 0;
        if (cat) {
            const catStr = cat.toString().trim();
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

    const exclusions = [...sheetConfigs.map(s => s.name.toLowerCase()), ...sheetConfigs.map(s => s.type.toLowerCase()), 'checking account', 'savings account', 'credit card'];
    bsNames.forEach(n => {
        if (exclusions.includes(n.toLowerCase())) return;
        const amt = catStats[n] ? catStats[n].total : 0;
        if (Math.abs(amt) > 0.01) reports.bs.push({ label: n, value: amt });
    });

    reports.vendor = Object.entries(vendorStats).sort((a, b) => a[1] - b[1]).slice(0, 10).map(([k, v]) => ({ label: k, value: v }));
    reports.customer = Object.entries(customerStats).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([k, v]) => ({ label: k, value: v }));

    reports.plSub = [];
    pnlNames.forEach(name => {
        const total = catStats[name] ? catStats[name].total : 0;
        reports.plSub.push({ type: 'main', category: name, value: total });
        if (catStats[name] && catStats[name].subCats) {
            const subs = catStats[name].subCats;
            const valid = Object.keys(subs).filter(s => s !== '(No Sub-Cat)' && Math.abs(subs[s]) > 0.01).sort();
            if (valid.length > 1) valid.forEach(s => reports.plSub.push({ type: 'sub', subCategory: s, value: subs[s] }));
        }
    });

    function printSection(title, rows) {
        console.log(`\n--- ${title} ---`);
        if (!rows.length) { console.log('(No Data)'); return; }
        const max = Math.max(...rows.map(r => r.label.length), 10);
        rows.forEach(r => console.log(`${r.label.padEnd(max + 5)} : ${r.value.toFixed(2).padStart(10)}`));
    }

    if (showAll || showPL) { printSection('PROFIT & LOSS', reports.pl); console.log(`\n=== NET INCOME: ${netIncome.toFixed(2)} ===\n`); }
    if (showAll || showBS) printSection('BALANCE SHEET', reports.bs);
    if (showAll || showVendor) printSection('TOP VENDORS', reports.vendor);
    if (showAll || showCustomer) printSection('TOP CUSTOMERS', reports.customer);

    if (printOnly) return;

    if (summarySheet) workbook.removeWorksheet(summarySheet.id);
    summarySheet = workbook.addWorksheet('Summary');
    summarySheet.getCell('A1').value = 'Financial Summary (' + new Date().toLocaleString() + ')';
    summarySheet.getCell('A1').font = { size: 14, bold: true };
    let row = 3;
    summarySheet.getCell(`A${row}`).value = 'Profit & Loss'; row++;
    reports.pl.forEach(r => { summarySheet.getCell(`A${row}`).value = r.label; summarySheet.getCell(`B${row}`).value = r.value; row++; });
    summarySheet.getCell(`A${row}`).value = 'NET INCOME'; summarySheet.getCell(`B${row}`).value = netIncome; row += 3;
    summarySheet.getCell(`A${row}`).value = 'Balance Sheet'; row++;
    reports.bs.forEach(r => { summarySheet.getCell(`A${row}`).value = r.label; summarySheet.getCell(`B${row}`).value = r.value; row++; });

    try {
        await workbook.xlsx.writeFile(filename);
        console.log(`\nSuccessfully updated financials in ${filename}`);
    } catch (e) {
        console.error('Error saving file:', e.message);
    }
}
updateFinancials();

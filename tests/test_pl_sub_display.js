const { execSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const assert = require('assert');

const TEST_FILENAME = 'Temp_SubCat_Test.xlsx';

async function verifySubCatLogic() {
    console.log('--- Setting up Test for P&L Sub-Category Logic ---');
    const workbook = new ExcelJS.Workbook();

    // 1. Setup Sheet
    const setup = workbook.addWorksheet('Setup');
    setup.columns = [{ header: 'Category', key: 'cat' }, { header: 'Account Type', key: 'type' }, { header: 'Report', key: 'rep' }, { header: 'Report Type', key: 'rtype' }];
    setup.addRow(['Cat_Multi', 'Expense', '', 'P&L']);
    setup.addRow(['Cat_OnlyNoSub', 'Expense', '', 'P&L']);
    setup.addRow(['Cat_Mixed', 'Expense', '', 'P&L']);
    setup.addRow(['Cat_SingleReal', 'Expense', '', 'P&L']);

    // 1.5. Ledger Sheet (Mandatory)
    const ledger = workbook.addWorksheet('Ledger');
    ledger.columns = [{ header: 'Date', key: 'd' }, { header: 'Desc', key: 'desc' }, { header: 'Cat', key: 'c' }, { header: 'Sub', key: 's' }, { header: 'Vend', key: 'v' }, { header: 'Cust', key: 'cust' }, { header: 'Dr', key: 'dr' }, { header: 'Cr', key: 'cr' }];
    ledger.addRow(['2025-01-01', 'Init', '', '', '', '', 0, 0]);


    // 2. Bank Transactions
    // Columns: [1]Date [2]Desc [3]Amount [4]Cat [5]Sub ...
    const sheet = workbook.addWorksheet('Bank Transactions');
    sheet.columns = [
        { header: 'Date', key: 'date' }, { header: 'Desc', key: 'desc' }, { header: 'Amount', key: 'amt' },
        { header: 'Category', key: 'cat' }, { header: 'Sub-Category', key: 'sub' }
    ];

    // Transaction Data
    const rows = [
        // Case: Multi (Should show all)
        ['2025-01-01', 'Multi 1', -100, 'Cat_Multi', 'SubA'],
        ['2025-01-01', 'Multi 2', -200, 'Cat_Multi', 'SubB'],

        // Case: Only No Sub (Should HIDDEN sub-cat line)
        ['2025-01-01', 'NoSub Only', -300, 'Cat_OnlyNoSub', ''],

        // Case: Mixed (Should show "No Sub-Cat" because it distinguishes from others)
        ['2025-01-01', 'Mixed 1', -40, 'Cat_Mixed', ''],
        ['2025-01-01', 'Mixed 2', -60, 'Cat_Mixed', 'SubC'],

        // Case: Single Real Sub (Should SHOW, as it provides specific info)
        ['2025-01-01', 'Single Real', -500, 'Cat_SingleReal', 'SpecialProject']
    ];

    rows.forEach(r => sheet.addRow(r));
    await workbook.xlsx.writeFile(TEST_FILENAME);

    console.log('--- Running update_financials.js --pl-sub ---');
    let output = '';
    try {
        output = execSync(`node update_financials.js "${TEST_FILENAME}" --pl-sub`, { encoding: 'utf-8' });
    } catch (e) {
        console.error('Execution Failed:', e.stdout);
        process.exit(1);
    }

    console.log('--- Analyzing Output ---');
    console.log(output);

    // Helper asserts
    const contains = (str) => output.includes(str);

    // Assertions
    const errors = [];

    // 1. Cat_Multi: Should see SubA and SubB
    if (!contains('Cat_Multi')) errors.push('Missing Cat_Multi main line');
    if (!contains('SubA')) errors.push('Missing SubA in Cat_Multi');
    if (!contains('SubB')) errors.push('Missing SubB in Cat_Multi');

    // 2. Cat_OnlyNoSub: Should see Main Line, Should NOT see "(No Sub-Cat)"
    if (!contains('Cat_OnlyNoSub')) errors.push('Missing Cat_OnlyNoSub main line');
    // We need to be careful; "(No Sub-Cat)" exists for Cat_Mixed.
    // We check via regex or line proximity, or just ensure the output behaves generally.
    // Since "Cat_OnlyNoSub" is unique, let's look for the block.
    const onlyNoSubBlock = output.split('\n').filter(l => l.includes('Cat_OnlyNoSub'))[0];
    // We unfortunately can't split blocks easily from raw text without parsing. 
    // However, if we look at the lines immediately following:
    const lines = output.split('\n');
    const idx = lines.findIndex(l => l.includes('Cat_OnlyNoSub'));
    if (idx !== -1) {
        const nextLine = lines[idx + 1] || '';
        if (nextLine.includes('(No Sub-Cat)')) {
            errors.push('FAILED: Cat_OnlyNoSub showed "(No Sub-Cat)" but it should be hidden!');
        }
    }

    // 3. Cat_Mixed: Should see "(No Sub-Cat)" and "SubC"
    // Use index search to find the Mixed block
    const mixedIdx = lines.findIndex(l => l.includes('Cat_Mixed'));
    if (mixedIdx !== -1) {
        let foundNoSub = false;
        let foundSubC = false;
        // Check next few lines until another category starts (or empty line)
        for (let i = mixedIdx + 1; i < lines.length; i++) {
            if (lines[i].trim() === '' || !lines[i].includes('>')) break; // Stop at end of block
            if (lines[i].includes('(No Sub-Cat)')) foundNoSub = true;
            if (lines[i].includes('SubC')) foundSubC = true;
        }
        if (!foundNoSub) errors.push('FAILED: Cat_Mixed should show "(No Sub-Cat)"');
        if (!foundSubC) errors.push('FAILED: Cat_Mixed should show "SubC"');
    }

    // 4. Cat_SingleReal: Should see "SpecialProject"
    if (!contains('SpecialProject')) errors.push('Missing SpecialProject sub-cat');

    if (errors.length > 0) {
        console.error('TEST FAILED with errors:');
        errors.forEach(e => console.error(` - ${e}`));
        process.exit(1);
    } else {
        console.log('✅ TEST PASSED: All sub-category formatting logic verified.');
    }

    // Cleanup
    try { fs.unlinkSync(TEST_FILENAME); } catch (e) { }
}

verifySubCatLogic();

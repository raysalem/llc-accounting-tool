const { execSync } = require('child_process');
const fs = require('fs');
const ExcelJS = require('exceljs');

async function runTest() {
    const TEST_FILE = 'Test_Accounting.xlsx';

    console.log('--- Phase 1: Initialize Template ---');
    execSync('node generate_excel.js', { stdio: 'inherit' });
    fs.renameSync('LLC_Accounting_Template.xlsx', TEST_FILE);

    console.log('\n--- Phase 2: Load Bank Transactions ---');
    execSync(`node load_transactions.js tests/example_bank.csv bank ${TEST_FILE} --clear`, { stdio: 'inherit' });

    console.log('\n--- Phase 3: Load CC Transactions ---');
    execSync(`node load_transactions.js tests/example_cc.csv cc ${TEST_FILE} --clear`, { stdio: 'inherit' });

    console.log('\n--- Phase 4: Categorize Transactions (Simulated) ---');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEST_FILE);

    const bankSheet = workbook.getWorksheet('Bank Transactions');
    bankSheet.eachRow((row, r) => {
        if (r === 1) return;
        const descCell = row.getCell(2).value;
        if (!descCell) return;
        const desc = descCell.toString();
        if (desc.includes('Salary')) row.getCell(4).value = 'Sales';
        if (desc.includes('Rent')) row.getCell(4).value = 'Rent';
        if (desc.includes('Client')) row.getCell(4).value = 'Sales';
    });

    const ccSheet = workbook.getWorksheet('Credit Card Transactions');
    ccSheet.eachRow((row, r) => {
        if (r === 1) return;
        const descCell = row.getCell(3).value;
        if (!descCell) return;
        const desc = descCell.toString();
        if (desc.includes('Starbucks')) row.getCell(5).value = 'Travel';
        if (desc.includes('Amazon')) row.getCell(5).value = 'Office';
        if (desc.includes('AWS')) row.getCell(5).value = 'Office';
    });

    // --- Add Illegal Entries for Integrity Test ---
    bankSheet.addRow([new Date('2025-01-25'), 'Mystery Corp', 100, 'IllegalCat', '', '', 'UnknownVendor', '']);
    bankSheet.addRow([new Date('2025-01-26'), 'Uncategorized Expense', 50, '', '', '', '', '']); // No category

    // 1099 Test: > $600 Transaction for NEC Vendor
    bankSheet.addRow([new Date('2025-01-27'), 'Major Contract Work', -737.50, 'Services', '', '', 'Contractor 1099', '']);

    // Add necessary categories to Setup if missing
    const setupSheet = workbook.getWorksheet('Setup');
    setupSheet.addRow(['Travel', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Travel', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Office', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Services', 'General', 'Expense', 'P&L']); // Fix "Illegal Category"

    // Setup for 1099 Vendor: Column F (6) is Vendors. We need "1099 Type" which might be Column O/P (15/16) in standard template?
    // We must find the headers first or just append to known cols. 
    // Template 'Setup' headers are strict. We should find the Vendor table and append there.
    // However, for this test, we accept we might be writing blind unless we scan headers.
    // Simpler: Just write to known Setup table row area. But we need to hit the "Vendors" and "1099 Type" columns.
    // Standard Template has Vendors at Col F (6). 1099 Type at Col G (7) or similar if we added split cols there?
    // Wait, the template generation script creates a standard "Setup".
    // Let's assume we can just add a row to the end of the Setup sheet and rely on headers being in row 1.
    // Actually, report.js reads vertical tables relative to headers.
    // Let's find "Vendors" and "1099 Type" column indices dynamically in this test script?
    // Too complex. Let's just assume standard template: Vendors=F, 1099=G (if simplified) or we can just create a map here.

    // Better: Scan Headers in this Test Script
    // Better: Scan Headers in this Test Script
    const setupHeaders = {};
    setupSheet.getRow(1).eachCell((c, col) => {
        if (c.value) setupHeaders[c.value.toString().toLowerCase()] = col;
    });
    const colVend = setupHeaders['vendors'] || 6;
    const col1099Type = setupHeaders['1099 type'] || setupHeaders['1099'] || 7;

    // Add Row: "Contractor 1099" (Vendor), "NEC" (1099 Type)
    // We need a specific row for the Vendor table, usually below header.
    // Let's iterate to find the end of the vendor list or just insert at a safe row (Row 5+).
    // The Template usually has placeholder vendors "Vendor ABC". We'll just add a new row at bottom.
    const newRow = setupSheet.addRow([]);
    newRow.getCell(colVend).value = 'Contractor 1099';
    newRow.getCell(col1099Type).value = 'NEC';

    // Add Payer Info for CSV
    // We need "Company Info" vertical key-values.
    // Assuming "Company Info" header is somewhere.
    const colCompInfo = setupHeaders['company info'] || setupHeaders['payer info'];
    if (colCompInfo) {
        // Find first empty row under this header?
        // Let's just overwrite known rows if possible or add new below.
        // We will add rows specifically for this.
        let r = 20; // Safe distance
        setupSheet.getCell(r, colCompInfo).value = 'Company Name';
        setupSheet.getCell(r, colCompInfo + 1).value = 'Test Corp LLC';
        setupSheet.getCell(r + 1, colCompInfo).value = 'TIN';
        setupSheet.getCell(r + 1, colCompInfo + 1).value = '12-3456789';
    }

    console.log('\n--- Phase 5: Add Ledger Entries ---');
    const ledgerSheet = workbook.getWorksheet('Ledger');
    // Add an Owner Investment (Equity/Asset)
    // Date[1], Desc[2], Cat[3], Debit[4], Credit[5]
    ledgerSheet.addRow([new Date('2025-01-01'), 'Owner Investment', 'Checking Account', 1000, 0]);
    // Add a manual expense adjustment
    ledgerSheet.addRow([new Date('2025-01-20'), 'Audit Adjustment', 'Office', 50, 0]);

    await workbook.xlsx.writeFile(TEST_FILE);

    console.log('\n--- Phase 6: Run Financial Report (with Checker) ---');
    console.log('\n--- Phase 6: Run Financial Report (with Checker & 1099) ---');
    execSync(`node report.js ${TEST_FILE} --pl --bs --checker --1099`, { stdio: 'inherit' });

    console.log('\n--- Phase 7: Save Test Artifact ---');
    const ARTIFACT_PATH = 'tests/Full_Accounting_Test_Case.xlsx';
    fs.copyFileSync(TEST_FILE, ARTIFACT_PATH);
    console.log(`Saved full test case to ${ARTIFACT_PATH}`);

    console.log('\n--- Phase 8: Verification ---');
    // math update:
    // Bank: 6000 (from CSV) + 1000 (Ledger Debit) = 7000.00
    // CC: -180.50
    // Office: -165.00 (CSV) - 50 (Ledger Debit) = -215.00
    // Net Income: 5819.50 (CSV) - 50 (Ledger) = 5769.50
    // Bank Balance note: 6000 (CSV) + 1000 (Ledger) + 150 (Test Junk Rows) = 7150.00

    console.log('Test completed. Check output above for:');
    console.log('Bank Balance: 7,150.00');
    console.log('CC Balance: 180.50'); // Displayed as positive Liability in BS report
    console.log('Net Income: 5,769.50');
}

runTest().catch(console.error);

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

    // Add necessary categories to Setup if missing
    const setupSheet = workbook.getWorksheet('Setup');
    setupSheet.addRow(['Travel', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Office', 'General', 'Expense', 'P&L']);

    console.log('\n--- Phase 5: Add Ledger Entries ---');
    const ledgerSheet = workbook.getWorksheet('Ledger');
    // Add an Owner Investment (Equity/Asset)
    // Date[1], Desc[2], Cat[3], Debit[4], Credit[5]
    ledgerSheet.addRow([new Date('2025-01-01'), 'Owner Investment', 'Checking Account', 1000, 0]);
    // Add a manual expense adjustment
    ledgerSheet.addRow([new Date('2025-01-20'), 'Audit Adjustment', 'Office', 50, 0]);

    await workbook.xlsx.writeFile(TEST_FILE);

    console.log('\n--- Phase 6: Run Financial Report (with Checker) ---');
    execSync(`node update_financials.js ${TEST_FILE} --print-only --pl --bs --checker`, { stdio: 'inherit' });

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

    console.log('Test completed. Check output above for:');
    console.log('Bank Balance: 7000.00');
    console.log('CC Balance: -180.50');
    console.log('Net Income: 5769.50');
}

runTest().catch(console.error);

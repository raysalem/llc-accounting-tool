/**
 * Test: Extended --details Flag
 * 
 * Verifies that the --details flag supports filtering by:
 * 1. Category
 * 2. Vendor
 * 3. Customer
 */

const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const testFile = path.join(__dirname, 'Test_Details_Extended.xlsx');

async function createTestWorkbook() {
    const wb = new ExcelJS.Workbook();

    // 1. Setup
    const setup = wb.addWorksheet('Setup');
    setup.getRow(1).values = ['Category', '', 'Type', 'Report', '', '', 'Vendors', '', 'Customers'];
    setup.getRow(2).values = ['DetailsCat', '', 'Expense', 'P&L'];
    setup.getRow(2).getCell(7).value = 'DetailsVendor';
    setup.getRow(2).getCell(9).value = 'DetailsCustomer';

    // Sheet Config
    setup.getRow(1).getCell(15).value = 'SheetName';
    setup.getRow(1).getCell(16).value = 'Type';
    setup.getRow(1).getCell(17).value = 'Category';
    setup.getRow(1).getCell(18).value = 'Offset';
    setup.getRow(2).getCell(15).value = 'Bank';
    setup.getRow(2).getCell(16).value = 'Asset';
    setup.getRow(2).getCell(17).value = 'Bank';
    setup.getRow(2).getCell(18).value = 1;

    // 2. Ledger
    wb.addWorksheet('Ledger').getRow(1).values = ['Date', 'Description', 'Category', 'Debit', 'Credit'];

    // 3. Transactions (Bank)
    const bank = wb.addWorksheet('Bank');
    bank.getRow(1).values = ['Date', 'Description', 'Amount', 'Category', 'SubCategory', '', 'Vendor', 'Customer'];

    // Row 2: Match Category Only
    bank.getRow(2).values = ['2025-01-01', 'Cat trans', -10, 'DetailsCat', '', '', '', ''];
    // Row 3: Match Vendor Only
    bank.getRow(3).values = ['2025-01-02', 'Vendor trans', -20, 'OtherCat', '', '', 'DetailsVendor', ''];
    // Row 4: Match Customer Only
    bank.getRow(4).values = ['2025-01-03', 'Cust trans', 30, 'OtherCat', '', '', '', 'DetailsCustomer'];

    await wb.xlsx.writeFile(testFile);
}

function runCheck(filter, expectedDesc, label) {
    console.log(`\nChecking --details "${filter}"...`);
    try {
        const output = execSync(`node report.js "${testFile}" --details "${filter}"`, {
            cwd: path.join(__dirname, '..'), // Run from root
            encoding: 'utf8'
        });

        if (output.includes(expectedDesc)) {
            console.log(`✓ PASS: Found "${expectedDesc}" when filtering by ${label}`);
        } else {
            console.error(`❌ FAIL: Did not find "${expectedDesc}" when filtering by ${label}`);
            console.log('Output:', output);
            process.exit(1);
        }
    } catch (e) {
        console.error(`❌ FAIL: Command failed for ${label}`);
        process.exit(1);
    }
}

async function runTest() {
    try {
        await createTestWorkbook();

        // 1. Check Category Match
        runCheck('DetailsCat', 'Cat trans', 'Category');

        // 2. Check Vendor Match
        runCheck('DetailsVendor', 'Vendor trans', 'Vendor');

        // 3. Check Customer Match
        runCheck('DetailsCustomer', 'Cust trans', 'Customer');

        console.log('\n✅ ALL TESTS PASSED - Extended Details Logic Verified\n');
        fs.unlinkSync(testFile);

    } catch (e) {
        console.error('Test Failed:', e);
        if (fs.existsSync(testFile)) fs.unlinkSync(testFile);
        process.exit(1);
    }
}

runTest();

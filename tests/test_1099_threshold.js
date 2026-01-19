/**
 * Test: 1099 Threshold Compliance
 * 
 * This test verifies that the --vendor report correctly applies threshold logic
 * to the "Required" column:
 * - NEC vendors must have >= $600 to show "Required: YES"
 * - INT vendors must have > $0 to show "Required: YES"
 * - Vendors below threshold should show type but NOT "Required: YES"
 */

const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

console.log('=== 1099 Threshold Compliance Test ===\n');

// Create a test workbook with specific vendor amounts
const ExcelJS = require('exceljs');
const testFile = path.join(__dirname, 'Test_1099_Threshold.xlsx');

async function createTestWorkbook() {
    const wb = new ExcelJS.Workbook();

    // Setup sheet
    const setup = wb.addWorksheet('Setup');
    setup.getRow(1).values = ['Category', 'Sub-Category', 'Type', 'Report', '', '', 'Vendors', '1099', 'Customers'];

    // Categories
    setup.getRow(2).values = ['Services', '', 'Expense', 'P&L'];

    // Vendors with 1099 status
    setup.getRow(2).getCell(7).value = 'Vendor A'; // Will pay $700 (NEC, should be required)
    setup.getRow(2).getCell(8).value = 'NEC';

    setup.getRow(3).getCell(7).value = 'Vendor B'; // Will pay $500 (NEC, should NOT be required)
    setup.getRow(3).getCell(8).value = 'NEC';

    setup.getRow(4).getCell(7).value = 'Vendor C'; // Will pay $100 (INT, should be required)
    setup.getRow(4).getCell(8).value = 'INT';

    setup.getRow(5).getCell(7).value = 'Vendor D'; // Will pay $0 (INT, should NOT be required)
    setup.getRow(5).getCell(8).value = 'INT';

    // Ledger sheet (required by report.js)
    const ledger = wb.addWorksheet('Ledger');
    ledger.getRow(1).values = ['Date', 'Description', 'Category', 'Debit', 'Credit'];

    // Bank Transactions sheet
    const bank = wb.addWorksheet('Bank Transactions');
    bank.getRow(1).values = ['Date', 'Description', 'Amount', 'Category', 'Vendor'];
    bank.getRow(2).values = ['2025-01-15', 'Payment to A', -700, 'Services', 'Vendor A'];
    bank.getRow(3).values = ['2025-01-16', 'Payment to B', -500, 'Services', 'Vendor B'];
    bank.getRow(4).values = ['2025-01-17', 'Interest to C', -100, 'Services', 'Vendor C'];
    bank.getRow(5).values = ['2025-01-18', 'Zero payment D', 0, 'Services', 'Vendor D'];

    await wb.xlsx.writeFile(testFile);
    console.log(`✓ Created test workbook: ${testFile}`);
}

async function runTest() {
    try {
        // Create test data
        await createTestWorkbook();

        // Run report
        console.log('\nRunning vendor report...\n');
        const output = execSync(`node report.js "${testFile}" --vendor`, {
            cwd: path.join(__dirname, '..'),
            encoding: 'utf8'
        });

        console.log(output);

        // Parse output and verify
        const lines = output.split('\n');
        const vendorLines = lines.filter(l => l.includes('Vendor'));

        let passed = true;
        const checks = [
            { vendor: 'Vendor A', amount: 700, type: 'NEC', shouldBeRequired: true },
            { vendor: 'Vendor B', amount: 500, type: 'NEC', shouldBeRequired: false },
            { vendor: 'Vendor C', amount: 100, type: 'INT', shouldBeRequired: true },
            { vendor: 'Vendor D', amount: 0, type: 'INT', shouldBeRequired: false }
        ];

        console.log('\n=== Verification ===\n');

        for (const check of checks) {
            const line = vendorLines.find(l => l.includes(check.vendor));
            if (!line) {
                console.log(`✗ FAIL: ${check.vendor} not found in output`);
                passed = false;
                continue;
            }

            const hasType = line.includes(check.type);
            const hasRequired = line.includes('YES');

            if (!hasType) {
                console.log(`✗ FAIL: ${check.vendor} should show type ${check.type}`);
                passed = false;
            } else if (hasRequired !== check.shouldBeRequired) {
                console.log(`✗ FAIL: ${check.vendor} ($${check.amount}, ${check.type}) - Required should be ${check.shouldBeRequired ? 'YES' : 'blank'}, got ${hasRequired ? 'YES' : 'blank'}`);
                passed = false;
            } else {
                console.log(`✓ PASS: ${check.vendor} ($${check.amount}, ${check.type}) - Required correctly shows ${hasRequired ? 'YES' : 'blank'}`);
            }
        }

        // Cleanup
        fs.unlinkSync(testFile);
        console.log(`\n✓ Cleaned up test file`);

        if (passed) {
            console.log('\n✅ ALL TESTS PASSED - 1099 threshold logic is compliant\n');
            process.exit(0);
        } else {
            console.log('\n❌ SOME TESTS FAILED - Review threshold logic\n');
            process.exit(1);
        }

    } catch (error) {
        console.error('Test failed with error:', error.message);
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
        process.exit(1);
    }
}

runTest();

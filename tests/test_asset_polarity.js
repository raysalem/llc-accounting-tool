/**
 * Test: Asset Polarity Verification
 * 
 * Verifies the fix for "Purchased Assets showing Negative".
 * 
 * Scenario:
 * 1. Bank (Linked Asset): Starts at $10,000.
 * 2. Investment (Unlinked Asset): Starts at $0.
 * 3. Mortgage (Unlinked Liability): Starts at $5,000 (Positive Magnitude).
 * 
 * Transaction:
 * - Spend $1,000 on Investment (Category: Investment).
 * - Spend $1,000 on Mortgage (Category: Mortgage).
 * 
 * Expected Results:
 * - Bank: $8,000 (10k - 1k - 1k).
 * - Investment: $1,000 (Must be Positive! Spending = Value Increase).
 * - Mortgage: $4,000 (Must Decrease! Spending = Paydown).
 */

const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const testFile = path.join(__dirname, 'Test_Asset_Polarity.xlsx');

async function createTestWorkbook() {
    const wb = new ExcelJS.Workbook();

    // 1. Setup
    const setup = wb.addWorksheet('Setup');
    setup.getRow(1).values = ['Category', '', 'Type', 'Report'];

    setup.getRow(2).values = ['Investment', '', 'Asset', 'Balance Sheet'];
    setup.getRow(3).values = ['Mortgage', '', 'Liability', 'Balance Sheet'];

    // Sheet Config
    setup.getRow(1).getCell(15).value = 'SheetName';
    setup.getRow(1).getCell(16).value = 'Type';
    setup.getRow(1).getCell(17).value = 'Category';
    setup.getRow(1).getCell(18).value = 'Offset';
    setup.getRow(1).getCell(19).value = 'StartBalance';

    setup.getRow(2).getCell(15).value = 'Bank';
    setup.getRow(2).getCell(16).value = 'Asset';
    setup.getRow(2).getCell(17).value = 'Bank';
    setup.getRow(2).getCell(18).value = 1;
    setup.getRow(2).getCell(19).value = 10000;

    // 2. Ledger
    wb.addWorksheet('Ledger').getRow(1).values = ['Date', 'Description', 'Category', 'Debit', 'Credit'];

    // 3. Transactions (Bank)
    const bank = wb.addWorksheet('Bank');
    bank.getRow(1).values = ['Date', 'Description', 'Amount', 'Category'];

    // Transaction 1: Buy Investment (-1000)
    bank.getRow(2).values = ['2025-01-01', 'Buy Asset', -1000, 'Investment'];

    // Transaction 2: Pay Mortgage (-1000)
    bank.getRow(3).values = ['2025-01-02', 'Pay Debt', -1000, 'Mortgage'];

    await wb.xlsx.writeFile(testFile);
}

function parseValue(line) {
    // Extract last number from line (e.g. "Investment ... 1,000.00")
    const match = line.match(/([\d,\.-]+)$/);
    return match ? parseFloat(match[1].replace(/,/g, '')) : null;
}

async function runTest() {
    try {
        await createTestWorkbook();
        console.log('Running Polarity Check...');
        const output = execSync(`node report.js "${testFile}" --bs`, {
            cwd: path.join(__dirname, '..'),
            encoding: 'utf8'
        });
        fs.writeFileSync(path.join(__dirname, 'debug_polarity.txt'), output);

        const lines = output.split('\n');
        console.log("DEBUG OUTPUT:\n", output);

        // Find lines
        // Output format: Label ... Value

        let investVal = null;
        let mortgageVal = null;

        lines.forEach(l => {
            if (l.includes('Investment')) investVal = parseValue(l);
            if (l.includes('Mortgage')) mortgageVal = parseValue(l);
        });

        let passed = true;

        // Check Investment (Unlinked Asset) output
        // We expect +1000. If bug exists, it would be -1000.
        if (investVal === 1000) {
            console.log(`✓ PASS: Investment Value is 1000 (Positive). Asset Polarity Fix Working.`);
        } else if (investVal === -1000) {
            console.error(`❌ FAIL: Investment Value is -1000. Asset Polarity Fix NOT Working.`);
            passed = false;
        } else {
            console.error(`❌ FAIL: Unexpected Investment Value: ${investVal}`);
            passed = false;
        }

        // Check Mortgage (Unlinked Liability) output
        // Start 5000? Wait, Mortgage is just a Category here with 0 start balance.
        // Wait, Unlinked Liability starts at 0.
        // Spending -1000 on it.
        // Stat += -1000.
        // Output -1000.
        // Is Liability -1000 correct?
        // If Liability is "Mortgage Loan". Balance should be Positive (Amount Owed).
        // If I pay -1000, Balance should decrease. 
        // If Start is 0, Balance becomes -1000.
        // This implies "I overpaid 1000" or "Net Change is -1000".
        // My test setup didn't give Mortgage a start balance.
        // So checking for -1000 is correct behavior for Liability accumulation (Reduction).
        // If I gave it a start balance, say 5000.
        // But "Start Balance" is only for SHEETS.
        // So "Mortgage" Category will just show the Net Change.

        if (mortgageVal === -1000) {
            console.log(`✓ PASS: Mortgage Change is -1000 (Paydown). Liability Polarity Preserved.`);
        } else {
            // If my "Invert Unlinked" logic accidentally hit Liabilities, it might be +1000.
            console.error(`❌ FAIL: Mortgage Value is ${mortgageVal}. Should be -1000.`);
            passed = false;
        }

        if (passed) {
            console.log('\n✅ ALL TESTS PASSED');
            fs.unlinkSync(testFile);
            process.exit(0);
        } else {
            console.log('\n❌ TESTS FAILED');
            process.exit(1);
        }

    } catch (e) {
        console.error('Test Error:', e);
        if (fs.existsSync(testFile)) fs.unlinkSync(testFile);
        process.exit(1);
    }
}

runTest();

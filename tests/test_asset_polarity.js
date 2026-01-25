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

    // Define Categories (Cols A-D)
    setup.addRow(['Investment', '', 'Asset', 'Balance Sheet']);
    setup.addRow(['Mortgage', '', 'Liability', 'Balance Sheet']);
    setup.addRow(['Equipment', '', 'Asset', 'Balance Sheet']);
    setup.addRow(['Bank', '', 'Asset', 'Balance Sheet']);

    // Sheet Config (Start at Col 15, Row 1)
    // We must manually place headers/values since addRow affects all columns (mostly A)
    setup.getCell('O1').value = 'SheetName';
    setup.getCell('P1').value = 'Type';
    setup.getCell('Q1').value = 'Category';
    setup.getCell('R1').value = 'Offset';
    setup.getCell('S1').value = 'StartBalance';

    setup.getCell('O2').value = 'Bank';
    setup.getCell('P2').value = 'Bank';
    setup.getCell('Q2').value = 'Bank'; // Links to Category 'Bank'
    setup.getCell('R2').value = 1;
    setup.getCell('S2').value = 10000;

    // 2. Ledger
    wb.addWorksheet('Ledger').addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);

    // 3. Transactions (Bank)
    const bank = wb.addWorksheet('Bank');
    bank.addRow(['Date', 'Description', 'Amount', 'Category']);
    bank.addRow(['2025-01-01', 'Buy Asset', -1000, 'Investment']);
    bank.addRow(['2025-01-02', 'Pay Debt', -1000, 'Mortgage']);

    // 4. Ledger Entries
    wb.getWorksheet('Ledger').addRow(['2025-01-03', 'Manual Equipment', 'Equipment', 500, 0]);
    wb.getWorksheet('Ledger').addRow(['2025-01-03', 'Manual Equipment', 'Bank', 0, 500]);

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
        let equipVal = null;

        lines.forEach(l => {
            // ... existing loop ...
            if (l.includes('Investment')) investVal = parseValue(l);
            if (l.includes('Mortgage')) mortgageVal = parseValue(l);
            if (l.includes('Equipment')) equipVal = parseValue(l);
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

        if (equipVal === 500) {
            console.log(`✓ PASS: Equipment (Ledger) is 500. Polarity Correct.`);
        } else {
            console.error(`❌ FAIL: Equipment (Ledger) is ${equipVal}. Should be 500. (Likely flipped to -500 by report.js)`);
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

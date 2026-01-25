const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const ExcelJS = require('exceljs');

// Test: Ledger Strictness (Hard Fail on Missing Date)
// 1. Create a dummy Excel file with a bad ledger row (Missing Date).
// 2. Run report.js.
// 3. Functional Success: Exit Code 1 (Crash).
// 4. Verification: Check output for "[CRITICAL ERROR]" message.

const TEST_FILE = path.join(__dirname, 'temp_ledger_strictness.xlsx');

async function createTestFile() {
    const workbook = new ExcelJS.Workbook();

    // Setup Sheet
    const setup = workbook.addWorksheet('Setup');
    setup.addRow(['Category', 'Type', 'Report', 'SheetName', 'Type', 'ShortName', 'FlipPolarity']);
    setup.addRow(['Office Supplies', 'Expense', 'P&L', 'Bank', 'Bank', 'Bank', 'No']);

    // Ledger Sheet
    const ledger = workbook.addWorksheet('General Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']); // Header
    ledger.addRow([new Date(), 'Valid Entry', 'Office Supplies', 100, 0]);
    ledger.addRow([null, 'Invalid Entry (No Date)', 'Office Supplies', 50, 0]); // BAD ROW
    ledger.addRow([new Date(), 'Valid Entry 2', 'Office Supplies', 0, 100]); // Credit to balance

    await workbook.xlsx.writeFile(TEST_FILE);
}

async function runTest() {
    await createTestFile();
    console.log('[Test] Created test file with invalid ledger row.');

    try {
        // Run report.js --checker
        // We EXPECT this to fail using the catch block
        execSync(`node report.js "${TEST_FILE}" --checker 2>&1`, { encoding: 'utf8' });

        // If we get here, it didn't crash -> FAIL
        console.error('[FAIL] Script failed to crash. Zero tolerance enforcement is missing.');
        process.exit(1);

    } catch (e) {
        // Expecting exit code 1
        const output = e.stdout || e.toString();

        if (output.includes('[CRITICAL ERROR] Ledger Row 3: Missing Date')) {
            console.log('[PASS] Script crashed correctly with Critical Error.');
        } else {
            console.error('[FAIL] Script crashed but message was unexpected.');
            console.error('Output:', output);
            process.exit(1);
        }
    } finally {
        // Cleanup
        if (fs.existsSync(TEST_FILE)) fs.unlinkSync(TEST_FILE);
    }
}

runTest();

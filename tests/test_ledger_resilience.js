const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const ExcelJS = require('exceljs');

// Test: Ledger Resilience (Soft Fail on Missing Date)
// 1. Create a dummy Excel file with a bad ledger row (Missing Date).
// 2. Run report.js.
// 3. functional success: Exit Code 0.
// 4. Verification: Check output for "[CRITICAL WARNING]" message.

const TEST_FILE = path.join(__dirname, 'temp_ledger_resilience.xlsx');

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
        const output = execSync(`node report.js "${TEST_FILE}" --checker 2>&1`, { encoding: 'utf8' });

        console.log('[Test] Execution finished. Output snippet:');
        // console.log(output);

        if (output.includes('[CRITICAL WARNING] Ledger Row 3: Missing Date')) {
            console.log('[PASS] Warnings detected correctly.');
        } else {
            console.error('[FAIL] Expected warning message not found.');
            process.exit(1);
        }

        if (output.includes('Valid Entry 2')) {
            // How to check if Valid Entry 2 was processed? 
            // We can check if the total Debit includes 100+200=300, or at least 100.
            // But main goal is crash prevention.
            console.log('[PASS] Script continued execution.');
        }

    } catch (e) {
        console.error('[FAIL] Script crashed with exit code:', e.status);
        console.error(e.stderr);
        process.exit(1);
    } finally {
        // Cleanup
        if (fs.existsSync(TEST_FILE)) fs.unlinkSync(TEST_FILE);
    }
}

runTest();

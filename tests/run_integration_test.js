const { execSync } = require('child_process');
const fs = require('fs');
const ExcelJS = require('exceljs');
// Test files live in a temp directory so test runs never modify the repo.
const TMP_DIR = require('fs').mkdtempSync(require('path').join(require('os').tmpdir(), 'llc-test-'));

const TEST_FILE = require('path').join(TMP_DIR, 'Test_Accounting.xlsx');
let failures = 0;

function assertIncludes(output, text, label) {
    if (output.includes(text)) {
        console.log(`✅ [PASS] ${label}`);
    } else {
        console.error(`❌ [FAIL] ${label} — expected output to contain: ${JSON.stringify(text)}`);
        failures++;
    }
}

// Runs report.js and returns { code, output } without throwing, so the test can
// assert on both the numbers and the exit code.
function runReport(args) {
    try {
        const output = execSync(`node report.js "${TEST_FILE}" ${args}`, { encoding: 'utf-8', stdio: 'pipe' });
        return { code: 0, output };
    } catch (e) {
        return { code: e.status, output: (e.stdout || '').toString() + (e.stderr || '').toString() };
    }
}

async function runTest() {
    console.log('--- Phase 1: Initialize Template ---');
    execSync(`node generate_excel.js "${TEST_FILE}"`, { stdio: 'inherit' });

    console.log('\n--- Phase 2: Load Bank Transactions ---');
    execSync(`node load_transactions.js tests/example_bank.csv bank "${TEST_FILE}" --clear`, { stdio: 'inherit' });

    console.log('\n--- Phase 3: Load CC Transactions ---');
    execSync(`node load_transactions.js tests/example_cc.csv cc "${TEST_FILE}" --clear`, { stdio: 'inherit' });

    console.log('\n--- Phase 4: Categorize Transactions (Simulated) ---');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEST_FILE);

    // Transaction sheets: rows 1-2 are TOTAL/SUBTOTAL, row 3 is the header.
    const bankSheet = workbook.getWorksheet('Bank Transactions');
    bankSheet.eachRow((row, r) => {
        if (r <= 3) return;
        const desc = (row.getCell(2).value || '').toString();
        if (desc.includes('Salary')) row.getCell(4).value = 'Sales';
        if (desc.includes('Rent')) row.getCell(4).value = 'Rent';
        if (desc.includes('Client')) row.getCell(4).value = 'Sales';
    });

    const ccSheet = workbook.getWorksheet('Credit Card Transactions');
    ccSheet.eachRow((row, r) => {
        if (r <= 3) return;
        const desc = (row.getCell(3).value || '').toString();
        if (desc.includes('Starbucks')) row.getCell(5).value = 'Travel';
        if (desc.includes('Amazon')) row.getCell(5).value = 'Office';
        if (desc.includes('AWS')) row.getCell(5).value = 'Office';
    });

    // 1099 Test: > $600 payment to an NEC vendor
    bankSheet.addRow([new Date('2025-01-27'), 'Major Contract Work', -737.50, 'Services', '', '', 'Contractor 1099', '']);

    const setupSheet = workbook.getWorksheet('Setup');
    setupSheet.addRow(['Travel', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Office', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Services', 'General', 'Expense', 'P&L']);
    setupSheet.addRow(['Owner Equity', 'Capital', 'Equity', 'Balance Sheet']);

    const setupHeaders = {};
    setupSheet.getRow(1).eachCell((c, col) => {
        if (c.value) setupHeaders[c.value.toString().toLowerCase()] = col;
    });
    const colVend = setupHeaders['vendors'] || 6;
    // The template has no 1099 column; add one in a free column so report.js detects it.
    const col1099Type = setupHeaders['1099 type'] || 14;
    setupSheet.getRow(1).getCell(col1099Type).value = '1099 Type';
    const newRow = setupSheet.addRow([]);
    newRow.getCell(colVend).value = 'Contractor 1099';
    newRow.getCell(col1099Type).value = 'NEC';

    console.log('\n--- Phase 5: Add Ledger Entries (balanced) ---');
    const ledgerSheet = workbook.getWorksheet('Ledger');
    // Date, Desc, Category, Debit, Credit
    ledgerSheet.addRow([new Date('2025-01-01'), 'Owner Investment', 'Checking Account', 1000, 0]);
    ledgerSheet.addRow([new Date('2025-01-01'), 'Owner Investment', 'Owner Equity', 0, 1000]);
    ledgerSheet.addRow([new Date('2025-01-20'), 'Audit Adjustment', 'Office', 50, 0]);
    ledgerSheet.addRow([new Date('2025-01-20'), 'Audit Adjustment', 'Checking Account', 0, 50]);

    await workbook.xlsx.writeFile(TEST_FILE);

    console.log('\n--- Phase 6: Clean run (P&L + BS + 1099) ---');
    // Expected math:
    //   Bank:   5000 - 1500 + 2500 - 737.50 (CSV + contractor) + 1000 - 50 (ledger) = 6212.50
    //   CC:     15.50 + 120 + 45 = 180.50 liability
    //   Office: 165 (CC) + 50 (ledger) = 215.00
    //   Net income: 7500 - 1500 - 737.50 - 15.50 - 215 = 5032.00
    const clean = runReport('--pl --bs --year=2025');
    console.log(clean.output);
    if (clean.code === 0) console.log('✅ [PASS] Clean run exits 0');
    else { console.error(`❌ [FAIL] Clean run exited with code ${clean.code}`); failures++; }
    assertIncludes(clean.output, '7,500.00', 'Sales total');
    assertIncludes(clean.output, '5,032.00', 'Net income');
    assertIncludes(clean.output, '6,212.50', 'Bank balance (CSV + ledger)');
    assertIncludes(clean.output, '180.50', 'CC liability');

    // 1099: the contractor is over the 2025 $600 threshold but has no TIN/address,
    // so the vendor is listed and the run must fail until details are filled in.
    const run1099 = runReport('--1099 --year=2025');
    assertIncludes(run1099.output, 'Contractor 1099', '1099-NEC vendor listed');
    assertIncludes(run1099.output, 'Incomplete Data', 'Missing vendor 1099 details flagged');
    if (run1099.code !== 0) console.log('✅ [PASS] 1099 run exits non-zero when vendor details are missing');
    else { console.error('❌ [FAIL] 1099 run exited 0 despite missing vendor details'); failures++; }

    console.log('\n--- Phase 7: Tax year filter ---');
    const otherYear = runReport('--pl --year=2026');
    assertIncludes(otherYear.output, 'Tax year: 2026 (1099-NEC threshold: $2,000)', '2026 uses the $2,000 NEC threshold');
    assertIncludes(otherYear.output, 'Skipped 11 transaction row(s) dated outside 2026', '2025 rows (7 transactions + 4 ledger) skipped when reporting 2026');
    assertIncludes(otherYear.output, 'NET INCOME: 0.00', 'No 2025 income in a 2026 report');

    console.log('\n--- Phase 8: Integrity checker flags bad rows ---');
    await workbook.xlsx.readFile(TEST_FILE);
    const dirtyBank = workbook.getWorksheet('Bank Transactions');
    dirtyBank.addRow([new Date('2025-01-25'), 'Mystery Corp', 100, 'IllegalCat', '', '', 'UnknownVendor', '']);
    dirtyBank.addRow([new Date('2025-01-26'), 'Uncategorized Expense', 50, '', '', '', '', '']);
    await workbook.xlsx.writeFile(TEST_FILE);

    const dirty = runReport('--checker --year=2025');
    if (dirty.code !== 0) console.log('✅ [PASS] Checker run exits non-zero on bad data');
    else { console.error('❌ [FAIL] Checker run exited 0 despite bad data'); failures++; }
    assertIncludes(dirty.output, 'IllegalCat', 'Illegal category flagged');
    assertIncludes(dirty.output, 'Uncategorized Expense', 'Missing category flagged');

    console.log('\n--- Phase 9: Save Test Artifact ---');
    // Only refresh the committed artifact when asked, so test runs don't modify the repo.
    if (process.env.SAVE_ARTIFACT) {
        fs.copyFileSync(TEST_FILE, 'tests/Full_Accounting_Test_Case.xlsx');
        console.log('Saved tests/Full_Accounting_Test_Case.xlsx');
    }
    fs.rmSync(TMP_DIR, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} assertion(s) failed.`);
        process.exit(1);
    }
    console.log('\nIntegration test passed.');
}

runTest().catch(err => {
    console.error(err);
    process.exit(1);
});

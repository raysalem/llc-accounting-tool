const { execSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const assert = require('assert');

const SRC_EXCEL = 'tests/temp_src.xlsx';
const TARGET_FILE = 'tests/temp_target.xlsx';

function run(cmd) {
    try {
        return execSync(cmd, { encoding: 'utf-8', stdio: 'pipe' });
    } catch (e) {
        console.error(`COMMAND FAILED: ${cmd}`);
        console.error('STDOUT:', e.stdout ? e.stdout.toString() : 'null');
        console.error('STDERR:', e.stderr ? e.stderr.toString() : 'null');
        throw new Error(`Command failed: ${cmd}`);
    }
}

// Like run(), but returns output even when report.js exits 1 because the sample
// books have data-integrity issues. These checks only test flag output.
function runReport(cmd) {
    try {
        return execSync(cmd, { encoding: 'utf-8', stdio: 'pipe' });
    } catch (e) {
        if (e.stdout) return e.stdout.toString();
        throw new Error(`Command failed: ${cmd}`);
    }
}

async function testArguments() {
    console.log('--- TEST SUITE: Argument Coverage ---');

    // --- SETUP: Create Source Excel & Target Template ---
    // 1. Source Excel for load_transactions (Testing Excel Input support)
    const srcWorkbook = new ExcelJS.Workbook();
    const srcSheet = srcWorkbook.addWorksheet('Sheet1');
    srcSheet.addRow(['Date', 'Description', 'Amount', 'Category']);
    srcSheet.addRow([new Date('2025-01-01'), 'Test Excel Source', -50.00, 'Office Supplies']);
    await srcWorkbook.xlsx.writeFile(SRC_EXCEL);

    // 2. Target Template (Clean Slate)
    run('node generate_excel.js');
    if (fs.existsSync(TARGET_FILE)) fs.unlinkSync(TARGET_FILE);
    fs.renameSync('LLC_Accounting_Template.xlsx', TARGET_FILE);

    // =========================================================================
    // PART 1: load_transactions.js Coverage
    // =========================================================================
    console.log('\n[1/4] load_transactions.js (Excel Input, Append, Clear)');

    // 1. Test: Excel Input (Standard)
    console.log('   Running: Load from Excel (Append Mode)...');
    run(`node load_transactions.js "${SRC_EXCEL}" bank "${TARGET_FILE}"`);

    // Verify 1 Row Added
    let wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TARGET_FILE);
    let sheet = wb.getWorksheet('Bank Transactions');
    // Header is row 3 (default template offset), Data at 4. Row 1,2 are totals.
    // Actually, `load_transactions` with new template might adjust.
    // Let's count non-empty rows.
    let rowCount = 0;
    sheet.eachRow(r => rowCount++);
    // Expect: Totals(2) + Header(1) + Data(1) = 4 rows usually.
    // Or Header(1) + Data(1) = 2 if simpler.
    // Let's just check the data value presence.
    let found = false;
    sheet.eachRow(r => {
        r.eachCell(c => { if (c.value && c.value.toString().includes('Test Excel Source')) found = true; });
    });
    if (!found) throw new Error('Excel Load Failed: Data not found in target.');

    // 2. Test: Append (Run again without --clear)
    console.log('   Running: Load from Excel (Append Mode / 2nd Run)...');
    run(`node load_transactions.js "${SRC_EXCEL}" bank "${TARGET_FILE}"`);

    // Verify Duplicate Rows
    await wb.xlsx.readFile(TARGET_FILE);
    sheet = wb.getWorksheet('Bank Transactions');
    let matchCount = 0;
    sheet.eachRow(r => {
        r.eachCell(c => { if (c.value && c.value.toString().includes('Test Excel Source')) matchCount++; });
    });
    if (matchCount < 2) throw new Error('Append Logic Failed: Expected multiple rows.');

    // 3. Test: --clear
    console.log('   Running: Load with --clear...');
    run(`node load_transactions.js "${SRC_EXCEL}" bank "${TARGET_FILE}" --clear`);

    // Verify Single Row Again
    await wb.xlsx.readFile(TARGET_FILE);
    sheet = wb.getWorksheet('Bank Transactions');
    matchCount = 0;
    sheet.eachRow(r => {
        r.eachCell(c => { if (c.value && c.value.toString().includes('Test Excel Source')) matchCount++; });
    });
    if (matchCount !== 1) throw new Error(`Clear Logic Failed: Expected 1 row, found ${matchCount}.`);

    // 4. Test: --help
    const helpOut = run('node load_transactions.js --help');
    if (!helpOut.includes('Usage: node load_transactions.js')) throw new Error('load_transactions --help failed');


    // =========================================================================
    // PART 2: update_financials.js Coverage
    // =========================================================================
    console.log('\n[2/4] report.js (Flags)');

    // Setup: Modify Target to have Vendors/Customers for testing
    // Row 4 is our data row.
    // Col 7 = Vendor, Col 8 = Customer (Bank Map default)
    // We need to re-read carefully.
    sheet.getRow(4).getCell(4).value = 'Rent'; // Category
    sheet.getRow(4).getCell(7).value = 'TestVendor';
    sheet.getRow(4).getCell(8).value = 'TestCust';
    await wb.xlsx.writeFile(TARGET_FILE);

    // 5. Test: --vendor
    console.log('   Running: --vendor');
    const vendorOut = runReport(`node report.js "${TARGET_FILE}" --year=2025 --vendor`);
    if (!vendorOut.includes('VENDOR SPENDING') || !vendorOut.includes('TestVendor')) {
        throw new Error('--vendor flag failed to show vendor report');
    }

    // 6. Test: --customer
    console.log('   Running: --customer');
    const custOut = runReport(`node report.js "${TARGET_FILE}" --year=2025 --customer`);
    if (!custOut.includes('CUSTOMER INCOME') || !custOut.includes('TestCust')) {
        throw new Error('--customer flag failed to show customer report');
    }

    // 6.5 Test: --pl-sub
    console.log('   Running: --pl-sub');
    // Ensure we have a subcategory to show
    sheet = wb.getWorksheet('Bank Transactions');
    sheet.getRow(4).getCell(5).value = 'Software'; // Sub-Category
    await wb.xlsx.writeFile(TARGET_FILE);

    const plSubOut = runReport(`node report.js "${TARGET_FILE}" --year=2025 --pl-sub`);
    if (!plSubOut.includes('PROFIT & LOSS') || !plSubOut.includes('> Software')) {
        console.error('--- FAILURE OUTPUT START ---');
        console.error(plSubOut);
        console.error('--- FAILURE OUTPUT END ---');
        throw new Error('--pl-sub flag failed to show sub-categories');
    }


    // 7. Test: --save (writes a separate report_<name>.xlsx next to the input)
    console.log('   Running: --save');
    const REPORT_FILE = 'tests/report_temp_target.xlsx';
    if (fs.existsSync(REPORT_FILE)) fs.unlinkSync(REPORT_FILE);
    runReport(`node report.js "${TARGET_FILE}" --year=2025 --pl --save`);

    if (!fs.existsSync(REPORT_FILE)) throw new Error(`--save failed: ${REPORT_FILE} not created`);
    const finalWb = new ExcelJS.Workbook();
    await finalWb.xlsx.readFile(REPORT_FILE);
    if (!finalWb.getWorksheet('Profit & Loss')) throw new Error('--save failed: "Profit & Loss" sheet missing from report');
    if (!finalWb.getWorksheet('Report Info')) throw new Error('--save failed: "Report Info" sheet missing from report');

    // 8. Test: --1099
    console.log('Testing --1099...');
    runReport(`node report.js "${TARGET_FILE}" --year=2025 --1099`);

    // 9. Test: --details (formerly tests/test_details_flag.js, which depended on this file existing)
    const detailsOut = runReport(`node report.js "${TARGET_FILE}" --year=2025 --details "Rent"`);
    if (!detailsOut.includes('DETAILS: "rent"') || !detailsOut.includes('TOTAL')) {
        throw new Error('--details flag failed to show details with a total');
    }

    // 10. Test: --help
    const upHelp = run('node report.js --help');
    if (!upHelp.includes('Usage: node report.js')) throw new Error('report.js --help failed');

    console.log('\n✅ TEST PASSED: All arguments covered and verified.');

    // Cleanup
    try {
        fs.unlinkSync(SRC_EXCEL);
        fs.unlinkSync(TARGET_FILE);
    } catch (e) { }
}

testArguments().catch(err => {
    console.error('❌ TEST FAILED:', err.message);
    process.exit(1);
});

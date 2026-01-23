const { execSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const TEST_FILE = 'tests/comprehensive_test_data.xlsx';

function run(cmd, allowPartialFailure = false) {
    try {
        console.log(`Running: ${cmd}`);
        return execSync(cmd, { encoding: 'utf-8', stdio: 'pipe' });
    } catch (e) {
        // If command failed but we got some output, return it for partial checking
        if (allowPartialFailure && e.stdout) {
            console.warn(`COMMAND PARTIALLY FAILED (returning output): ${cmd}`);
            return e.stdout.toString();
        }
        console.error(`COMMAND FAILED: ${cmd}`);
        console.error('STDOUT:', e.stdout ? e.stdout.toString() : 'null');
        console.error('STDERR:', e.stderr ? e.stderr.toString() : 'null');
        throw new Error(`Command failed: ${cmd}`);
    }
}

async function createTestData() {
    const wb = new ExcelJS.Workbook();

    // --- SETUP SHEET ---
    const setup = wb.addWorksheet('Setup');
    // Row 1: Headers (used by getHeaderMap in report.js)
    setup.addRow([
        'Category', 'Sub-Category', 'Account Type', 'Report',
        '',
        'Vendors', '1099 Type', '1099 Required', 'Business Name', 'TIN', 'Address', 'Email', 'Phone',
        'Customers',
        '',
        'Sheet Name', 'Type', 'Flip', 'Offset', 'Short Name',
        'Company Info', 'Value'
    ]);

    // Data rows
    setup.addRow(['Sales', 'Software', 'Income', 'P&L']);
    setup.addRow(['Sales', 'Consulting', 'Income', 'P&L']);
    setup.addRow(['Office Exp', 'Supplies', 'Expense', 'P&L']);
    setup.addRow(['Office Exp', 'Rent', 'Expense', 'P&L']);
    setup.addRow(['Professional Fees', 'Legal', 'Expense', 'P&L']);
    setup.addRow(['Cash', 'Bank', 'Asset', 'BS']);
    setup.addRow(['CC Liability', 'CC', 'Liability', 'BS']);

    // Set Vendor data (Starting Row 2)
    const setVend = (row, v, type, req, biz, tin) => {
        setup.getCell(`F${row}`).value = v;
        setup.getCell(`G${row}`).value = type;
        setup.getCell(`H${row}`).value = req;
        setup.getCell(`I${row}`).value = biz;
        setup.getCell(`J${row}`).value = tin;
    };
    setVend(2, 'Staples', 'NEC', 'No', 'Staples Inc', '');
    setVend(3, 'Landlord', '', 'No', '', '');
    setVend(4, 'Lawyer 1099', 'NEC', 'Yes', 'Law Firm LLC', '12-3456789');
    setVend(5, 'Bank Interest', 'INT', 'Yes', 'Bank ABC', '99-8888888');

    // Customers (Col N)
    setup.getCell(`N2`).value = 'Client A';
    setup.getCell(`N3`).value = 'Client B';

    // Sheet Config (Col P, starting at Row 15)
    setup.getCell(`P14`).value = 'Sheet Name';
    setup.getCell(`Q14`).value = 'Type';
    setup.getCell(`R14`).value = 'Flip';
    setup.getCell(`S14`).value = 'Offset';
    setup.getCell(`T14`).value = 'Short Name';

    const setSheet = (row, name, type, flip, off, short) => {
        setup.getCell(`P${row}`).value = name;
        setup.getCell(`Q${row}`).value = type;
        setup.getCell(`R${row}`).value = flip;
        setup.getCell(`S${row}`).value = off;
        setup.getCell(`T${row}`).value = short;
    };
    setSheet(15, 'Bank ABC', 'Bank', 'No', 1, 'Bank');
    setSheet(16, 'CC Amex', 'CC', 'Yes', 1, 'Amex');
    setSheet(17, 'Ledger', 'Ledger', 'No', 1, 'Ledger');

    // Payer Info (Col U-V, starting at Row 15)
    setup.getCell('U1').value = 'Company Info'; // Key Header for Column U detection
    setup.getCell('U15').value = 'Business Name';
    setup.getCell('V15').value = 'Test Company';
    setup.getCell('U16').value = 'TIN';
    setup.getCell('V16').value = '00-0000000';

    // --- BANK SHEET ---
    const bank = wb.addWorksheet('Bank ABC');
    bank.addRow(['Date', 'Description', 'Amount', 'Category', 'Sub-Category', 'Vendor', 'Customer']);
    bank.addRow(['2025-01-01', 'Client Payment', 5000, 'Sales', 'Software', '', 'Client A']); // Income
    bank.addRow(['2025-01-02', 'Office Rent', -2000, 'Office Exp', 'Rent', 'Landlord', '']);   // Expense
    bank.addRow(['2025-01-03', 'Categorized Item', -50, 'Office Exp', 'Supplies', 'Staples', '']); // Categorized for balance
    bank.addRow(['2025-01-04', 'Legal Fees Bank', -800, 'Professional Fees', 'Legal', 'Lawyer 1099', '']); // 1099 on Bank sheet
    bank.addRow(['2025-01-05', 'Junk Row', -1.00, '', '', 'Unknown Vendor', '']); // To trigger Checker

    // --- CC SHEET ---
    const cc = wb.addWorksheet('CC Amex');
    cc.addRow(['Date', 'Description', 'Amount', 'Category', 'Sub-Category', 'Vendor', 'Customer']); // Amount is POSITIVE (Flip=Yes)
    cc.addRow(['2025-01-05', 'Supplies', 150, 'Office Exp', 'Supplies', 'Staples', '']);
    cc.addRow(['2025-01-06', 'Legal Consult', 1200, 'Professional Fees', 'Legal', 'Lawyer 1099', '']); // 1099 Candidate (>600)

    // --- LEDGER SHEET ---
    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Sub-Category', 'Debit', 'Credit', 'Vendor', 'Customer']);
    // Adjusting Entries
    // Reclass: Move 100 from Supplies to Rent?
    // Dr Rent 100, Cr Supplies 100
    ledger.addRow(['2025-01-31', 'Reclass', 'Office Exp', 'Rent', 100, 0, 'Landlord', '']);
    ledger.addRow(['2025-01-31', 'Reclass', 'Office Exp', 'Supplies', 0, 100, 'Staples', '']);

    // Add Interest Income (Bank Interest) via Ledger
    // Dr Bank (Asset) 10, Cr Interest Income (Sales/General?) 10
    // REMOVED 'Bank Interest' vendor from Asset side to avoid usage conflict (Asset vs P&L)
    ledger.addRow(['2025-01-31', 'Interest', 'Cash', '', 10, 0, '', '']); // Debit = Increase Asset
    ledger.addRow(['2025-01-31', 'Interest', 'Sales', 'Consulting', 0, 10, 'Bank Interest', '']); // Credit = Income

    await wb.xlsx.writeFile(TEST_FILE);
    console.log(`Created Test Data: ${TEST_FILE}`);
}

async function runTests() {
    await createTestData();

    let failureCount = 0;

    const check = (name, output, includes = [], excludes = []) => {
        const miss = includes.filter(s => !output.includes(s));
        const bad = excludes.filter(s => output.includes(s));

        // Check for runtime errors
        const hasError = output.includes('[ERROR]');

        if (miss.length || bad.length || hasError) {
            console.error(`❌ [FAIL] ${name}`);
            if (miss.length) console.error(`   Missing: ${miss.join(', ')}`);
            if (bad.length) console.error(`   Found Illegal: ${bad.join(', ')}`);
            if (hasError) {
                console.error(`   Runtime Error Detected!`);
                // Extract error line for debugging
                const errorLines = output.split('\n').filter(line => line.includes('[ERROR]'));
                errorLines.forEach(line => console.error(`   ${line.trim()}`));
            }
            console.error(`   Actual Output (Truncated 500 chars): ${output.substring(0, 500).replace(/\n/g, ' ')}...`);
            failureCount++;
        } else {
            console.log(`✅ [PASS] ${name}`);
        }
    };

    // 1. Basic PL
    const plOut = run(`node report.js "${TEST_FILE}" --pl`, true);
    check('P&L Standard', plOut, ['PROFIT & LOSS', 'Sales', 'Office Exp', 'NET INCOME']);

    // 2. PL Sub (Check New Header Format)
    const plSubOut = run(`node report.js "${TEST_FILE}" --pl-sub`, true);
    // Look for new header strings "Additions" "Subtractions" "Net"
    // Look for Sheet Columns "Bank ABC" "CC Amex" "Ledger"
    check('P&L Sub-Report (Detailed)', plSubOut,
        ['Additions', 'Subtractions', 'Net', 'Bank', 'Amex', 'Ledger']);

    // 3. BS Sub
    const bsSubOut = run(`node report.js "${TEST_FILE}" --bs-sub`, true);
    check('BS Sub-Report (Detailed)', bsSubOut,
        ['BALANCE SHEET', 'Additions', 'Subtractions', 'Net', 'Bank']);

    // 4. Customer Sub
    const custSubOut = run(`node report.js "${TEST_FILE}" --customer-sub`, true);
    check('Customer Sub-Report', custSubOut, ['CUSTOMER INCOME', 'Client A', 'Bank']);

    // 5. Vendor Sub
    const vendSubOut = run(`node report.js "${TEST_FILE}" --vendor-sub`, true);
    check('Vendor Sub-Report', vendSubOut, ['VENDOR SPENDING', 'Staples', 'Landlord', 'Additions', 'Subtractions']);

    // 6. Checker
    const checkerOut = run(`node report.js "${TEST_FILE}" --checker`, true);
    check('Checker', checkerOut, ['DATA INTEGRITY ISSUES', 'Junk Row', 'MISSING CATEGORY']);

    // 7. Details
    const detailOut = run(`node report.js "${TEST_FILE}" --details "Sales"`, true);
    check('Details Flag', detailOut, ['DETAILS: "sales"', 'Client Payment', '5000.00']);

    // 8. 1099 Generation within Excel Report
    const reportFile = 'tests/report_comprehensive_test_data.xlsx';
    try { fs.unlinkSync(reportFile); } catch (e) { }
    const out1099 = run(`node report.js "${TEST_FILE}" --1099 --save`, true);
    console.log(out1099); // Print full output for debugging

    if (fs.existsSync(reportFile)) {
        const testWb = new ExcelJS.Workbook();
        await testWb.xlsx.readFile(reportFile);
        const sheet1099 = testWb.getWorksheet('1099 Data');

        if (sheet1099) {
            console.log('✅ [PASS] 1099 Data Sheet Found in Excel');
            let foundLawyer = false;
            let lawyerAmount = 0;

            sheet1099.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;
                const name = row.getCell(1).value;
                const amtVal = row.getCell(13).value;
                console.log(`   [DEBUG 1099 Row] Name: "${name}", Amount: ${amtVal}`);
                if (name && (name.toString().includes('Lawyer 1099') || name.toString().includes('Law Firm LLC'))) {
                    foundLawyer = true;
                    lawyerAmount = parseFloat(amtVal);
                }
            });

            if (foundLawyer && Math.abs(lawyerAmount - 2000) < 0.01) {
                console.log('✅ [PASS] 1099 Data Content Valid in Excel');
            } else {
                console.error('❌ [FAIL] 1099 Data Content Invalid in Excel');
                console.error(`   Expected "Law Firm LLC" (or Lawyer 1099) with amount 2000. Found: ${foundLawyer}, Amount: ${lawyerAmount}`);
                failureCount++;
            }
        } else {
            console.error('❌ [FAIL] 1099 Data Sheet Missing in Excel');
            failureCount++;
        }
    } else {
        console.error('❌ [FAIL] Excel Report Not Created');
        failureCount++;
    }

    if (failureCount > 0) process.exit(1);
    console.log('\nALL TESTS PASSED.');
}

runTests();

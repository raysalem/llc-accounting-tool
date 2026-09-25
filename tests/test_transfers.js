// Automates the three transfer scenarios in docs/testing.md:
//   1. Credit card payment is not double counted
//   2. Refunds reduce expense and liability; payments are excluded from linkage
//   3. A transfer category used on the wrong sheet raises a warning
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-transfers-'));
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

// Returns the number printed after "label :" on the first matching line,
// e.g. "Checking            :            0.00 ->          900.00" -> 900.
function valueFor(output, label) {
    const line = output.split('\n').find(l => new RegExp(`^\\s*${label}\\s*:`).test(l));
    if (!line) return null;
    const nums = line.match(/-?[\d,]+\.\d{2}/g);
    return nums ? parseFloat(nums[nums.length - 1].replace(/,/g, '')) : null;
}

function runReport(file, flags) {
    // Capture stdout and stderr: report.js prints warnings with console.warn.
    const r = spawnSync(process.execPath, ['report.js', file, '--year=2025', ...flags.split(' ')], { cwd: ROOT, encoding: 'utf8' });
    return { code: r.status, output: r.stdout + r.stderr };
}

// Bank rows and Amex rows are [date, description, amount, category].
async function buildWorkbook(file, bankRows, amexRows) {
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.addRow([
        'Category', 'Sub-Category', 'Account Type', 'Report', 'Transfer Account',
        'Vendors', 'Customers',
        'Sheet Name', 'Type', 'Flip', 'Offset', 'Short Name', 'Link Asset'
    ]);
    const categories = [
        ['Sales', '', 'Income', 'P&L'],
        ['Office Supplies', '', 'Expense', 'P&L'],
        ['Checking', '', 'Asset', 'Balance Sheet'],
        ['AX CC', '', 'Liability', 'Balance Sheet'],
        ['Owner Equity', '', 'Equity', 'Balance Sheet'],
        ['Transfer AX CC', '', 'Transfer', 'Transfer', 'AX CC'],
    ];
    categories.forEach((c, i) => c.forEach((v, j) => { setup.getRow(i + 2).getCell(j + 1).value = v; }));

    // Sheet config (columns H-M)
    const sheets = [
        ['Bank', 'Bank', 'No', 1, 'Bank', 'Checking'],
        ['Amex', 'CC', 'Yes', 1, 'Amex', 'AX CC'],
    ];
    sheets.forEach((s, i) => s.forEach((v, j) => { setup.getRow(i + 2).getCell(8 + j).value = v; }));

    const header = ['Date', 'Description', 'Amount', 'Category'];
    const bank = wb.addWorksheet('Bank');
    bank.addRow(header);
    bankRows.forEach(r => bank.addRow(r));
    const amex = wb.addWorksheet('Amex');
    amex.addRow(header);
    amexRows.forEach(r => amex.addRow(r));

    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);

    await wb.xlsx.writeFile(file);
}

async function main() {
    // --- Scenarios 1 & 2 ---
    // Bank: +1000 sales, -100 payment to the Amex (categorized to the liability).
    // Amex (statement shows charges as positive, so Flip = Yes):
    //   +300 office supplies charge, -100 payment received (Transfer), -50 refund.
    // Expected:
    //   Checking = 1000 - 100 = 900
    //   AX CC    = 300 - 50 (refund) - 100 (bank payment only) = 150
    //   Office Supplies expense = 300 - 50 = 250; net income = 750
    //   A = L + E: 900 = 150 + 750
    const good = path.join(TMP, 'transfers.xlsx');
    await buildWorkbook(good,
        [
            [new Date('2025-02-01'), 'Client payment', 1000, 'Sales'],
            [new Date('2025-02-15'), 'Amex autopay', -100, 'AX CC'],
        ],
        [
            [new Date('2025-02-03'), 'Office supplies', 300, 'Office Supplies'],
            [new Date('2025-02-15'), 'Payment received', -100, 'Transfer AX CC'],
            [new Date('2025-02-20'), 'Supplies refund', -50, 'Office Supplies'],
        ]);

    const r = runReport(good, '--pl --bs');
    const out = r.output;
    check('Clean transfer books exit 0', r.code === 0, `exit ${r.code}`);
    check('Checking = 900.00', valueFor(out, 'Checking') === 900, `got ${valueFor(out, 'Checking')}`);
    check('AX CC liability = 150.00 (payment counted once)', valueFor(out, 'AX CC') === 150, `got ${valueFor(out, 'AX CC')}`);
    check('Office Supplies = -250.00 (refund reduces expense)', valueFor(out, 'Office Supplies') === -250, `got ${valueFor(out, 'Office Supplies')}`);
    check('Net income = 750.00', out.includes('NET INCOME: 750.00'));
    check('Balance sheet balances', out.includes('[OK] (A = L + E)'));
    check('Transfer category not on the P&L', !/^\s*Transfer AX CC\s*:/m.test(out.split('--- BALANCE SHEET')[0]));

    // --- Scenario 3: transfer category used on the wrong sheet ---
    const wrong = path.join(TMP, 'transfers_wrong_sheet.xlsx');
    await buildWorkbook(wrong,
        [
            [new Date('2025-02-01'), 'Client payment', 1000, 'Sales'],
            [new Date('2025-02-15'), 'Amex autopay', -100, 'Transfer AX CC'],
        ],
        [
            [new Date('2025-02-03'), 'Office supplies', 300, 'Office Supplies'],
        ]);
    const w = runReport(wrong, '--bs');
    check('Transfer on wrong sheet raises CRITICAL WARNING',
        /CRITICAL WARNING.*"Transfer AX CC" expects Transfer Account "AX CC".*Sheet "Bank"/.test(w.output));

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll transfer tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

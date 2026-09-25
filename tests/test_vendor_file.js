// --vendor-file, --ignore-vendors, --all and --debug.
// The workbook's NEC contractor has no tax details in Setup; a separate
// vendor.csv supplies them.
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-vendorfile-'));
const BOOKS_DIR = path.join(TMP, 'books');
const VENDOR_DIR = path.join(TMP, 'vendors');
const FILE = path.join(BOOKS_DIR, 'books.xlsx');
const VENDOR_CSV = path.join(VENDOR_DIR, 'contractors.csv');
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

async function build() {
    fs.mkdirSync(BOOKS_DIR);
    fs.mkdirSync(VENDOR_DIR);
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.addRow(['Category', 'Sub-Category', 'Account Type', 'Report', 'Vendors', '1099 Type',
        'Sheet Name', 'Sheet Type', 'Flip', 'Offset', 'Link Asset', 'Company Info', 'Value']);
    setup.addRow(['Contract Labor', '', 'Expense', 'P&L', 'Jane Contractor', 'NEC',
        'Bank', 'Bank', 'No', 1, 'Checking', 'Company Name', 'Example LLC']);
    setup.addRow(['Checking', '', 'Asset', 'Balance Sheet', '', '', '', '', '', '', '', 'TIN', '12-3456789']);
    setup.addRow(['Owner Equity', '', 'Equity', 'Balance Sheet']);

    const bank = wb.addWorksheet('Bank');
    bank.addRow(['Date', 'Description', 'Amount', 'Category', 'Vendor']);
    bank.addRow([new Date('2025-02-01'), 'Owner deposit', 5000, 'Owner Equity', '']);
    bank.addRow([new Date('2025-03-01'), 'Website build', -1200, 'Contract Labor', 'Jane Contractor']);

    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);
    await wb.xlsx.writeFile(FILE);

    fs.writeFileSync(VENDOR_CSV, [
        'Vendor,Business Name,TIN,Address,City,State,Zip',
        'Jane Contractor,Jane Builds LLC,98-7654321,1 Main St,Springfield,IL,62701',
    ].join('\n'));
}

function run(...flags) {
    const r = spawnSync(process.execPath, ['report.js', FILE, '--year=2025', ...flags], { cwd: ROOT, encoding: 'utf8' });
    return { code: r.status, output: r.stdout + r.stderr };
}

async function main() {
    await build();

    const without = run('--1099');
    check('Without a vendor file, missing 1099 details are flagged', without.output.includes('Incomplete Data'));
    check('Without a vendor file, the run exits 1', without.code === 1, `exit ${without.code}`);

    const withFile = run('--1099', '--vendor-file', VENDOR_CSV);
    check('--vendor-file supplies the TIN and address', !withFile.output.includes('Incomplete Data'));
    check('--vendor-file run exits 0', withFile.code === 0, `exit ${withFile.code}`);
    check('Contractor is listed for 1099-NEC', withFile.output.includes('Jane Contractor'));

    // A vendor.csv next to the workbook is found automatically; --ignore-vendors skips it.
    fs.copyFileSync(VENDOR_CSV, path.join(BOOKS_DIR, 'vendor.csv'));
    const auto = run('--1099');
    check('vendor.csv next to the workbook is loaded automatically', !auto.output.includes('Incomplete Data'));
    const ignored = run('--1099', '--ignore-vendors');
    check('--ignore-vendors skips vendor.csv', ignored.output.includes('Incomplete Data'));

    const all = run('--all', '--vendor-file', VENDOR_CSV);
    for (const section of ['PROFIT & LOSS', 'BALANCE SHEET', 'VENDOR SPENDING', '1099 PREPARATION']) {
        check(`--all prints ${section}`, all.output.includes(section));
    }
    check('--all run exits 0', all.code === 0, `exit ${all.code}`);

    const debug = run('--pl', '--debug', '--vendor-file', VENDOR_CSV);
    check('--debug prints debug output', debug.output.includes('[DEBUG]'));
    check('--debug does not change results', debug.output.includes('NET INCOME: -1,200.00'));

    fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll vendor-file and flag tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

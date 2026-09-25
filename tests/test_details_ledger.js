// --details must include Ledger rows, and must not change any totals.
// (Previously a typo made every Ledger row throw when --details was used, which
// silently dropped the row's vendor/customer totals.)
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-details-'));
const FILE = path.join(TMP, 'details_ledger.xlsx');
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

async function build() {
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.addRow(['Category', 'Sub-Category', 'Account Type', 'Report', 'Vendors', 'Sheet Name', 'Type', 'Flip', 'Offset', 'Link Asset']);
    setup.addRow(['Office', '', 'Expense', 'P&L', 'Staples', 'Bank', 'Bank', 'No', 1, 'Checking']);
    setup.addRow(['Checking', '', 'Asset', 'Balance Sheet']);
    setup.addRow(['Owner Equity', '', 'Equity', 'Balance Sheet']);

    const bank = wb.addWorksheet('Bank');
    bank.addRow(['Date', 'Description', 'Amount', 'Category', 'Vendor']);
    bank.addRow([new Date('2025-04-01'), 'Owner deposit', 500, 'Owner Equity', '']);
    bank.addRow([new Date('2025-04-02'), 'Paper', -50, 'Office', 'Staples']);

    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit', 'Vendor']);
    ledger.addRow([new Date('2025-04-30'), 'Owner paid Staples invoice', 'Office', 100, 0, 'Staples']);
    ledger.addRow([new Date('2025-04-30'), 'Owner paid Staples invoice', 'Owner Equity', 0, 100, '']);
    await wb.xlsx.writeFile(FILE);
}

function run(...flags) {
    const r = spawnSync(process.execPath, ['report.js', FILE, '--year=2025', ...flags], { cwd: ROOT, encoding: 'utf8' });
    return r.stdout + r.stderr;
}

function staplesTotal(output) {
    const line = output.split('\n').find(l => l.startsWith('Staples'.padEnd(30)));
    return line ? parseFloat(line.match(/-?[\d,]+\.\d{2}/)[0].replace(/,/g, '')) : null;
}

async function main() {
    await build();

    const byCategory = run('--details', 'office');
    check('--details by category lists the bank row', byCategory.includes('Paper'));
    check('--details by category lists the ledger row', byCategory.includes('Owner paid Staples invoice'));

    const byVendor = run('--details', 'staples');
    check('--details by vendor lists the ledger row', byVendor.includes('Owner paid Staples invoice'));

    const plain = staplesTotal(run('--vendor'));
    const withDetails = staplesTotal(run('--vendor', '--details', 'office'));
    check('Staples total is 150.00 (50 bank + 100 ledger)', plain === 150, `got ${plain}`);
    check('--details does not change vendor totals', withDetails === plain, `got ${withDetails} vs ${plain}`);

    fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll --details ledger tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

// An unbalanced ledger (debits != credits) is an error: exit code 1. Without --checker the reports
// still print and the imbalance is repeated at the end; with --checker the run stops before them.
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-ledger-'));

async function build(file, ledgerRows) {
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.getRow(1).values = ['Category', 'Sub-Category', 'Type', 'Report'];
    setup.getRow(2).values = ['Office', '', 'Expense', 'P&L'];
    setup.getRow(3).values = ['Owner Equity', '', 'Equity', 'Balance Sheet'];
    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);
    ledgerRows.forEach(r => ledger.addRow(r));
    await wb.xlsx.writeFile(file);
}

function run(file, extra = []) {
    const r = spawnSync(process.execPath, ['report.js', file, '--year=2025', '--pl', ...extra], { cwd: ROOT, encoding: 'utf8' });
    return { code: r.status, output: r.stdout + r.stderr };
}

async function main() {
    let failures = 0;
    const check = (label, ok) => {
        console.log(`${ok ? '✅ [PASS]' : '❌ [FAIL]'} ${label}`);
        if (!ok) failures++;
    };

    const unbalanced = path.join(TMP, 'unbalanced.xlsx');
    await build(unbalanced, [[new Date('2025-05-01'), 'One-sided entry', 'Office', 100, 0]]);
    const u = run(unbalanced);
    check('Unbalanced ledger exits 1', u.code === 1);
    check('Unbalanced ledger reports the difference', u.output.includes('Ledger is UNBALANCED') && u.output.includes('100.00'));
    check('Without --checker: the P&L still prints', u.output.includes('PROFIT & LOSS'));
    check('Without --checker: the imbalance is repeated in the final summary', u.output.includes('Ledger: UNBALANCED by 100.00'));
    const strict = run(unbalanced, ['--checker']);
    check('With --checker: exits 1', strict.code === 1);
    check('With --checker: stops before the reports', strict.output.includes('Ledger is UNBALANCED') && !strict.output.includes('PROFIT & LOSS'));

    const balanced = path.join(TMP, 'balanced.xlsx');
    await build(balanced, [
        [new Date('2025-05-01'), 'Owner paid expense', 'Office', 100, 0],
        [new Date('2025-05-01'), 'Owner paid expense', 'Owner Equity', 0, 100],
    ]);
    const b = run(balanced);
    check('Balanced ledger does not report UNBALANCED', !b.output.includes('UNBALANCED'));

    fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) process.exit(1);
    console.log('\nAll ledger balance tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

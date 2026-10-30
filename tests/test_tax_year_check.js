// Tax year check: every row of a book must be dated in the tax year. A row from another year
// (e.g. a ledger date typed as 2016 instead of 2026) is listed with its sheet, row, date and
// amount, the reports still print, and with --book the run exits 1.
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-taxyear-'));
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

function run(script, args) {
    const r = spawnSync(process.execPath, [script, ...args], { cwd: ROOT, encoding: 'utf8' });
    return { code: r.status, out: r.stdout + r.stderr };
}

async function addLedgerRows(file, rows) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    const ws = wb.getWorksheet('Ledger');
    // Ledger header is on row 3 (example book); columns: Entry ID, Date, Description, Category, Sub-Category, Debit, Credit
    rows.forEach((r, i) => { ws.getRow(4 + i).values = r; });
    await wb.xlsx.writeFile(file);
}

async function main() {
    fs.copyFileSync(path.join(ROOT, 'examples', 'example.book.toml'), path.join(TMP, 'example.book.toml'));
    fs.copyFileSync(path.join(ROOT, 'examples', 'example.private.toml'), path.join(TMP, 'example.private.toml'));
    const book = path.join(TMP, 'example.book.toml');
    const gen = run('generate_excel.js', ['--book', book]);
    check('Workbook generated', gen.code === 0, gen.out);
    const wbFile = path.join(TMP, 'Example_2026.xlsx');

    // One balanced entry in 2026, then the same kind of entry with the year typed as 2016.
    await addLedgerRows(wbFile, [
        [1, new Date(Date.UTC(2026, 9, 19)), 'Paper', 'Office Supplies', 'Paper', 100, null],
        [1, new Date(Date.UTC(2026, 9, 19)), 'Paper', 'Owner Investment', null, null, 100],
        [2, new Date(Date.UTC(2016, 9, 1)), 'Software', 'Office Supplies', 'Software', 60, null],
        [2, new Date(Date.UTC(2016, 9, 1)), 'Software', 'Owner Investment', null, null, 60],
    ]);

    const r = run('report.js', ['--book', book, '--pl']);
    check('Exit 1: a book row outside the tax year is an error', r.code === 1, `exit ${r.code}`);
    check('Block names how many rows are outside', r.out.includes('TAX YEAR CHECK: 2 row(s) NOT dated 2026'), r.out);
    check('Each row is listed with sheet, row and date', r.out.includes('"Ledger" row 6: 2016-10-01') && r.out.includes('"Ledger" row 7: 2016-10-01'), r.out);
    check('Summary line counts them', r.out.includes('Tax year check: 2 of 4 rows are NOT dated 2026'), r.out);
    check('The reports still print', r.out.includes('PROFIT & LOSS'));
    check('Says how to fix it', r.out.includes('Fix these dates'));

    // Fix the year: the check passes.
    await addLedgerRows(wbFile, [
        [1, new Date(Date.UTC(2026, 9, 19)), 'Paper', 'Office Supplies', 'Paper', 100, null],
        [1, new Date(Date.UTC(2026, 9, 19)), 'Paper', 'Owner Investment', null, null, 100],
        [2, new Date(Date.UTC(2026, 9, 1)), 'Software', 'Office Supplies', 'Software', 60, null],
        [2, new Date(Date.UTC(2026, 9, 1)), 'Software', 'Owner Investment', null, null, 60],
    ]);
    const ok = run('report.js', ['--book', book, '--pl']);
    check('All rows in the year: summary says so', ok.out.includes('Tax year check: all 4 rows are dated 2026'), ok.out);
    check('All rows in the year: no TAX YEAR CHECK block', !ok.out.includes('TAX YEAR CHECK:'));

    fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll tax year check tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

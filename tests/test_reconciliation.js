// Every bank/card sheet must carry an End Balance (statement ending balance) and
// the report prints a stamp of what it was made from. Covers: all accounts
// reconciled, a missing End Balance, an explicit 0, a mismatch, the mixed-case
// summary counts, and the stamp in the console and in the --save Report Info sheet.
const { spawnSync } = require('child_process');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-recon-'));
const HASH_HEX_LENGTH = 12;
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

// Three accounts in 2025:
//   Checking  +1000 - 100 = 900
//   Savings   +200
//   Amex      (liability, Flip = Yes) 50 + 30 = 80
// endBalances = [checking, savings, amex]; null leaves the cell empty.
async function buildWorkbook(file, endBalances) {
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.addRow(['Category', 'Sub-Category', 'Account Type', 'Report', 'Vendors', 'Customers',
        'Sheet Name', 'Type', 'Flip', 'Offset', 'Short Name', 'Link Asset', 'End Balance']);
    const categories = [
        ['Sales', '', 'Income', 'P&L'],
        ['Office Supplies', '', 'Expense', 'P&L'],
        ['Checking', '', 'Asset', 'Balance Sheet'],
        ['Savings', '', 'Asset', 'Balance Sheet'],
        ['Amex Card', '', 'Liability', 'Balance Sheet'],
    ];
    categories.forEach((c, i) => c.forEach((v, j) => { setup.getRow(i + 2).getCell(j + 1).value = v; }));
    const sheets = [
        ['Bank', 'Bank', 'No', 1, 'Bank', 'Checking'],
        ['Savings', 'Bank', 'No', 1, 'Savings', 'Savings'],
        ['Amex', 'CC', 'Yes', 1, 'Amex', 'Amex Card'],
    ];
    sheets.forEach((s, i) => {
        s.forEach((v, j) => { setup.getRow(i + 2).getCell(7 + j).value = v; });
        if (endBalances[i] !== null) setup.getRow(i + 2).getCell(7 + s.length).value = endBalances[i];
    });

    const header = ['Date', 'Description', 'Amount', 'Category'];
    const bank = wb.addWorksheet('Bank');
    bank.addRow(header);
    bank.addRow([new Date('2025-03-01'), 'Client payment', 1000, 'Sales']);
    bank.addRow([new Date('2025-03-02'), 'Supplies', -100, 'Office Supplies']);
    const savings = wb.addWorksheet('Savings');
    savings.addRow(header);
    savings.addRow([new Date('2025-03-03'), 'Client payment', 200, 'Sales']);
    const amex = wb.addWorksheet('Amex');
    amex.addRow(header);
    amex.addRow([new Date('2025-03-04'), 'Supplies', 50, 'Office Supplies']);
    amex.addRow([new Date('2025-03-05'), 'More supplies', 30, 'Office Supplies']);

    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);
    await wb.xlsx.writeFile(file);
}

function run(file, ...flags) {
    const r = spawnSync(process.execPath, ['report.js', file, '--year=2025', ...flags], { cwd: ROOT, encoding: 'utf8' });
    return { code: r.status, output: r.stdout + r.stderr };
}

async function main() {
    // --- All reconciled ---
    const good = path.join(TMP, 'good.xlsx');
    await buildWorkbook(good, [900, 200, 80]);
    const g = run(good, '--pl');
    check('All accounts reconciled: exit 0', g.code === 0, `exit ${g.code}\n${g.output}`);
    check('All accounts reconciled: summary "3 of 3"', g.output.includes('Reconciliation: 3 of 3 accounts match their End Balance'));
    check('All accounts reconciled: no UNRECONCILED or MISMATCH', !g.output.includes('UNRECONCILED') && !g.output.includes('MISMATCH'));

    // --- One End Balance missing ---
    const missing = path.join(TMP, 'missing.xlsx');
    await buildWorkbook(missing, [900, null, 80]);
    const m = run(missing, '--pl');
    check('Missing End Balance: exit 1', m.code === 1, `exit ${m.code}`);
    check('Missing End Balance: UNRECONCILED names the sheet',
        /UNRECONCILED: "Savings" has no End Balance in Setup \(SheetInfo\)/.test(m.output), m.output);
    check('Missing End Balance: other sheets are not named', !/UNRECONCILED: "(Bank|Amex)"/.test(m.output));
    check('Missing End Balance: summary "2 of 3 ... 1 missing End Balance"',
        m.output.includes('Reconciliation: 2 of 3 accounts match (1 missing End Balance)'));

    // --- Explicit 0 that matches is accepted ---
    const zero = path.join(TMP, 'zero.xlsx');
    await buildWorkbook(zero, [900, 200, 80]);
    const zwb = new ExcelJS.Workbook();
    await zwb.xlsx.readFile(zero);
    // Savings has no activity in 2026, so End Balance 0 is the right statement balance for --year=2026.
    zwb.getWorksheet('Setup').getRow(3).getCell(13).value = 0;
    await zwb.xlsx.writeFile(zero);
    const z = spawnSync(process.execPath, ['report.js', zero, '--year=2026', '--pl'], { cwd: ROOT, encoding: 'utf8' });
    const zOut = z.stdout + z.stderr;
    check('Explicit End Balance 0 is accepted (not reported missing)', !/UNRECONCILED: "Savings"/.test(zOut), zOut);
    check('Explicit End Balance 0 matches a zero balance', !/MISMATCH: "Savings"/.test(zOut));
    // The other two accounts have End Balances for 2025, so in 2026 they mismatch: 1 of 3 accounts match.
    check('Explicit 0 counts as matching in the summary', zOut.includes('Reconciliation: 1 of 3 accounts match (2 mismatches)'), zOut);

    // --- Mismatch still fails ---
    const bad = path.join(TMP, 'mismatch.xlsx');
    await buildWorkbook(bad, [901, 200, 80]);
    const b = run(bad, '--pl');
    check('Mismatch: exit 1', b.code === 1, `exit ${b.code}`);
    check('Mismatch: END BALANCE MISMATCH names the sheet and both numbers',
        b.output.includes('END BALANCE MISMATCH: "Bank" (Calc: 900.00 vs Exp: 901.00)'), b.output);
    check('Mismatch: summary "2 of 3 ... 1 mismatch"', b.output.includes('Reconciliation: 2 of 3 accounts match (1 mismatch)'));

    // --- Mixed: one mismatch and one missing ---
    const mixed = path.join(TMP, 'mixed.xlsx');
    await buildWorkbook(mixed, [901, null, 80]);
    const x = run(mixed, '--pl');
    check('Mixed: exit 1', x.code === 1);
    check('Mixed: both problems are listed together', /UNRECONCILED: "Savings"/.test(x.output) && /MISMATCH: "Bank"/.test(x.output));
    check('Mixed: summary counts', x.output.includes('Reconciliation: 1 of 3 accounts match (1 mismatch, 1 missing End Balance)'), x.output);

    // --- Report stamp on the console ---
    const expectedHash = crypto.createHash('sha256').update(fs.readFileSync(good)).digest('hex').slice(0, HASH_HEX_LENGTH);
    const size = fs.statSync(good).size;
    const stampLine = g.output.split('\n').find(l => l.includes('Input: Workbook'));
    const hashMatch = stampLine && stampLine.match(/sha256 ([0-9a-f]+)/);
    check('Stamp: workbook hash is 12 hex digits', hashMatch && hashMatch[1].length === HASH_HEX_LENGTH, stampLine);
    check('Stamp: workbook hash matches SHA-256 computed by the test', hashMatch && hashMatch[1] === expectedHash, `${hashMatch && hashMatch[1]} vs ${expectedHash}`);
    check('Stamp: workbook size is printed', stampLine && stampLine.includes(`${size.toLocaleString('en-US')} bytes`), stampLine);
    check('Stamp: tax year is printed', /Tax Year\s*: 2025/.test(g.output));
    check('Stamp: flags are printed as given', /Flags\s*:.*good\.xlsx --year=2025 --pl/.test(g.output));
    check('Stamp: git commit field is present (a short sha, with +dirty if changed, or unknown)',
        /Git Commit\s*: ([0-9a-f]{7,40}(\+dirty)?|unknown)\s*$/m.test(g.output), g.output);
    check('Stamp: run time is ISO with a time zone', /Run Time\s*: \d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2}/.test(g.output));
    check('Stamp: tool version is printed', /Tool Version\s*: \d+\.\d+\.\d+/.test(g.output));

    // --- Stamp in the --save Report Info sheet ---
    // --save exits 1 whenever any warning was printed, so read the saved file instead of the exit code.
    run(good, '--pl', '--save');
    const saved = new ExcelJS.Workbook();
    await saved.xlsx.readFile(path.join(TMP, 'report_good.xlsx'));
    const info = saved.getWorksheet('Report Info');
    check('--save: Report Info sheet exists', Boolean(info));
    const rows = {};
    if (info) info.eachRow(row => { rows[String(row.getCell(1).value)] = String(row.getCell(2).value); });
    check('--save: Report Info has Tool Version', /^\d+\.\d+\.\d+$/.test(rows['Tool Version'] || ''), JSON.stringify(rows));
    check('--save: Report Info has Git Commit', /^([0-9a-f]{7,40}(\+dirty)?|unknown)$/.test(rows['Git Commit'] || ''), rows['Git Commit']);
    check('--save: Report Info has Run Time', /^\d{4}-\d{2}-\d{2}T/.test(rows['Run Time'] || ''));
    check('--save: Report Info has Flags', (rows.Flags || '').endsWith('--year=2025 --pl --save'), rows.Flags);
    check('--save: Report Info has Tax Year 2025', rows['Tax Year'] === '2025');
    check('--save: Report Info has the workbook hash and size',
        (rows['Input: Workbook'] || '').includes(`good.xlsx  sha256 ${expectedHash}  ${size.toLocaleString('en-US')} bytes`), rows['Input: Workbook']);

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll reconciliation tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

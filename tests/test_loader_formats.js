// End-to-end check of load_transactions.js for bank-style amount formats and
// dates around the year boundary. Runs the loader under several time zones to
// make sure a row never moves to a different calendar day or tax year.
const { execFileSync, spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-loader-'));
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

const CSV = [
    'Date,Name,Memo,Amount',
    '12/31/2025,Year end deposit,,"$1,234.56"',
    '2025-12-31,Year end fee,,(50.00)',
    '01/01/2026,New year charge,,"1,000.00-"',
    '2026-01-01,New year deposit,,25',
].join('\n');

const EXPECTED = [
    { desc: 'Year end deposit', date: '2025-12-31', amount: 1234.56 },
    { desc: 'Year end fee', date: '2025-12-31', amount: -50 },
    { desc: 'New year charge', date: '2026-01-01', amount: -1000 },
    { desc: 'New year deposit', date: '2026-01-01', amount: 25 },
];

async function loadAndRead(tz) {
    const book = path.join(TMP, `books_${tz.replace(/\W/g, '_')}.xlsx`);
    const csv = path.join(TMP, 'statement.csv');
    fs.writeFileSync(csv, CSV);
    const env = { ...process.env, TZ: tz };
    execFileSync(process.execPath, ['generate_excel.js', book], { cwd: ROOT, env, stdio: 'pipe' });
    execFileSync(process.execPath, ['load_transactions.js', csv, 'bank', book, '--clear'], { cwd: ROOT, env, stdio: 'pipe' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(book);
    const rows = [];
    wb.getWorksheet('Bank Transactions').eachRow((row, r) => {
        if (r <= 3) return; // TOTAL, SUBTOTAL, header
        const [date, desc, amount] = [row.getCell(1).value, row.getCell(2).value, row.getCell(3).value];
        if (date instanceof Date) rows.push({ desc, date: date.toISOString().slice(0, 10), amount });
    });
    return { book, rows };
}

async function main() {
    for (const tz of ['UTC', 'America/Los_Angeles', 'Asia/Tokyo']) {
        const { book, rows } = await loadAndRead(tz);
        for (const exp of EXPECTED) {
            const got = rows.find(r => r.desc === exp.desc);
            check(`[${tz}] ${exp.desc}: date ${exp.date}`, got && got.date === exp.date, got ? `got ${got.date}` : 'row missing');
            check(`[${tz}] ${exp.desc}: amount ${exp.amount}`, got && got.amount === exp.amount, got ? `got ${got.amount}` : 'row missing');
        }

        // Each tax year sees only its own two rows.
        const r25 = spawnSync(process.execPath, ['report.js', book, '--year=2025', '--pl'], { cwd: ROOT, encoding: 'utf8', env: { ...process.env, TZ: tz } });
        check(`[${tz}] 2025 report skips the two 2026 rows`, r25.stdout.includes('Skipped 2 transaction row(s) dated outside 2025'));
        const r26 = spawnSync(process.execPath, ['report.js', book, '--year=2026', '--pl'], { cwd: ROOT, encoding: 'utf8', env: { ...process.env, TZ: tz } });
        check(`[${tz}] 2026 report skips the two 2025 rows`, r26.stdout.includes('Skipped 2 transaction row(s) dated outside 2026'));
    }

    fs.rmSync(TMP, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll loader format tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

// When Setup has no sheet configuration, report.js falls back to the default
// "Bank Transactions" / "Credit Card Transactions" sheets. Those sheets have their
// header on row 3 when filled by load_transactions.js, and the fallback must find it.
const { execFileSync, spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-fallback-'));
const FILE = path.join(TMP, 'books.xlsx');

async function main() {
    execFileSync(process.execPath, ['generate_excel.js', FILE], { cwd: ROOT, stdio: 'pipe' });
    execFileSync(process.execPath, ['load_transactions.js', path.join(ROOT, 'tests/example_bank.csv'), 'bank', FILE, '--clear'], { cwd: ROOT, stdio: 'pipe' });

    // Remove the sheet configuration (Setup columns I-L) and categorize the rows.
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(FILE);
    const setup = wb.getWorksheet('Setup');
    for (let r = 1; r <= 10; r++) for (let c = 9; c <= 12; c++) setup.getRow(r).getCell(c).value = null;
    const bank = wb.getWorksheet('Bank Transactions');
    bank.eachRow((row, r) => {
        if (r <= 3) return;
        const desc = String(row.getCell(2).value || '');
        row.getCell(4).value = desc.includes('Rent') ? 'Rent' : 'Sales';
    });
    await wb.xlsx.writeFile(FILE);

    const r = spawnSync(process.execPath, ['report.js', FILE, '--year=2025', '--pl', '--checker'], { cwd: ROOT, encoding: 'utf8' });
    const out = r.stdout + r.stderr;
    let failures = 0;
    const check = (label, ok) => { console.log(`${ok ? '✅ [PASS]' : '❌ [FAIL]'} ${label}`); if (!ok) failures++; };
    check('Fallback sheet configuration is used', out.includes('No sheet configurations found in Setup'));
    check('Header row 3 is found (no missing Date column)', !out.includes("'Date' column not found"));
    check('Sales = 7,500.00', /Sales\s*:\s*7,500\.00/.test(out));
    check('Net income = 6,000.00', out.includes('NET INCOME: 6,000.00'));

    fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) process.exit(1);
    console.log('\nFallback sheet test passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

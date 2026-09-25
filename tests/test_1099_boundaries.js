// 1099-NEC "Required" flag at the exact threshold, just under it, after a refund,
// and across the 2025 ($600) / 2026 ($2,000) threshold change.
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-1099-'));
const FILE = path.join(TMP, 'books_1099.xlsx');
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

async function buildWorkbook() {
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.getRow(1).values = ['Category', 'Sub-Category', 'Type', 'Report', '', '', 'Vendors', '1099', 'Customers'];
    setup.getRow(2).values = ['Services', '', 'Expense', 'P&L'];
    ['At 600', 'Under 600', 'Refunded', 'At 2000', 'Under 2000'].forEach((v, i) => {
        setup.getRow(i + 2).getCell(7).value = v;
        setup.getRow(i + 2).getCell(8).value = 'NEC';
    });

    const ledger = wb.addWorksheet('Ledger');
    ledger.getRow(1).values = ['Date', 'Description', 'Category', 'Debit', 'Credit'];

    const bank = wb.addWorksheet('Bank Transactions');
    bank.addRow(['Date', 'Description', 'Amount', 'Category', 'Vendor']);
    const rows = [
        // 2025 (threshold $600)
        ['2025-03-01', 'Exactly 600', -600.00, 'At 600'],
        ['2025-03-02', 'Just under', -599.99, 'Under 600'],
        ['2025-03-03', 'Paid 700', -700.00, 'Refunded'],
        ['2025-03-20', 'Refund 150', 150.00, 'Refunded'], // net 550
        // 2026 (threshold $2,000)
        ['2026-03-01', 'Same 600 in 2026', -600.00, 'At 600'],
        ['2026-03-02', 'Exactly 2000', -2000.00, 'At 2000'],
        ['2026-03-03', 'Just under 2000', -1999.99, 'Under 2000'],
    ];
    rows.forEach(([d, desc, amt, vendor]) => bank.addRow([d, desc, amt, 'Services', vendor]));
    await wb.xlsx.writeFile(FILE);
}

// Returns { total, required } from the vendor report line, or null.
function vendorLine(output, vendor) {
    const line = output.split('\n').find(l => l.startsWith(vendor.padEnd(30)));
    if (!line) return null;
    const total = parseFloat((line.match(/-?[\d,]+\.\d{2}/) || ['NaN'])[0].replace(/,/g, ''));
    return { total, required: /\bYES\b/.test(line) };
}

function runVendorReport(year) {
    const r = spawnSync(process.execPath, ['report.js', FILE, `--year=${year}`, '--vendor'], { cwd: ROOT, encoding: 'utf8' });
    return r.stdout + r.stderr;
}

async function main() {
    await buildWorkbook();

    const cases = [
        { year: 2025, vendor: 'At 600', total: 600, required: true },
        { year: 2025, vendor: 'Under 600', total: 599.99, required: false },
        { year: 2025, vendor: 'Refunded', total: 550, required: false },
        { year: 2026, vendor: 'At 600', total: 600, required: false },
        { year: 2026, vendor: 'At 2000', total: 2000, required: true },
        { year: 2026, vendor: 'Under 2000', total: 1999.99, required: false },
    ];
    const outputs = { 2025: runVendorReport(2025), 2026: runVendorReport(2026) };

    for (const c of cases) {
        const got = vendorLine(outputs[c.year], c.vendor);
        const label = `${c.year} ${c.vendor}: $${c.total} → ${c.required ? 'Required' : 'not required'}`;
        check(label, got && Math.abs(got.total - c.total) < 0.005 && got.required === c.required,
            got ? `got $${got.total}, required=${got.required}` : 'vendor missing from report');
    }

    fs.rmSync(TMP, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll 1099 boundary tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

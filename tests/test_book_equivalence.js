// --book must give the same reports as the Setup sheet. One synthetic book is built
// twice: as a workbook with Setup tables, and as a TOML book (with a workbook that has
// no Setup sheet, and again with a stale Setup sheet that must be ignored). report.js runs
// on each with --all and --1099, and the output must be identical once the lines that
// legitimately differ (the report stamp, "Setup sheet ignored", the payer-source line, the
// book-mode 1099 hint) are removed. Also: a tax ID never appears in console output, in
// any run, and is in the saved 1099 Data sheet.
//
// The book (year 2026), with the numbers the tests assert:
//   Opening balances (2025 closing): Checking 1,000, Credit Card 200, Equipment 600, Owner Investment 1,000
//   Bank (linked to Checking): +5,000 Sales/Consulting (Client A); -2,500 Acme Plumbing (Contract Labor);
//     -300 Office Depot (Office Supplies/Paper); -100 card autopay (to the Credit Card liability); -400 Equipment
//   Card (linked to Credit Card, charges positive so flipped): +150 Office Depot (Office Supplies/Software);
//     -100 payment received (Transfer Credit Card: left out of the card's balance)
//   Ledger: Dr Checking 500 / Cr Owner Investment 500
//   Checking     = 1,000 + (5,000 - 2,500 - 300 - 100 - 400) + 500 = 3,200
//   Credit Card  = 200 + 150 (charge) - 100 (bank autopay only)      = 250
//   Equipment    = 600 + 400 (unlinked asset: spending increases it)  = 1,000
//   Owner Inv.   = 1,000 + 500                                        = 1,500
//   Net income   = 5,000 - 2,500 - (300 + 150)                        = 2,050
//   Opening plug = 1,000 + 600 - 200 - 1,000                          = 400
//   A = L + E:   (3,200 + 1,000) = 250 + (1,500 + 400 + 2,050)        = 4,200
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-book-equiv-'));
const FAKE_COMPANY_TIN = '00-0000000';
const FAKE_VENDOR_TIN = '00-0000001';
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

function valueFor(output, label) {
    const line = output.split('\n').find(l => new RegExp(`^\\s*${label}\\s*:`).test(l));
    if (!line) return null;
    const nums = line.match(/-?[\d,]+\.\d{2}/g);
    return nums ? parseFloat(nums[nums.length - 1].replace(/,/g, '')) : null;
}

const CLOSING = 'closing.toml'; // same file name in both runs: its name is printed
const CLOSING_TEXT = 'year = 2025\n[accounts]\nChecking = 1000\n"Credit Card" = 200\nEquipment = 600\n"Owner Investment" = 1000\n';

const CATEGORIES = [
    // [name, bookType, setup Account Type, setup Report, subcategories, transfer account]
    ['Sales', 'income', 'Income', 'P&L', ['Consulting', 'Products'], ''],
    ['Office Supplies', 'expense', 'Expense', 'P&L', ['Software', 'Paper'], ''],
    ['Contract Labor', 'expense', 'Expense', 'P&L', [], ''],
    ['Checking', 'asset', 'Asset', 'Balance Sheet', [], ''],
    ['Credit Card', 'liability', 'Liability', 'Balance Sheet', [], ''],
    ['Equipment', 'asset', 'Asset', 'Balance Sheet', [], ''],
    ['Owner Investment', 'equity', 'Equity', 'Balance Sheet', [], ''],
    ['Transfer Credit Card', 'transfer', 'Transfer', 'Transfer', [], 'Credit Card'],
];

const COMPANY = [['Company Name', 'Equiv Holdings LLC'], ['Address', '1 Main St'], ['City', 'Springfield'], ['State', 'CA'],
    ['Zip', '90000'], ['Country', 'US'], ['Email', 'books@example.com'], ['Phone', '555-0100']];

const BANK_ROWS = [
    // [date, description, amount, category, sub-category, vendor, customer]
    [new Date('2026-03-01'), 'Client payment', 5000, 'Sales', 'Consulting', '', 'Client A'],
    [new Date('2026-03-02'), 'Plumbing job', -2500, 'Contract Labor', '', 'Acme Plumbing', ''],
    [new Date('2026-03-03'), 'Paper', -300, 'Office Supplies', 'Paper', 'Office Depot', ''],
    [new Date('2026-03-04'), 'Card autopay', -100, 'Credit Card', '', '', ''],
    [new Date('2026-03-05'), 'Buy equipment', -400, 'Equipment', '', '', ''],
];
const CARD_ROWS = [
    [new Date('2026-03-06'), 'Software', 150, 'Office Supplies', 'Software', 'Office Depot', ''],
    [new Date('2026-03-07'), 'Payment received', -100, 'Transfer Credit Card', '', '', ''],
];
const LEDGER_ROWS = [
    [new Date('2026-04-01'), 'Owner investment', 'Checking', 500, null],
    [new Date('2026-04-01'), 'Owner investment', 'Owner Investment', null, 500],
];
const END_BALANCES = { Bank: 3200, Card: 250 };

function addTransactionSheets(wb) {
    [['Bank', BANK_ROWS], ['Card', CARD_ROWS]].forEach(([name, rows]) => {
        const ws = wb.addWorksheet(name);
        ws.addRow(['Date', 'Description', 'Amount', 'Category', 'Sub-Category', 'Vendor', 'Customer']);
        rows.forEach(r => ws.addRow(r));
    });
    const ledger = wb.addWorksheet('Ledger');
    ledger.addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);
    LEDGER_ROWS.forEach(r => ledger.addRow(r));
}

// The Setup version: header-row detection, as in docs/setup.md.
function addSetupSheet(wb) {
    const setup = wb.addWorksheet('Setup');
    setup.addRow(['Category', 'Sub-Category', 'Account Type', 'Report', 'Transfer Account',
        'Vendors', 'Business Name', '1099 Type', '1099 Required', 'Address', 'Customers',
        'Sheet Name', 'Type', 'Flip', 'Offset', 'Short Name', 'Link Asset', 'Start Balance', 'End Balance', 'Company Info', '']);
    let row = 2;
    CATEGORIES.forEach(([name, , accountType, report, subs, transfer]) => {
        (subs.length ? subs : ['']).forEach(sub => {
            [name, sub, accountType, report, transfer].forEach((v, j) => { setup.getRow(row).getCell(j + 1).value = v; });
            row++;
        });
    });
    [['Acme Plumbing', 'Acme Plumbing Inc', 'NEC', 'YES', '2 Pipe Rd'], ['Office Depot']].forEach((v, i) => {
        v.forEach((x, j) => { setup.getRow(i + 2).getCell(6 + j).value = x; });
    });
    setup.getRow(2).getCell(11).value = 'Client A';
    [['Bank', 'Bank', '', 1, 'Bank', 'Checking', null, END_BALANCES.Bank],
        ['Card', 'CC', '', 1, 'Card', 'Credit Card', null, END_BALANCES.Card],
        ['Ledger', 'Ledger', '', 1, '', '', null, null]].forEach((s, i) => {
        s.forEach((v, j) => { if (v !== null) setup.getRow(i + 2).getCell(12 + j).value = v; });
    });
    COMPANY.forEach(([k, v], i) => { setup.getRow(i + 2).getCell(20).value = k; setup.getRow(i + 2).getCell(21).value = v; });
}

// The same configuration as a book. extras: { private, vendorDetails } for the tax ID runs.
function bookText({ withPrivate = false, completeVendor = false, flags = ['--all'] } = {}) {
    const lines = [
        '[book]', 'id = "equiv"', 'workbook = "equiv.xlsx"', 'year = 2026', `opening_balances = "${CLOSING}"`, `default_flags = [${flags.map(f => `"${f}"`).join(', ')}]`,
        ...(withPrivate ? ['private = "equiv.private.toml"'] : []),
        '', '[company]',
        ...COMPANY.map(([k, v]) => `${k === 'Company Name' ? 'name' : k.toLowerCase()} = "${v}"`), '',
        '[[sheets]]', 'name = "Bank"', 'type = "bank"', 'header_row = 1', 'link = "Checking"', `end_balance = ${END_BALANCES.Bank}`, '',
        '[[sheets]]', 'name = "Card"', 'type = "cc"', 'header_row = 1', 'link = "Credit Card"', `end_balance = ${END_BALANCES.Card}`, '',
        '[[sheets]]', 'name = "Ledger"', 'type = "ledger"', 'header_row = 1', '',
    ];
    CATEGORIES.forEach(([name, type, , , subs, transfer]) => {
        lines.push('[[categories]]', `name = "${name}"`, `type = "${type}"`);
        if (subs.length) lines.push(`subcategories = [${subs.map(s => `"${s}"`).join(', ')}]`);
        if (transfer) lines.push(`transfer_account = "${transfer}"`);
        lines.push('');
    });
    lines.push('[[vendors]]', 'name = "Acme Plumbing"', 'business_name = "Acme Plumbing Inc"', 'form_1099 = "NEC"', 'address = "2 Pipe Rd"');
    if (completeVendor) lines.push('city = "Springfield"', 'state = "CA"', 'zip = "90001"');
    lines.push('', '[[vendors]]', 'name = "Office Depot"', '', '[[customers]]', 'name = "Client A"', '');
    return lines.join('\n');
}

// Removes the lines that differ by design between Setup and --book runs.
function normalize(text) {
    return text.split('\n')
        .filter(l => !/^ {2}(Tool Version|Git Commit|Run Time|Flags|Tax Year|Input: )/.test(l))
        .filter(l => !l.startsWith('[WARNING] Setup sheet is missing formal Excel table') && !l.startsWith('Falling back to row-scanning'))
        .filter(l => l !== 'Report stamp:' && !l.startsWith('Setup sheet ignored') && !l.startsWith('Payer Information (from')
            && !l.includes('With --book: put address fields'))
        .join('\n');
}

function run(args, cwd) {
    const r = spawnSync(process.execPath, [path.join(ROOT, 'report.js'), ...args], { cwd, encoding: 'utf8' });
    return { code: r.status, out: r.stdout, err: r.stderr, all: r.stdout + r.stderr };
}

async function main() {
    // --- the three folders: Setup version, book + workbook without Setup, book + workbook with a stale Setup ---
    const dirs = { setup: path.join(TMP, 'setup'), book: path.join(TMP, 'book'), stale: path.join(TMP, 'stale') };
    Object.values(dirs).forEach(d => fs.mkdirSync(d));
    Object.values(dirs).forEach(d => fs.writeFileSync(path.join(d, CLOSING), CLOSING_TEXT));

    let wb = new ExcelJS.Workbook();
    addSetupSheet(wb);
    addTransactionSheets(wb);
    await wb.xlsx.writeFile(path.join(dirs.setup, 'equiv.xlsx'));

    wb = new ExcelJS.Workbook();
    addTransactionSheets(wb);
    await wb.xlsx.writeFile(path.join(dirs.book, 'equiv.xlsx'));
    fs.writeFileSync(path.join(dirs.book, 'equiv.book.toml'), bookText());

    wb = new ExcelJS.Workbook();
    addSetupSheet(wb);
    addTransactionSheets(wb);
    await wb.xlsx.writeFile(path.join(dirs.stale, 'equiv.xlsx'));
    fs.writeFileSync(path.join(dirs.stale, 'equiv.book.toml'), bookText());

    // --- all reports (--all comes from the book's default_flags in book mode) ---
    const setupRun = run([path.join(dirs.setup, 'equiv.xlsx'), '--all', '--year=2026', '--opening-balances', path.join(dirs.setup, CLOSING), '--ignore-vendors'], dirs.setup);
    const bookRun = run(['--book', path.join(dirs.book, 'equiv.book.toml')], dirs.book);
    const staleRun = run(['--book', path.join(dirs.stale, 'equiv.book.toml')], dirs.stale);

    // The numbers first, so a wrong book cannot pass by being wrong in both.
    const bsRun = run([path.join(dirs.setup, 'equiv.xlsx'), '--bs', '--year=2026', '--opening-balances', path.join(dirs.setup, CLOSING), '--ignore-vendors'], dirs.setup);
    const nums = bsRun.out;
    check('Setup run: Checking = 3,200.00', valueFor(nums, 'Checking') === 3200, `got ${valueFor(nums, 'Checking')}`);
    check('Setup run: Credit Card = 250.00', valueFor(nums, 'Credit Card') === 250, `got ${valueFor(nums, 'Credit Card')}`);
    check('Setup run: Equipment (unlinked asset) = 1,000.00', valueFor(nums, 'Equipment') === 1000, `got ${valueFor(nums, 'Equipment')}`);
    check('Setup run: Owner Investment = 1,500.00', valueFor(nums, 'Owner Investment') === 1500, `got ${valueFor(nums, 'Owner Investment')}`);
    check('Setup run: net income 2,050.00', setupRun.out.includes('NET INCOME: 2,050.00'), String(setupRun.out.match(/NET INCOME[^\n]*/)));
    check('Setup run: opening plug 400.00 and A = L + E', valueFor(nums, 'Retained Earnings \\(Opening Balance\\)') === 400 && nums.includes('[OK] (A = L + E)'));
    check('Setup run: the transfer is left out of the card (Credit Card 250, not 150)', valueFor(nums, 'Credit Card') !== 150);
    check('Setup run: 1099 lists Acme Plumbing at 2,500.00', /Acme Plumbing\s+Acme Plumbing Inc\s+2500\.00\s+NEC/.test(setupRun.out), setupRun.out);

    for (const [label, r] of [['book without Setup', bookRun], ['book with a stale Setup sheet', staleRun]]) {
        check(`${label}: same exit code as the Setup run (${setupRun.code})`, r.code === setupRun.code, `exit ${r.code}\n${r.all}`);
        check(`${label}: stdout identical to the Setup run`, normalize(r.out) === normalize(setupRun.out), diffHint(normalize(setupRun.out), normalize(r.out)));
        check(`${label}: stderr identical to the Setup run`, normalize(r.err) === normalize(setupRun.err), diffHint(normalize(setupRun.err), normalize(r.err)));
        check(`${label}: no Setup-sheet or vendor.xlsx noise`, !r.all.includes('Setup sheet is missing') && !r.all.includes('vendor.xlsx" or'), r.all);
    }
    check('Stale Setup sheet: one line says it was ignored', staleRun.out.includes('Setup sheet ignored: config comes from equiv.book.toml'));
    check('Book without Setup: no "ignored" line', !bookRun.out.includes('Setup sheet ignored'));
    check('Book run: stamp lists workbook, book and opening-balances file',
        /Input: Workbook\s+: equiv\.xlsx/.test(bookRun.out) && /Input: Book\s+: equiv\.book\.toml/.test(bookRun.out) && /Input: Opening balances\s+: closing\.toml/.test(bookRun.out), bookRun.out);
    check('Book run: opening_balances in the book is used (summary line printed)', bookRun.out.includes('Opening balances: 4 accounts from closing.toml (closing year 2025)'));

    // Explicit --1099 (not from default_flags) gives the same 1099 section too
    const setup1099 = run([path.join(dirs.setup, 'equiv.xlsx'), '--1099', '--year=2026', '--opening-balances', path.join(dirs.setup, CLOSING), '--ignore-vendors'], dirs.setup);
    fs.writeFileSync(path.join(dirs.book, 'only1099.book.toml'), bookText({ flags: [] }));
    const book1099 = run(['--book', path.join(dirs.book, 'only1099.book.toml'), '--1099'], dirs.book);
    check('--1099 only: identical to the Setup run', normalize(book1099.all) === normalize(setup1099.all) && book1099.code === setup1099.code, diffHint(normalize(setup1099.all), normalize(book1099.all)));

    // --- tax IDs: never on the console, present in the saved 1099 Data sheet ---
    const taxDir = path.join(TMP, 'tax');
    fs.mkdirSync(taxDir);
    fs.writeFileSync(path.join(taxDir, CLOSING), CLOSING_TEXT);
    wb = new ExcelJS.Workbook();
    addTransactionSheets(wb);
    await wb.xlsx.writeFile(path.join(taxDir, 'equiv.xlsx'));
    fs.writeFileSync(path.join(taxDir, 'equiv.book.toml'), bookText({ withPrivate: true, completeVendor: true }));
    const privateText = `[company]\ntin = "${FAKE_COMPANY_TIN}"\n\n[vendors."Acme Plumbing"]\ntin = "${FAKE_VENDOR_TIN}"\n`;
    fs.writeFileSync(path.join(taxDir, 'equiv.private.toml'), privateText);
    const bookFile = path.join(taxDir, 'equiv.book.toml');

    const taxRuns = {
        '--1099': run(['--book', bookFile, '--1099'], taxDir),
        '--all (from default_flags)': run(['--book', bookFile], taxDir),
        '--all --debug --checker': run(['--book', bookFile, '--debug', '--checker'], taxDir),
        '--1099 --details': run(['--book', bookFile, '--1099', '--details', 'Acme Plumbing'], taxDir),
        '--1099 --save': run(['--book', bookFile, '--1099', '--save'], taxDir),
    };
    for (const [label, r] of Object.entries(taxRuns)) {
        check(`Tax IDs: not in stdout or stderr (${label})`, !r.all.includes(FAKE_COMPANY_TIN) && !r.all.includes(FAKE_VENDOR_TIN), r.all);
    }
    const t = taxRuns['--1099'];
    check('Tax IDs: complete vendor details and a private file give exit 0', t.code === 0, t.all);
    check('Tax IDs: payer and recipient show "on file", not the number', t.out.includes('Tax ID:  on file') && t.out.includes('TIN:     on file') && /Acme Plumbing\s+Acme Plumbing Inc\s+on file\s+2500\.00\s+NEC/.test(t.out), t.out);
    check('Tax IDs: no "missing" complaint when they are on file', !t.all.includes('Incomplete Data'), t.all);
    check('Tax IDs: the private file is in the stamp (hash only)', /Input: Private file\s+: equiv\.private\.toml\s+sha256 [0-9a-f]{12}/.test(t.out));

    const saved = path.join(taxDir, 'report_equiv.xlsx');
    check('--save wrote the report workbook', fs.existsSync(saved), taxRuns['--1099 --save'].all);
    if (fs.existsSync(saved)) {
        const rep = new ExcelJS.Workbook();
        await rep.xlsx.readFile(saved);
        const data = rep.getWorksheet('1099 Data');
        const cells = [];
        data.eachRow(row => row.eachCell(c => cells.push(String(c.value))));
        check('Saved 1099 Data sheet has the company tax ID', cells.includes(FAKE_COMPANY_TIN), cells.join('|'));
        check('Saved 1099 Data sheet has the vendor tax ID', cells.includes(FAKE_VENDOR_TIN), cells.join('|'));
        // ...and nowhere else in the saved report (the Processing Log is a copy of the console)
        const elsewhere = [];
        rep.eachSheet(ws => {
            if (ws.name === '1099 Data') return;
            ws.eachRow(row => row.eachCell(c => { if (String(c.value).includes('00-000000')) elsewhere.push(`${ws.name}: ${c.value}`); }));
        });
        check('Tax IDs are in no other sheet of the saved report (Processing Log, Report Info)', elsewhere.length === 0, elsewhere.join('\n'));
    }

    // Without the private file the IDs are reported as missing (still never printed)
    const noPrivDir = path.join(TMP, 'nopriv');
    fs.mkdirSync(noPrivDir);
    fs.writeFileSync(path.join(noPrivDir, CLOSING), CLOSING_TEXT);
    wb = new ExcelJS.Workbook();
    addTransactionSheets(wb);
    await wb.xlsx.writeFile(path.join(noPrivDir, 'equiv.xlsx'));
    fs.writeFileSync(path.join(noPrivDir, 'equiv.book.toml'), bookText({ completeVendor: true }));
    const missing = run(['--book', path.join(noPrivDir, 'equiv.book.toml'), '--1099'], noPrivDir);
    check('No private file: the missing tax ID is reported and the run exits 1',
        missing.code === 1 && missing.all.includes('Missing: Tax ID (SSN/EIN)') && missing.all.includes('Tax ID:  MISSING'), missing.all);

    // A private file named in the book but absent is a clear error (only when 1099 is asked for)
    fs.writeFileSync(path.join(noPrivDir, 'equiv.book.toml'), bookText({ withPrivate: true, completeVendor: true, flags: [] }));
    const absent = run(['--book', path.join(noPrivDir, 'equiv.book.toml'), '--1099'], noPrivDir);
    check('Private file named but missing: exit 1 with the fix', absent.code === 1 && absent.all.includes('Private file') && absent.all.includes('not found'), absent.all);
    const notNeeded = run(['--book', path.join(noPrivDir, 'equiv.book.toml'), '--pl'], noPrivDir);
    check('Private file named but missing, no 1099 asked: runs normally', notNeeded.code === 0, notNeeded.all);

    // The CLI --opening-balances wins over the book's opening_balances
    const other = path.join(taxDir, 'other-closing.toml');
    fs.writeFileSync(other, 'year = 2025\n[accounts]\nChecking = 1\n');
    const override = run(['--book', bookFile, '--bs', '--opening-balances', other], taxDir);
    check('CLI --opening-balances overrides the book\'s file', override.all.includes('from other-closing.toml'), override.all);

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll book equivalence tests passed.');
}

// The first line where two outputs differ, for a failure message.
function diffHint(a, b) {
    const x = a.split('\n');
    const y = b.split('\n');
    for (let i = 0; i < Math.max(x.length, y.length); i++) {
        if (x[i] !== y[i]) return `first difference at line ${i + 1}:\n  setup: ${x[i]}\n  book : ${y[i]}`;
    }
    return '';
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

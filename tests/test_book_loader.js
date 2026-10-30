// Tests load_transactions.js --book:
//   - the lock check runs first (before the export is read) and changes nothing
//   - the generated Lists sheet (contents, note, protection)
//   - dropdowns and the Report Type formula point at Lists, not Setup
//   - --refresh-lists rewrites only Lists
//   - without --book the old Setup behavior is unchanged
const { spawn, spawnSync } = require('child_process');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { assertWorkbookUnlocked } = require('../lib/book/lock');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-book-loader-'));
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

function run(script, args, cwd = ROOT) {
    const r = spawnSync(process.execPath, [path.join(ROOT, script), ...args], { cwd, encoding: 'utf8' });
    return { code: r.status, out: r.stdout + r.stderr };
}

const sha = file => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

const BOOK = `
[book]
id = "loader"
workbook = "Loader_2026.xlsx"

[[sheets]]
name = "Chase Checking"
type = "bank"
link = "Checking"
end_balance = 0

[[sheets]]
name = "Visa Card"
type = "cc"
link = "Credit Card"
end_balance = 0

[[sheets]]
name = "Ledger"
type = "ledger"

[[categories]]
name = "Sales"
type = "income"
subcategories = ["Consulting", "Products"]

[[categories]]
name = "Office Supplies"
type = "expense"
subcategories = ["Software", "Consulting"]

[[categories]]
name = "Checking"
type = "asset"

[[categories]]
name = "Credit Card"
type = "liability"

[[categories]]
name = "Transfer Credit Card"
type = "transfer"
transfer_account = "Credit Card"

[[vendors]]
name = "Acme Plumbing"

[[vendors]]
name = "Office Depot"

[[customers]]
name = "Client A"
`;

function colValues(ws, col) {
    const out = [];
    for (let r = 2; r <= ws.rowCount; r++) {
        const v = ws.getCell(`${col}${r}`).value;
        if (v !== null && v !== undefined && v !== '') out.push(String(v));
    }
    return out;
}

function testLockFunction() {
    // A mocked fs: simulates what Windows reports when Excel holds the file.
    const failing = code => ({ openSync: () => { const e = new Error(code); e.code = code; throw e; }, closeSync: () => {} });
    for (const code of ['EBUSY', 'EPERM', 'EACCES']) {
        let msg = null;
        try { assertWorkbookUnlocked('Books.xlsx', failing(code)); } catch (e) { msg = e.message; }
        check(`Lock check (mocked fs): ${code} gives a "close it in Excel" error`, msg !== null && msg.includes('close it in Excel') && msg.includes(code) && msg.includes('Nothing was changed'), msg);
    }
    let missing = null;
    try { assertWorkbookUnlocked('Books.xlsx', failing('ENOENT')); } catch (e) { missing = e.message; }
    check('Lock check (mocked fs): ENOENT says the workbook is not found', missing !== null && missing.includes('not found'), missing);
    let other = null;
    try { assertWorkbookUnlocked('Books.xlsx', failing('EMFILE')); } catch (e) { other = e.code; }
    check('Lock check (mocked fs): an unrelated error is not hidden', other === 'EMFILE');
    const calls = [];
    assertWorkbookUnlocked('Books.xlsx', { openSync: (f, flag) => { calls.push(`open ${flag}`); return 7; }, closeSync: fd => calls.push(`close ${fd}`) });
    check('Lock check (mocked fs): opens for update and closes the handle', calls.join(',') === 'open r+,close 7', calls.join(','));
}

// Holds an exclusive lock on file for a few seconds, as Excel does (Windows only). Resolves with a release().
function holdLock(file) {
    return new Promise((resolve, reject) => {
        const script = `$f=[System.IO.File]::Open('${file.replace(/'/g, "''")}','Open','ReadWrite','None'); Write-Output LOCKED; Start-Sleep -Seconds 60`;
        const child = spawn('powershell', ['-NoProfile', '-Command', script], { stdio: ['ignore', 'pipe', 'ignore'] });
        let done = false;
        child.stdout.on('data', d => {
            if (!done && d.toString().includes('LOCKED')) { done = true; resolve(() => child.kill()); }
        });
        child.on('error', reject);
        child.on('exit', () => { if (!done) reject(new Error('lock helper exited early')); });
    });
}

async function main() {
    testLockFunction();

    const dir = path.join(TMP, 'books');
    fs.mkdirSync(dir);
    const bookFile = path.join(dir, 'loader.book.toml');
    fs.writeFileSync(bookFile, BOOK);
    const wbFile = path.join(dir, 'Loader_2026.xlsx');
    const bankCsv = path.join(ROOT, 'tests', 'example_bank.csv');
    const ccCsv = path.join(ROOT, 'tests', 'example_cc.csv');

    const gen = run('generate_excel.js', ['--book', bookFile]);
    check('Workbook generated from the book', gen.code === 0 && fs.existsSync(wbFile), gen.out);
    let wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(wbFile);
    check('Generated workbook has the book sheets, a Ledger and Lists, and no Setup',
        ['Chase Checking', 'Visa Card', 'Ledger', 'Lists'].every(n => wb.getWorksheet(n)) && !wb.getWorksheet('Setup'), wb.worksheets.map(s => s.name).join(','));

    // --- load a bank export and a card export ---
    const bank = run('load_transactions.js', ['--book', bookFile, bankCsv, 'Chase Checking']);
    check('Bank import exits 0', bank.code === 0 && bank.out.includes('Successfully processed'), bank.out);
    const card = run('load_transactions.js', ['--book', bookFile, ccCsv, 'Visa Card']);
    check('Card import exits 0', card.code === 0 && card.out.includes('Successfully processed'), card.out);

    wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(wbFile);
    const lists = wb.getWorksheet('Lists');
    check('Lists: Categories column in book order', colValues(lists, 'A').join('|') === 'Sales|Office Supplies|Checking|Credit Card|Transfer Credit Card', colValues(lists, 'A').join('|'));
    check('Lists: SubCategories unique across categories', colValues(lists, 'B').join('|') === 'Consulting|Products|Software', colValues(lists, 'B').join('|'));
    check('Lists: Vendors and Customers', colValues(lists, 'C').join('|') === 'Acme Plumbing|Office Depot' && colValues(lists, 'D').join('|') === 'Client A');
    check('Lists: header row names the four lists', ['Categories', 'SubCategories', 'Vendors', 'Customers'].every((h, i) => lists.getCell(1, i + 1).value === h));
    check('Lists: a note says it is generated and how to refresh it', /GENERATED from loader\.book\.toml/.test(String(lists.getCell('G1').value)) && String(lists.getCell('G1').value).includes('--refresh-lists'));
    check('Lists: the sheet is protected', lists.sheetProtection && lists.sheetProtection.sheet === true);

    const bankWs = wb.getWorksheet('Chase Checking');
    const bankRows = bankWs.rowCount - 3;
    check('Bank sheet: header on row 3, TOTAL above, rows loaded', bankWs.getCell('A1').value === 'TOTAL' && bankWs.getCell('A3').value === 'Date' && bankRows === 3, `rows ${bankRows}`);
    const dv = c => (bankWs.getCell(c).dataValidation || {}).formulae;
    // Sub-Category follows the row's Category (column D here); every sub-category while Category is empty.
    const subFor = cell => [`IF(${cell}="",Lists!$B$2:$B$4,OFFSET(Lists!$J$2,0,MATCH(${cell},Lists!$J$1:$N$1,0)-1,MAX(1,COUNTA(OFFSET(Lists!$J$2,0,MATCH(${cell},Lists!$J$1:$N$1,0)-1,200,1))),1))`];
    check('Bank dropdowns point at Lists (category, sub-category by category, vendor, customer)',
        JSON.stringify([dv('D4'), dv('E4'), dv('G4'), dv('H4')]) === JSON.stringify([['Lists!$A$2:$A$6'], subFor('$D4'), ['Lists!$C$2:$C$3'], ['Lists!$D$2:$D$2']]),
        JSON.stringify([dv('D4'), dv('E4'), dv('G4'), dv('H4')]));
    check('Lists: sub-categories per category from column J', lists.getCell('J1').value === 'Sales' && lists.getCell('J2').value === 'Consulting' && lists.getCell('J3').value === 'Products');
    const reportType = bankWs.getCell('I4').value.formula;
    check('Report Type formula looks up the category in Lists', reportType.includes('Lists!$E$2:$E$2000') && reportType.includes('MATCH(D4,Lists!$A$2:$A$2000,0)'), reportType);
    const cardWs = wb.getWorksheet('Visa Card');
    const cdv = c => (cardWs.getCell(c).dataValidation || {}).formulae;
    check('Card dropdowns (columns E, F, H, I) point at Lists',
        JSON.stringify([cdv('E4'), cdv('F4'), cdv('H4'), cdv('I4')]) === JSON.stringify([['Lists!$A$2:$A$6'], subFor('$E4'), ['Lists!$C$2:$C$3'], ['Lists!$D$2:$D$2']]));
    check('Card Report Type formula uses column E and Lists', cardWs.getCell('L4').value.formula.includes('MATCH(E4,Lists!$A$2:$A$2000,0)'));
    let mentionsSetup = false;
    [bankWs, cardWs].forEach(ws => ws.eachRow(row => row.eachCell(c => {
        if ((c.dataValidation && JSON.stringify(c.dataValidation).includes('Setup!')) || (c.value && c.value.formula && c.value.formula.includes('Setup!'))) mentionsSetup = true;
    })));
    check('No formula or dropdown refers to Setup', !mentionsSetup);

    // --- --refresh-lists after editing the book ---
    fs.writeFileSync(bookFile, BOOK.replace('[[customers]]', '[[categories]]\nname = "Rent"\ntype = "expense"\n\n[[customers]]\nname = "Client B"\n\n[[customers]]'));
    const before = {
        bankRows: bankWs.getRow(4).values.slice(1, 9).join('|'),
        cardRows: cardWs.rowCount,
        version: wb.getWorksheet('VERSION').rowCount,
    };
    const refresh = run('load_transactions.js', ['--book', bookFile, '--refresh-lists']);
    check('--refresh-lists exits 0 and says what it wrote', refresh.code === 0 && refresh.out.includes('Lists sheet written: 6 categories') && refresh.out.includes('2 customers'), refresh.out);
    wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(wbFile);
    check('--refresh-lists: Lists has the new category and customer', colValues(wb.getWorksheet('Lists'), 'A').includes('Rent') && colValues(wb.getWorksheet('Lists'), 'D').join('|') === 'Client B|Client A');
    check('--refresh-lists: the transaction sheets and VERSION are untouched',
        wb.getWorksheet('Chase Checking').getRow(4).values.slice(1, 9).join('|') === before.bankRows && wb.getWorksheet('Visa Card').rowCount === before.cardRows
        && wb.getWorksheet('VERSION').rowCount === before.version);
    check('--refresh-lists: still one Lists sheet, protected', wb.worksheets.filter(s => s.name === 'Lists').length === 1 && wb.getWorksheet('Lists').sheetProtection.sheet === true);

    // --- A second import lands right under the first ---
    // The dropdowns reach 1000 rows past the data; they must not push the next import down.
    const again = run('load_transactions.js', ['--book', bookFile, bankCsv, 'Chase Checking']);
    wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(wbFile);
    const afterTwo = wb.getWorksheet('Chase Checking');
    const dataRows = [];
    afterTwo.eachRow((row, r) => { if (r > 3) dataRows.push(r); });
    check('Second import: rows follow the first with no gap (rows 4-9)', again.code === 0 && dataRows.join(',') === '4,5,6,7,8,9', dataRows.join(','));

    // --- errors ---
    const unknown =run('load_transactions.js', ['--book', bookFile, bankCsv, 'Chase Checkin']);
    check('Unknown sheet: exit 1 with the closest sheet name', unknown.code === 1 && unknown.out.includes('Did you mean "Chase Checking"?'), unknown.out);
    const ledger = run('load_transactions.js', ['--book', bookFile, bankCsv, 'Ledger']);
    check('Loading into the ledger is refused', ledger.code === 1 && ledger.out.includes('is the ledger'), ledger.out);
    const noArgs = run('load_transactions.js', ['--book', bookFile]);
    check('Missing export and sheet: exit 1 with an example', noArgs.code === 1 && noArgs.out.includes('give <inputFile> and <sheet>'), noArgs.out);
    const missingExport = run('load_transactions.js', ['--book', bookFile, path.join(dir, 'nope.csv'), 'Chase Checking']);
    check('Missing export file: exit 1', missingExport.code === 1 && missingExport.out.includes('Input file not found'), missingExport.out);
    const badBook = path.join(dir, 'bad.book.toml');
    fs.writeFileSync(badBook, BOOK.replace('type = "bank"', 'type = "bnak"'));
    const bad = run('load_transactions.js', ['--book', badBook, bankCsv, 'Chase Checking']);
    check('A bad book stops the import with the fix', bad.code === 1 && bad.out.includes('Did you mean "bank"?'), bad.out);

    // --- the lock: Excel holds the file ---
    if (process.platform === 'win32') {
        const hashBefore = sha(wbFile);
        let release;
        try {
            release = await holdLock(wbFile);
        } catch (e) {
            console.warn(`(skipped the real lock test: ${e.message})`);
        }
        if (release) {
            // The export does not exist: the lock error must come first, because the export is not read before the lock is checked.
            const locked = run('load_transactions.js', ['--book', bookFile, path.join(dir, 'nope.csv'), 'Chase Checking']);
            const lockedRefresh = run('load_transactions.js', ['--book', bookFile, '--refresh-lists']);
            release();
            check('Locked workbook: exit 1 and "close it in Excel", before the export is read', locked.code === 1 && locked.out.includes('close it in Excel') && !locked.out.includes('Input file not found'), locked.out);
            check('Locked workbook: --refresh-lists is refused the same way', lockedRefresh.code === 1 && lockedRefresh.out.includes('close it in Excel'), lockedRefresh.out);
            await new Promise(r => setTimeout(r, 500));
            check('Locked workbook: the file is byte-identical afterwards', sha(wbFile) === hashBefore);
        }
    } else {
        console.log('(real file-lock test is Windows only; the lock function is covered with a mocked fs above)');
    }

    // --- without --book: the Setup sheet still drives everything ---
    const legacy = path.join(TMP, 'legacy.xlsx');
    check('Legacy: generate_excel.js still makes the Setup template', run('generate_excel.js', [legacy]).code === 0 && fs.existsSync(legacy));
    const imp = run('load_transactions.js', [bankCsv, 'bank', legacy]);
    check('Legacy: load_transactions.js without --book imports', imp.code === 0 && imp.out.includes('Successfully processed'), imp.out);
    wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(legacy);
    const legacyBank = wb.getWorksheet('Bank Transactions');
    check('Legacy: dropdowns still point at Setup and there is no Lists sheet',
        legacyBank.getCell('D4').dataValidation.formulae[0] === 'Setup!$A$2:$A$100' && legacyBank.getCell('I4').value.formula.includes('VLOOKUP(D4,Setup!A:D,4,FALSE)') && !wb.getWorksheet('Lists'));

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll book loader tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

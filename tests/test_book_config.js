// Tests --book (lib/book/): validating the book and private files, the registry,
// path resolution, default_flags and --year merging, and the generated example files.
//   1. Closest-match suggestions
//   2. Validation: every problem listed, each with a suggested fix
//   3. The private file (and that a tax ID never appears in a message)
//   4. Finding a book: path, books.toml in the current folder, LLC_BOOKS, relative paths
//   5. Merging with the command line (default_flags, --year, workbook)
//   6. End to end with report.js, using examples/example.book.toml and a workbook from generate_excel.js
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { closest, levenshtein } = require('../lib/book/suggest');
const { parseBookText, parsePrivateText, checkBookAgainstWorkbook, loadBook, mergeBookArgs } = require('../lib/book');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-book-config-'));
const FAKE_TIN = '00-0000000';
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

// Returns the error message thrown by fn, or null if it did not throw.
function errorOf(fn) {
    try { fn(); } catch (e) { return e.message; }
    return null;
}

// Passes if fn throws an error containing every fragment.
function rejects(label, fn, ...fragments) {
    const msg = errorOf(fn);
    check(label, msg !== null && fragments.every(f => msg.includes(f)), `got ${msg === null ? 'no error' : JSON.stringify(msg)}`);
}

const BOOK_FILE = path.join(TMP, 'acme.book.toml');

// A valid book. Tests change one thing in it with swap().
const VALID = `
[book]
id = "acme"
name = "Acme Holdings LLC"
workbook = "Acme_2026.xlsx"
year = 2026
opening_balances = "acme-2025-closing.toml"
private = "acme.private.toml"
default_flags = ["--pl", "--bs"]

[company]
name = "Acme Holdings LLC"
city = "Springfield"

[[sheets]]
name = "Chase Checking"
type = "bank"
link = "Checking"
end_balance = 900.00

[[sheets]]
name = "Visa"
type = "cc"
link = "Credit Card"
end_balance = 0

[[sheets]]
name = "Ledger"
type = "ledger"

[[categories]]
name = "Sales"
type = "income"
subcategories = ["Consulting"]

[[categories]]
name = "Office Supplies"
type = "expense"
subcategories = ["Software", "Paper"]

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
form_1099 = "NEC"

[[customers]]
name = "Client A"
`;

function swap(from, to) {
    if (!VALID.includes(from)) throw new Error(`test bug: "${from}" not in VALID`);
    return VALID.replace(from, to);
}

function parse(text) { return parseBookText(text, BOOK_FILE); }

function testSuggestions() {
    check('levenshtein: bnak -> bank is 2', levenshtein('bnak', 'bank') === 2);
    check('closest: a typo gets the nearest candidate', closest('Chase Checkin', ['Visa', 'Chase Checking']) === 'Chase Checking');
    check('closest: case is ignored', closest('CHECKING', ['Checking', 'Credit Card']) === 'Checking');
    check('closest: an unrelated word gets no suggestion', closest('zzzzzz', ['bank', 'cc', 'income']) === null);
    check('closest: a short word needs a tiny distance (xy vs cc)', closest('xy', ['cc']) === null);
}

function testValidation() {
    const ok = parse(VALID);
    check('Valid book: id, name, year', ok.id === 'acme' && ok.name === 'Acme Holdings LLC' && ok.year === 2026);
    check('Valid book: sheets keep order and defaults', ok.sheets.length === 3 && ok.sheets[0].name === 'Chase Checking'
        && ok.sheets[0].flip === null && ok.sheets[0].headerRow === 0 && ok.sheets[0].endBalance === 900 && ok.sheets[0].link === 'Checking');
    check('Valid book: relative paths resolve against the book file folder',
        ok.workbookPath === path.join(TMP, 'Acme_2026.xlsx') && ok.openingBalancesPath === path.join(TMP, 'acme-2025-closing.toml')
        && ok.privatePath === path.join(TMP, 'acme.private.toml'));
    check('Valid book: default_flags kept', ok.defaultFlags.join(' ') === '--pl --bs');
    check('Valid book: the ledger sheet needs no link or end balance', ok.sheets[2].type === 'ledger' && ok.sheets[2].link === null && ok.sheets[2].endBalance === null);

    rejects('Unknown [book] key gets the closest valid key', () => parse(swap('workbook = "Acme_2026.xlsx"', 'workbok = "Acme_2026.xlsx"')),
        '[book]: unknown key "workbok"', 'Did you mean "workbook"?', '"workbook" is missing');
    rejects('Unknown sheet key gets the closest valid key', () => parse(swap('end_balance = 900.00', 'end_ballance = 900.00')),
        'unknown key "end_ballance"', 'Did you mean "end_balance"?');
    rejects('Unknown top-level table gets the closest name', () => parse(`${VALID}\n[[vendor]]\nname = "X"\n`),
        'Unknown top-level key or table "vendor"', 'Did you mean "vendors"?');
    rejects('Bad sheet type suggests the closest and lists the valid ones', () => parse(swap('type = "bank"', 'type = "bnak"')),
        'type "bnak" is not valid', 'Did you mean "bank"?', 'Valid: bank, cc, income, expense, ledger');
    rejects('Bad category type suggests the closest', () => parse(swap('type = "expense"', 'type = "expence"')),
        'type "expence" is not valid', 'Did you mean "expense"?');
    rejects('Bad form_1099 suggests the closest', () => parse(swap('form_1099 = "NEC"', 'form_1099 = "NECC"')),
        'form_1099 "NECC" is not valid', 'Did you mean "NEC"?');
    rejects('A tax ID key in the book file points to the private file', () => parse(swap('city = "Springfield"', 'tin = "x"')),
        '"tin" does not belong in the book file', 'private file');

    rejects('Link to an unknown category suggests the closest', () => parse(swap('link = "Checking"', 'link = "Chekcing"')),
        'link "Chekcing" is not a category', 'Did you mean "Checking"?');
    rejects('Link to a P&L category is rejected', () => parse(swap('link = "Checking"', 'link = "Sales"')),
        'link "Sales" is a P&L category', 'Balance Sheet category');
    rejects('Link to a Transfer category is rejected', () => parse(swap('link = "Checking"', 'link = "Transfer Credit Card"')),
        'a Transfer category');
    rejects('A bank sheet without a link is rejected with the Balance Sheet categories', () => parse(swap('link = "Checking"\n', '')),
        '"link" is missing', '"Checking"', '"Credit Card"');
    // A year in progress has no December statement: the book loads; the report flags UNRECONCILED.
    check('A bank sheet without end_balance loads with endBalance null', parse(swap('end_balance = 900.00', '')).sheets[0].endBalance === null);
    check('An explicit end_balance of 0 is accepted (the card sheet)', parse(VALID).sheets[1].endBalance === 0);
    const described = swap('type = "income"', 'type = "income"\ndescription = "unit A plus connex"');
    check('A category description is accepted', parse(described).categories.length === parse(VALID).categories.length);
    rejects('A description that is not text is rejected', () => parse(swap('type = "income"', 'type = "income"\ndescription = 5')), 'description');
    const withBlank = t => swap('default_flags = ["--pl", "--bs"]', `default_flags = ["--pl", "--bs"]\nblank_sheets = ${t}`);
    check('blank_sheets is read', JSON.stringify(parse(withBlank('["Worksheet"]')).blankSheets) === '["Worksheet"]');
    rejects('A blank sheet named like a book sheet is rejected', () => parse(withBlank('["chase checking"]')), 'blank_sheets "chase checking" has the same name as sheet "Chase Checking"');
    rejects('A blank sheet named Lists is rejected', () => parse(withBlank('["Lists"]')), 'the generated Lists tab');
    rejects('A quoted end_balance is rejected', () => parse(swap('end_balance = 900.00', 'end_balance = "900.00"')),
        '"end_balance" must be a number', 'remove the quotes');
    rejects('An end_balance with 3 decimals is rejected', () => parse(swap('end_balance = 900.00', 'end_balance = 900.005')), 'more than 2 decimals');
    rejects('A ledger sheet with another name is rejected', () => parse(swap('name = "Ledger"', 'name = "Journal"')), 'must be named "Ledger" or "General Ledger"');
    rejects('A ledger sheet with a link is rejected', () => parse(swap('type = "ledger"', 'type = "ledger"\nlink = "Checking"')), 'ledger sheet has no "link"');

    rejects('Duplicate sheet names (case-insensitive) are rejected', () => parse(`${VALID}\n[[sheets]]\nname = "chase checking"\ntype = "bank"\nlink = "Checking"\nend_balance = 1\n`),
        '[[sheets]]: "chase checking" appears twice');
    rejects('Duplicate categories (case-insensitive) are rejected', () => parse(`${VALID}\n[[categories]]\nname = "SALES"\ntype = "income"\n`),
        '[[categories]]: "SALES" appears twice');
    rejects('Duplicate sub-categories in one category are rejected', () => parse(swap('subcategories = ["Software", "Paper"]', 'subcategories = ["Software", "software"]')),
        'subcategories: "software" appears twice');
    rejects('Duplicate vendors (spacing and case ignored) are rejected', () => parse(`${VALID}\n[[vendors]]\nname = "acme  plumbing"\n`),
        '[[vendors]]: "acme  plumbing" appears twice');
    rejects('Duplicate customers are rejected', () => parse(`${VALID}\n[[customers]]\nname = "CLIENT A"\n`),
        '[[customers]]: "CLIENT A" appears twice');
    rejects('A vendor named like a category is rejected', () => parse(`${VALID}\n[[vendors]]\nname = "Office Supplies"\n`),
        '[[vendors]] "Office Supplies" has the same name as a category');
    rejects('A customer named like a category is rejected', () => parse(`${VALID}\n[[customers]]\nname = "sales"\n`),
        '[[customers]] "sales" has the same name as a category');
    rejects('transfer_account on a non-transfer category is rejected', () => parse(swap('subcategories = ["Consulting"]', 'subcategories = ["Consulting"]\ntransfer_account = "Checking"')),
        '"transfer_account" only applies to type = "transfer"');
    rejects('transfer_account naming no Balance Sheet category is rejected with a suggestion', () => parse(swap('transfer_account = "Credit Card"', 'transfer_account = "Credit Crad"')),
        'transfer_account "Credit Crad" is not a Balance Sheet category', 'Did you mean "Credit Card"?');
    rejects('default_flags: an unknown flag gets the closest', () => parse(swap('"--pl", "--bs"', '"--pl-sbu"')),
        'unknown flag "--pl-sbu"', 'Did you mean "--pl-sub"?');
    rejects('default_flags: --year is refused with the fix', () => parse(swap('"--pl", "--bs"', '"--year=2025"')), 'set year = 2026 in [book] instead');
    rejects('[sheets] written as a single table is rejected with the fix', () => parse(VALID.replace(/\[\[sheets\]\][\s\S]*?(?=\[\[categories\]\])/, '[sheets]\nname = "X"\ntype = "bank"\n\n')),
        'Fix: use [[sheets]] instead of [sheets]');
    rejects('A missing [book] table is rejected', () => parse(VALID.replace(/\[book\][\s\S]*?\n\n/, '')), 'The [book] table is missing');

    rejects('A TOML duplicate key is reported by smol-toml', () => parse(swap('year = 2026', 'year = 2026\nyear = 2027')), 'is not valid TOML', 'redefine');

    const all = errorOf(() => parse(swap('type = "bank"', 'type = "bnak"').replace('workbook =', 'workbok =').replace('link = "Credit Card"', 'link = "Sales"')));
    check('Every problem is listed in one error', all !== null && /has \d+ problem\(s\)/.test(all) && all.includes('"bnak"') && all.includes('workbok') && all.includes('link "Sales"'), all);

    // The sheet check needs the workbook's worksheet names.
    const book = parse(VALID);
    rejects('A sheet not in the workbook gets the closest worksheet name', () => checkBookAgainstWorkbook(book, ['Chase Checkin', 'Visa', 'Ledger', 'Lists']),
        'Sheet "Chase Checking" is not in the workbook', 'Did you mean "Chase Checkin"?');
    check('All sheets present: no error', errorOf(() => checkBookAgainstWorkbook(book, ['Chase Checking', 'Visa', 'Ledger'])) === null);
}

function testPrivate() {
    const book = parse(VALID);
    const text = `[company]\ntin = "${FAKE_TIN}"\n\n[vendors."acme plumbing"]\ntin = "00-0000001"\n`;
    const priv = parsePrivateText(text, 'acme.private.toml', book);
    check('Private: company and vendor tax IDs read (vendor name matched ignoring case)', priv.companyTin === FAKE_TIN && priv.vendorTins.get('acme plumbing') === '00-0000001');

    const msgs = [
        errorOf(() => parsePrivateText(`[company]\nssn = "${FAKE_TIN}"\n`, 'p.toml', book)),
        errorOf(() => parsePrivateText(`[company]\ntin = ${FAKE_TIN} oops\n`, 'p.toml', book)),
        errorOf(() => parsePrivateText(`[company]\ntin = "${FAKE_TIN}"\ntin = "${FAKE_TIN}"\n`, 'p.toml', book)),
        errorOf(() => parsePrivateText(`[vendors."Acme Plumbng"]\ntin = "${FAKE_TIN}"\n`, 'p.toml', book)),
        errorOf(() => parsePrivateText(`[company]\ntin = 123456789\n`, 'p.toml', book)),
    ];
    check('Private: unknown key is an error naming the key', msgs[0] !== null && msgs[0].includes('unknown key "ssn"') && msgs[0].includes('Did you mean "tin"?'), msgs[0]);
    check('Private: invalid TOML is reported without quoting the line', msgs[1] !== null && msgs[1].includes('is not valid TOML'), msgs[1]);
    check('Private: duplicate key is reported', msgs[2] !== null && msgs[2].includes('is not valid TOML'), msgs[2]);
    check('Private: unknown vendor gets the closest vendor name', msgs[3] !== null && msgs[3].includes('Did you mean "Acme Plumbing"?'), msgs[3]);
    check('Private: a number instead of a string is an error', msgs[4] !== null && msgs[4].includes('must be a string in quotes'), msgs[4]);
    check('Private: no error message ever contains the tax ID', msgs.every(m => m !== null && !m.includes(FAKE_TIN) && !m.includes('123456789')), JSON.stringify(msgs));
}

function writeFile(rel, text) {
    const file = path.join(TMP, rel);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, text);
    return file;
}

function testFindingBooks() {
    // path form
    const file = writeFile('direct/acme.book.toml', VALID);
    check('Path form: loads the book', loadBook(file).id === 'acme');
    check('Path form: relative workbook resolves against the book folder', loadBook(file).workbookPath === path.join(TMP, 'direct', 'Acme_2026.xlsx'));
    rejects('Path form: a missing .toml file is a clear error', () => loadBook(path.join(TMP, 'direct', 'nope.book.toml')), 'not found');

    // registry in the current folder (./books.toml), paths relative to the registry
    writeFile('reg/books.toml', 'acme = "../direct/acme.book.toml"\n');
    check('Registry ./books.toml: id resolves', loadBook('acme', { cwd: path.join(TMP, 'reg'), env: {} }).id === 'acme');

    // registry named by LLC_BOOKS, from a different cwd
    const regFile = writeFile('elsewhere/my-books.toml', 'acme = "../direct/acme.book.toml"\nacmee = "../direct/acme.book.toml"\n');
    check('Registry via LLC_BOOKS: id resolves from any folder', loadBook('acme', { cwd: TMP, env: { LLC_BOOKS: regFile } }).id === 'acme');
    rejects('Registry: an unknown id suggests the closest and lists the known ones', () => loadBook('acm', { cwd: TMP, env: { LLC_BOOKS: regFile } }),
        'is not in the registry', 'Did you mean "acme"?', 'Known ids: acme, acmee');
    rejects('Registry: id and the book file\'s own id must agree', () => loadBook('acmee', { cwd: TMP, env: { LLC_BOOKS: regFile } }),
        'Registry id "acmee" does not match id "acme"');
    rejects('No registry and not a file: explains both fixes', () => loadBook('acme', { cwd: path.join(TMP, 'empty-folder-that-does-not-exist'), env: {} }),
        'is not a book file', 'books.toml', 'LLC_BOOKS');
    writeFile('badreg/books.toml', 'acme = 5\n');
    rejects('Registry: a non-string path is an error', () => loadBook('acme', { cwd: path.join(TMP, 'badreg'), env: {} }), 'must be the path of a book file');
}

function testMerge() {
    const book = parse(VALID);
    const merged = mergeBookArgs(['--book', 'acme', '--vendor'], book);
    check('Merge: --book removed, workbook first, default_flags then CLI flags, book year added',
        merged.join(' ') === `${book.workbookPath} --pl --bs --vendor --year=2026`, merged.join(' '));
    check('Merge: --year=YYYY on the command line wins', mergeBookArgs(['--book=acme', '--year=2025'], book).join(' ') === `${book.workbookPath} --pl --bs --year=2025`);
    check('Merge: "--year 2025" (two arguments) also wins', !mergeBookArgs(['--book', 'acme', '--year', '2025'], book).some(a => a === '--year=2026'));
    check('Merge: a book without a year adds no --year', !mergeBookArgs(['--book', 'x'], { ...book, year: null }).some(a => a.startsWith('--year')));
    rejects('Merge: a workbook path with --book is an error', () => mergeBookArgs(['--book', 'acme', 'Other.xlsx'], book), '--book already names the workbook');
    check('Merge: the value of --details is not mistaken for a workbook', errorOf(() => mergeBookArgs(['--book', 'acme', '--details', 'Sales'], book)) === null);
}

// --- End to end -------------------------------------------------------------

function run(script, args, options = {}) {
    const r = spawnSync(process.execPath, [path.join(ROOT, script), ...args], { cwd: options.cwd || ROOT, env: { ...process.env, ...(options.env || {}) }, encoding: 'utf8' });
    return { code: r.status, out: r.stdout + r.stderr };
}

function testEndToEnd() {
    // The documented commands: copy the example files, make the workbook, report on it.
    const dir = path.join(TMP, 'example');
    fs.mkdirSync(dir);
    ['example.book.toml', 'example.private.toml'].forEach(f => fs.copyFileSync(path.join(ROOT, 'examples', f), path.join(dir, f)));
    const bookFile = path.join(dir, 'example.book.toml');

    const gen = run('generate_excel.js', ['--book', bookFile]);
    check('generate_excel --book creates the workbook named in the book', gen.code === 0 && fs.existsSync(path.join(dir, 'Example_2026.xlsx')), gen.out);
    const again = run('generate_excel.js', ['--book', bookFile]);
    check('generate_excel --book refuses to overwrite an existing workbook', again.code === 1 && again.out.includes('not overwriting'), again.out);

    // default_flags (--pl --bs) and the book's year apply when the command line says nothing
    const r = run('report.js', ['--book', bookFile]);
    check('Example book: exit 0 on the empty generated workbook', r.code === 0, r.out);
    check('Example book: the book year (2026) is the tax year', r.out.includes('Tax year: 2026'));
    check('Example book: default_flags --pl and --bs are on', r.out.includes('--- PROFIT & LOSS ---') && r.out.includes('--- BALANCE SHEET ---'));
    check('Example book: --vendor was not on by default', !r.out.includes('VENDOR SPENDING'));
    check('Example book: stamp lists the book file as an input', /Input: Book\s+: example\.book\.toml\s+sha256 [0-9a-f]{12}/.test(r.out), r.out);
    check('Example book: the private file is not read without --1099', !r.out.includes('Private file'));
    check('Example book: no vendor.xlsx warning in book mode', !r.out.includes('vendor.xlsx'));

    // the command line wins for --year, and flags add to default_flags
    const y = run('report.js', ['--book', bookFile, '--year=2025', '--vendor']);
    check('--year=2025 on the command line overrides the book year', y.out.includes('Tax year: 2025') && !y.out.includes('Tax year: 2026'), y.out);
    check('CLI flags are added to default_flags', y.out.includes('--- PROFIT & LOSS ---') && y.out.includes('VENDOR'), y.out);

    // --1099 reads the private file: it is listed in the stamp (hash only) and no tax ID is printed
    const t = run('report.js', ['--book', bookFile, '--1099']);
    check('--1099 lists the private file in the stamp', /Input: Private file\s+: example\.private\.toml\s+sha256 [0-9a-f]{12}/.test(t.out), t.out);
    check('--1099 never prints a tax ID', !t.out.includes('00-0000000') && !t.out.includes('00-0000001'), t.out);
    check('--1099 says the payer tax ID is on file', t.out.includes('Tax ID:  on file'), t.out);

    // registry by id through LLC_BOOKS, and ./books.toml in the current folder
    const reg = writeFile('example/books.toml', 'example = "example.book.toml"\n');
    const byEnv = run('report.js', ['--book', 'example'], { cwd: TMP, env: { LLC_BOOKS: reg } });
    check('--book <id> through LLC_BOOKS', byEnv.code === 0 && byEnv.out.includes('Input: Book'), byEnv.out);
    const byCwd = run('report.js', ['--book=example'], { cwd: dir, env: { LLC_BOOKS: '' } });
    check('--book=<id> through ./books.toml in the current folder', byCwd.code === 0 && byCwd.out.includes('Input: Book'), byCwd.out);

    // problems stop the run before the workbook is opened, all listed, exit 1
    const broken = writeFile('example/broken.book.toml', fs.readFileSync(bookFile, 'utf8').replace('workbook =', 'workbok =').replace('type = "cc"', 'type = "ccc"'));
    const b = run('report.js', ['--book', broken]);
    check('A bad book exits 1 and lists every problem with a fix', b.code === 1 && b.out.includes('workbok') && b.out.includes('Did you mean "workbook"?') && b.out.includes('type "ccc"') && !b.out.includes('PROFIT'), b.out);
    const noValue = run('report.js', ['--book']);
    check('--book without a value exits 1', noValue.code === 1 && noValue.out.includes('--book needs'), noValue.out);

    // a missing worksheet is reported with the closest name
    const wbFile = path.join(dir, 'Example_2026.xlsx');
    const ExcelJS = require('exceljs');
    return new ExcelJS.Workbook().xlsx.readFile(wbFile).then(async wb => {
        wb.getWorksheet('Visa Card').name = 'Visa Crd';
        await wb.xlsx.writeFile(wbFile);
        const m = run('report.js', ['--book', bookFile]);
        check('A sheet missing from the workbook exits 1 with the closest worksheet name',
            m.code === 1 && m.out.includes('Sheet "Visa Card" is not in the workbook') && m.out.includes('Did you mean "Visa Crd"?'), m.out);
    });
}

async function main() {
    testSuggestions();
    testValidation();
    testPrivate();
    testFindingBooks();
    testMerge();
    await testEndToEnd();

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll book config tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

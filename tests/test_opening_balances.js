// Tests --opening-balances (lib/opening-balances.js and its use in report.js):
//   1. Parsing and validating the TOML file on its own
//   2. Checking the file against the workbook's Setup (categories, sheet links, year)
//   3. End to end: a linked bank account, an unlinked liability, an unlinked asset
//      and an equity account, with exact balance-sheet numbers
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { parseOpeningBalancesText, resolveOpeningBalances } = require('../lib/opening-balances');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-opening-'));
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

// Passes if parsing `text` throws an error mentioning every fragment.
function rejects(label, text, ...fragments) {
    const msg = errorOf(() => parseOpeningBalancesText(text, 'test.toml'));
    check(label, msg !== null && fragments.every(f => msg.includes(f)), `got ${msg === null ? 'no error' : JSON.stringify(msg)}`);
}

// Returns the number printed after "label :" on the first matching line (last number on the line).
function valueFor(output, label) {
    const line = output.split('\n').find(l => new RegExp(`^\\s*${label}\\s*:`).test(l));
    if (!line) return null;
    const nums = line.match(/-?[\d,]+\.\d{2}/g);
    return nums ? parseFloat(nums[nums.length - 1].replace(/,/g, '')) : null;
}

const VALID = 'year = 2025\nas_of_date = 2025-12-31\n[accounts]\n"Mortgage" = 3433.00\n"Solar Loan" = 4345\n';

function testParsing() {
    const ok = parseOpeningBalancesText(VALID, 'test.toml');
    check('Valid file: year', ok.year === 2025);
    check('Valid file: as_of_date is a YYYY-MM-DD string', ok.asOfDate === '2025-12-31', `got ${ok.asOfDate}`);
    check('Valid file: accounts keep order, name and amount',
        ok.accounts.length === 2 && ok.accounts[0].key === 'Mortgage' && ok.accounts[0].amount === 3433
        && ok.accounts[1].key === 'Solar Loan' && ok.accounts[1].amount === 4345);

    const noDate = parseOpeningBalancesText('year = 2025\n[accounts]\nA = 1\n', 'test.toml');
    check('as_of_date is optional', noDate.asOfDate === null);

    const quoted = parseOpeningBalancesText('year = 2025\nas_of_date = "2026-01-01"\n[accounts]\nA = 1\n', 'test.toml');
    check('Quoted-string as_of_date (Jan 1 of year + 1) is accepted', quoted.asOfDate === '2026-01-01');

    const mixed = parseOpeningBalancesText('year = 2025\n[accounts]\nInt = 5\nFloat = 12.5\nNeg = -20.25\n', 'test.toml');
    check('Integer, float and negative values are accepted',
        mixed.accounts[0].amount === 5 && mixed.accounts[1].amount === 12.5 && mixed.accounts[2].amount === -20.25);

    rejects('Duplicate key is rejected', 'year = 2025\n[accounts]\nA = 1\nA = 2\n', 'not valid TOML');
    rejects('Unknown top-level key is rejected', 'year = 2025\nbogus = 1\n[accounts]\nA = 1\n', 'Unknown top-level key "bogus"');
    rejects('[acounts] typo is rejected (unknown key and missing table)', 'year = 2025\n[acounts]\nA = 1\n',
        'Unknown top-level key "acounts"', '[accounts] table is missing');
    rejects('Missing year is rejected', '[accounts]\nA = 1\n', '"year" is missing');
    rejects('3-digit year is rejected', 'year = 999\n[accounts]\nA = 1\n', '"year" must be a 4-digit integer');
    rejects('Non-integer year is rejected', 'year = 2025.5\n[accounts]\nA = 1\n', '"year" must be a 4-digit integer');
    rejects('Bad as_of_date (not a date) is rejected', 'year = 2025\nas_of_date = "soon"\n[accounts]\nA = 1\n', '"as_of_date" must be a real date');
    rejects('Impossible as_of_date (Feb 30) is rejected', 'year = 2025\nas_of_date = "2025-02-30"\n[accounts]\nA = 1\n', '"as_of_date" must be a real date');
    rejects('as_of_date not matching the year is rejected', 'year = 2025\nas_of_date = 2024-12-31\n[accounts]\nA = 1\n', 'must be 2025-12-31 or 2026-01-01');
    rejects('Empty [accounts] is rejected', 'year = 2025\n[accounts]\n', '[accounts] table is empty');
    rejects('String value is rejected', 'year = 2025\n[accounts]\nA = "3433"\n', 'Account "A" must be a number');
    rejects('Boolean value is rejected', 'year = 2025\n[accounts]\nA = true\n', 'Account "A" must be a number');
    rejects('3-decimal value is rejected', 'year = 2025\n[accounts]\nA = 1.005\n', 'Account "A" has more than 2 decimals');
    rejects('Nested table in [accounts] is rejected', 'year = 2025\n[accounts]\nA = 1\n[accounts.B]\nC = 2\n', 'Account "B" must be a number, got a table');
    rejects('Infinite value is rejected', 'year = 2025\n[accounts]\nA = inf\n', 'Account "A" must be a finite number');

    const all = errorOf(() => parseOpeningBalancesText('bogus = 1\n[accounts]\nA = "x"\nB = 1.234\n', 'test.toml'));
    check('Every problem is listed in one error',
        all !== null && all.includes('4 problem(s)') && ['bogus', '"year" is missing', 'Account "A"', 'Account "B"'].every(f => all.includes(f)),
        `got ${JSON.stringify(all)}`);
}

// A ctx-like object: Setup categories plus sheet configs, as the report phases build them.
function makeCtx(overrides = {}) {
    const cat = (displayName, report, accountType) => [displayName.toLowerCase(), { report, accountType, displayName }];
    return {
        taxYear: 2026,
        uniqueCategories: new Map([
            cat('Bank', 'Balance Sheet', 'Asset'),
            cat('Mortgage', 'Balance Sheet', 'Liability'),
            cat('Solar Loan', 'Balance Sheet', 'Liability'),
            cat('Equipment', 'Balance Sheet', 'Asset'),
            cat('Owner Investment', 'Balance Sheet', 'Equity'),
            cat('Sales', 'P&L', 'Income'),
        ]),
        sheetConfigs: [{ name: 'Bank', linkedAccount: 'Bank', startBalance: 0 }],
        openingBalances: [],
        ...overrides,
    };
}

function resolveError(toml, ctx = makeCtx()) {
    return errorOf(() => resolveOpeningBalances(parseOpeningBalancesText(toml, 'test.toml'), ctx));
}

function testResolution() {
    const ctx = makeCtx();
    const res = resolveOpeningBalances(parseOpeningBalancesText(
        'year = 2025\n[accounts]\nbank = 1000\n"Solar Loan" = 4345\nEquipment = 200\n"owner_investment" = 343\n', 'test.toml'), ctx);
    check('Linked account sets its sheet startBalance', ctx.sheetConfigs[0].startBalance === 1000);
    check('Unlinked accounts are recorded on ctx.openingBalances',
        ctx.openingBalances.length === 3 && ctx.openingBalances.every(o => o.label !== 'Bank'));
    check('Normalized key "owner_investment" matches "Owner Investment"', ctx.openingBalances.some(o => o.label === 'Owner Investment' && o.amount === 343));
    check('Plug = assets - liabilities - equity (1000 + 200 - 4345 - 343)', res.plug === -3488, `got ${res.plug}`);
    check('Result reports the account count', res.count === 4);

    let msg = resolveError('year = 2025\n[accounts]\nNope = 5\n');
    check('Unknown category is rejected', msg !== null && msg.includes('Account "Nope" does not match any Setup category'), msg);

    msg = resolveError('year = 2025\n[accounts]\nSales = 5\n');
    check('P&L category is rejected', msg !== null && msg.includes('Account "Sales"') && msg.includes('P&L category'), msg);

    msg = resolveError('year = 2025\n[accounts]\n"Solar Loan" = 5\nsolar_loan = 6\n');
    check('Two keys naming the same category are rejected', msg !== null && msg.includes('"Solar Loan" and "solar_loan" are both "Solar Loan"'), msg);

    msg = resolveError('year = 2024\n[accounts]\nMortgage = 5\n');
    check('Year other than --year minus 1 is rejected', msg !== null && msg.includes('File year 2024 must be the year before --year=2026'), msg);

    msg = resolveError('year = 2025\n[accounts]\nBank = 1000\n', makeCtx({ sheetConfigs: [{ name: 'Bank', linkedAccount: 'Bank', startBalance: 500 }] }));
    check('Linked account with a different StartBalance is rejected', msg !== null && msg.includes('StartBalance 500.00') && msg.includes('file says 1000.00'), msg);

    const warnCtx = makeCtx({ sheetConfigs: [{ name: 'Bank', linkedAccount: 'Bank', startBalance: 1000 }] });
    const warnings = [];
    const origWarn = console.warn;
    console.warn = (m) => warnings.push(m);
    let warnErr = null;
    try { resolveOpeningBalances(parseOpeningBalancesText('year = 2025\n[accounts]\nBank = 1000\n', 'test.toml'), warnCtx); } catch (e) { warnErr = e.message; }
    console.warn = origWarn;
    check('Linked account with an equal StartBalance is a warning only', warnErr === null && warnings.some(w => w.includes('redundant')), warnErr || JSON.stringify(warnings));

    msg = resolveError('year = 2025\n[accounts]\nBank = 1\n', makeCtx({
        sheetConfigs: [{ name: 'A', linkedAccount: 'Bank', startBalance: 0 }, { name: 'B', linkedAccount: 'Bank', startBalance: 0 }],
    }));
    check('Account linked to two sheets is rejected', msg !== null && msg.includes('linked to several sheets'), msg);

    const negWarnings = [];
    console.warn = (m) => negWarnings.push(m);
    resolveOpeningBalances(parseOpeningBalancesText('year = 2025\n[accounts]\nMortgage = -50\n', 'test.toml'), makeCtx());
    console.warn = origWarn;
    check('Negative value prints a warning', negWarnings.some(w => w.includes('negative')), JSON.stringify(negWarnings));
}

// Setup uses header-row detection (see docs/setup.md). Columns A-G are the Categories table,
// H-N the SheetInfo table.
async function buildWorkbook(file) {
    const wb = new ExcelJS.Workbook();
    const setup = wb.addWorksheet('Setup');
    setup.addRow([
        'Category', 'Sub-Category', 'Account Type', 'Report', 'Transfer Account',
        'Vendors', 'Customers',
        'Sheet Name', 'Type', 'Flip', 'Offset', 'Short Name', 'Link Asset', 'Start Balance',
    ]);
    const categories = [
        ['Sales', '', 'Income', 'P&L'],
        ['Bank', '', 'Asset', 'Balance Sheet'],
        ['Mortgage', '', 'Liability', 'Balance Sheet'],
        ['Equipment', '', 'Asset', 'Balance Sheet'],
        ['Owner Investment', '', 'Equity', 'Balance Sheet'],
    ];
    categories.forEach((c, i) => c.forEach((v, j) => { setup.getRow(i + 2).getCell(j + 1).value = v; }));
    ['Bank', 'Bank', 'No', 1, 'Bank', 'Bank'].forEach((v, j) => { setup.getRow(2).getCell(8 + j).value = v; });

    const bank = wb.addWorksheet('Bank');
    bank.addRow(['Date', 'Description', 'Amount', 'Category']);
    bank.addRow([new Date('2026-03-01'), 'Mortgage payment', -500, 'Mortgage']);
    bank.addRow([new Date('2026-03-02'), 'Buy equipment', -400, 'Equipment']);
    // An out-of-year row must not matter.
    bank.addRow([new Date('2025-03-02'), 'Last year', -77, 'Equipment']);

    wb.addWorksheet('Ledger').addRow(['Date', 'Description', 'Category', 'Debit', 'Credit']);
    await wb.xlsx.writeFile(file);
}

function runReport(workbook, extraArgs) {
    const r = spawnSync(process.execPath, ['report.js', workbook, '--bs', '--year=2026', ...extraArgs], { cwd: ROOT, encoding: 'utf8' });
    return { code: r.status, output: r.stdout + r.stderr };
}

async function testEndToEnd() {
    const wbFile = path.join(TMP, 'opening.xlsx');
    await buildWorkbook(wbFile);
    const tomlFile = path.join(TMP, 'opening.toml');
    fs.writeFileSync(tomlFile, 'year = 2025\nas_of_date = 2025-12-31\n[accounts]\nBank = 1000\nMortgage = 3433.00\nEquipment = 4345\n"Owner Investment" = 343\n');

    for (const form of [['--opening-balances', tomlFile], [`--opening-balances=${tomlFile}`]]) {
        const r = runReport(wbFile, form);
        const out = r.output;
        const tag = form.length === 2 ? 'space form' : '= form';
        check(`End to end (${tag}): exit 0`, r.code === 0, `exit ${r.code}\n${out}`);
        check(`End to end (${tag}): Mortgage = 3433 - 500 = 2933`, valueFor(out, 'Mortgage') === 2933, `got ${valueFor(out, 'Mortgage')}`);
        check(`End to end (${tag}): Equipment = 4345 + 400 = 4745`, valueFor(out, 'Equipment') === 4745, `got ${valueFor(out, 'Equipment')}`);
        check(`End to end (${tag}): Bank = 1000 - 500 - 400 = 100`, valueFor(out, 'Bank') === 100, `got ${valueFor(out, 'Bank')}`);
        check(`End to end (${tag}): Owner Investment = 343`, valueFor(out, 'Owner Investment') === 343, `got ${valueFor(out, 'Owner Investment')}`);
        check(`End to end (${tag}): plug = 1000 + 4345 - 3433 - 343 = 1569`,
            valueFor(out, 'Retained Earnings \\(Opening Balance\\)') === 1569, `got ${valueFor(out, 'Retained Earnings \\(Opening Balance\\)')}`);
        check(`End to end (${tag}): balance sheet balances`, out.includes('[OK] (A = L + E)'));
        check(`End to end (${tag}): summary line printed`,
            out.includes('Opening balances: 4 accounts from opening.toml (closing year 2025). Plug to Retained Earnings (Opening Balance): 1,569.00'));
    }

    // Without the flag nothing changes: the unlinked accounts hold only this year's activity.
    const plain = runReport(wbFile, []);
    check('Without the flag: Mortgage = -500 (unchanged behavior)', valueFor(plain.output, 'Mortgage') === -500, `got ${valueFor(plain.output, 'Mortgage')}`);

    // A bad file stops the run with exit 1 and lists every problem.
    const badFile = path.join(TMP, 'bad.toml');
    fs.writeFileSync(badFile, 'year = 999\n[accounts]\nMortgage = "3433"\nEquipment = 1.005\n');
    const bad = runReport(wbFile, ['--opening-balances', badFile]);
    check('Bad file: exit 1', bad.code === 1, `exit ${bad.code}`);
    check('Bad file: every problem is printed',
        bad.output.includes('"year" must be a 4-digit integer') && bad.output.includes('Account "Mortgage" must be a number')
        && bad.output.includes('Account "Equipment" has more than 2 decimals'), bad.output);
    check('Bad file: no balance sheet is printed', !bad.output.includes('BALANCE SHEET'));

    // A file that parses but does not fit the workbook also stops the run.
    const misfit = path.join(TMP, 'misfit.toml');
    fs.writeFileSync(misfit, 'year = 2024\n[accounts]\nNope = 1\nSales = 2\n');
    const mis = runReport(wbFile, ['--opening-balances', misfit]);
    check('Setup-stage problems: exit 1 and all listed', mis.code === 1
        && mis.output.includes('File year 2024 must be the year before --year=2026')
        && mis.output.includes('Account "Nope" does not match')
        && mis.output.includes('Account "Sales"'), `exit ${mis.code}\n${mis.output}`);

    const missing = runReport(wbFile, ['--opening-balances', path.join(TMP, 'absent.toml')]);
    check('Missing file: exit 1 with a clear message', missing.code === 1 && missing.output.includes('not found'), missing.output);

    const noValue = spawnSync(process.execPath, ['report.js', wbFile, '--year=2026', '--opening-balances'], { cwd: ROOT, encoding: 'utf8' });
    check('Flag without a path: exit 1', noValue.status === 1 && (noValue.stdout + noValue.stderr).includes('needs a file path'));
}

async function main() {
    testParsing();
    testResolution();
    await testEndToEnd();

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });

    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll opening-balances tests passed.');
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

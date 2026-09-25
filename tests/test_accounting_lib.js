const assert = require('assert');
const { parseAmount, parseTaxYear, getYear, get1099Threshold, parseDateUTC, DEFAULT_TAX_YEAR } = require('../lib/accounting');

let failures = 0;
function test(name, fn) {
    try {
        fn();
        console.log(`✅ [PASS] ${name}`);
    } catch (e) {
        console.error(`❌ [FAIL] ${name}: ${e.message}`);
        failures++;
    }
}

test('parseAmount: numbers pass through', () => {
    assert.strictEqual(parseAmount(12.5), 12.5);
    assert.strictEqual(parseAmount(-3), -3);
});
test('parseAmount: blanks are zero', () => {
    for (const v of [null, undefined, '', '  ', '-']) assert.strictEqual(parseAmount(v), 0);
});
test('parseAmount: currency symbols and thousands separators', () => {
    assert.strictEqual(parseAmount('$1,234.56'), 1234.56);
    assert.strictEqual(parseAmount('1,234.56'), 1234.56);
    assert.strictEqual(parseAmount('-$5.00'), -5);
});
test('parseAmount: accounting-style negatives', () => {
    assert.strictEqual(parseAmount('(1,234.56)'), -1234.56);
    assert.strictEqual(parseAmount('($50)'), -50);
    assert.strictEqual(parseAmount('1,234.56-'), -1234.56);
    assert.strictEqual(parseAmount('100.00 CR'), -100);
});
test('parseAmount: text is NaN, not zero', () => {
    assert.ok(Number.isNaN(parseAmount('abc')));
    assert.ok(Number.isNaN(parseAmount('12abc')));
});
test('parseAmount: Excel formula results', () => {
    assert.strictEqual(parseAmount({ formula: 'A1+A2', result: 42 }), 42);
});

test('parseTaxYear: default', () => {
    assert.strictEqual(parseTaxYear([]), DEFAULT_TAX_YEAR);
    assert.strictEqual(DEFAULT_TAX_YEAR, 2026);
});
test('parseTaxYear: --year=YYYY and --year YYYY', () => {
    assert.strictEqual(parseTaxYear(['file.xlsx', '--year=2025']), 2025);
    assert.strictEqual(parseTaxYear(['--year', '2024', 'file.xlsx']), 2024);
});
test('parseTaxYear: rejects bad values', () => {
    assert.throws(() => parseTaxYear(['--year=25']));
    assert.throws(() => parseTaxYear(['--year']));
});

test('getYear: Date, Excel serial, ISO and US strings', () => {
    assert.strictEqual(getYear(new Date(Date.UTC(2025, 11, 31))), 2025);
    assert.strictEqual(getYear(new Date(Date.UTC(2026, 0, 1))), 2026);
    assert.strictEqual(getYear(45658), 2025); // 2025-01-01
    assert.strictEqual(getYear('2025-12-31'), 2025);
    assert.strictEqual(getYear('01/03/2026'), 2026);
    assert.strictEqual(getYear('1/3/25'), 2025);
    assert.strictEqual(getYear(''), null);
    assert.strictEqual(getYear('not a date'), null);
});

test('parseDateUTC: keeps the calendar day', () => {
    assert.strictEqual(parseDateUTC('2025-01-03').toISOString(), '2025-01-03T00:00:00.000Z');
    assert.strictEqual(parseDateUTC('12/31/2025').toISOString(), '2025-12-31T00:00:00.000Z');
    assert.strictEqual(parseDateUTC('1/1/26').toISOString(), '2026-01-01T00:00:00.000Z');
});

test('get1099Threshold: NEC/MISC by year, INT', () => {
    assert.strictEqual(get1099Threshold('NEC', 2025), 600);
    assert.strictEqual(get1099Threshold('MISC', 2025), 600);
    assert.strictEqual(get1099Threshold('NEC', 2026), 2000);
    assert.strictEqual(get1099Threshold('MISC', 2027), 2000);
    assert.strictEqual(get1099Threshold('INT', 2026), 0);
});

if (failures > 0) {
    console.error(`\n${failures} test(s) failed.`);
    process.exit(1);
}
console.log('\nAll accounting helper tests passed.');

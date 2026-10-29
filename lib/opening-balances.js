// Opening balances: last year's closing balance-sheet balances, read from a
// TOML file (--opening-balances <path>) and applied to the balance sheet.
// The file is parsed and checked on its own first (parseOpeningBalancesText),
// then checked against the workbook's Setup (resolveOpeningBalances).
// Every problem is collected so one run lists them all.
const fs = require('fs');
const { parse: parseToml } = require('smol-toml');
const { classifyBalanceSheetType } = require('./report/account-class');

const KNOWN_KEYS = ['year', 'as_of_date', 'accounts'];
const YEAR_PATTERN = /^\d{4}$/;
const DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})$/;
const DATE_LENGTH = 10; // length of "YYYY-MM-DD"
const CENTS_PER_UNIT = 100;
const CENT_EPSILON = 1e-9; // tolerance when checking a value has at most 2 decimals
const AMOUNT_EPSILON = 0.005; // amounts closer than half a cent are equal
const SEPARATOR_RUNS = /[\s_-]+/g;

function isTomlDate(v) {
    return v !== null && typeof v === 'object' && typeof v.isDate === 'function' && typeof v.toISOString === 'function';
}

function isTable(v) {
    return v !== null && typeof v === 'object' && !Array.isArray(v) && !isTomlDate(v);
}

function describe(v) {
    if (isTomlDate(v)) return 'a date';
    if (Array.isArray(v)) return 'an array';
    if (isTable(v)) return 'a table';
    if (typeof v === 'string') return 'a string';
    if (typeof v === 'boolean') return 'a boolean';
    return typeof v;
}

// A TOML number as a JS number. smol-toml returns numbers (and throws on an
// integer too large to hold exactly), but accept a BigInt in case it is ever
// switched on. Returns null for anything that is not a number.
function toNumber(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'bigint') return Number(v);
    return null;
}

// "YYYY-MM-DD" if y/m/d is a real calendar date, else null.
function realDate(y, m, d) {
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return dt.toISOString().slice(0, DATE_LENGTH);
}

// as_of_date as "YYYY-MM-DD", from a TOML date or a quoted string; null if invalid.
function readDate(v) {
    if (isTomlDate(v)) return v.isDate() ? v.toISOString().slice(0, DATE_LENGTH) : null;
    if (typeof v !== 'string') return null;
    const m = DATE_PATTERN.exec(v.trim());
    return m ? realDate(Number(m[1]), Number(m[2]), Number(m[3])) : null;
}

// Parses and validates the file text. Returns
// { year, asOfDate: 'YYYY-MM-DD' | null, accounts: [{ key, amount }] }
// or throws one Error listing every problem found.
function parseOpeningBalancesText(text, sourceName) {
    const label = sourceName || 'opening balances';
    let raw;
    try {
        raw = parseToml(text);
    } catch (e) {
        throw new Error(`Opening balances file "${label}" is not valid TOML:\n  - ${(e.message || '').split('\n')[0]}`);
    }

    const problems = [];
    Object.keys(raw).filter(k => !KNOWN_KEYS.includes(k)).forEach(k => {
        problems.push(`Unknown top-level key "${k}". Allowed: ${KNOWN_KEYS.join(', ')}.`);
    });

    let year = null;
    const yearNum = toNumber(raw.year);
    if (raw.year === undefined) problems.push('"year" is missing. Set it to the year that closed, e.g. year = 2025.');
    else if (yearNum === null || !Number.isInteger(yearNum) || !YEAR_PATTERN.test(String(yearNum))) problems.push(`"year" must be a 4-digit integer, got ${JSON.stringify(String(raw.year))}.`);
    else year = yearNum;

    let asOfDate = null;
    if (raw.as_of_date !== undefined) {
        asOfDate = readDate(raw.as_of_date);
        if (asOfDate === null) {
            problems.push(`"as_of_date" must be a real date (YYYY-MM-DD), got ${JSON.stringify(String(raw.as_of_date))}.`);
        } else if (year !== null && asOfDate !== `${year}-12-31` && asOfDate !== `${year + 1}-01-01`) {
            problems.push(`"as_of_date" ${asOfDate} must be ${year}-12-31 or ${year + 1}-01-01 to match year ${year}.`);
            asOfDate = null;
        }
    }

    const accounts = [];
    if (raw.accounts === undefined) {
        problems.push('The [accounts] table is missing.');
    } else if (!isTable(raw.accounts)) {
        problems.push(`"accounts" must be a table ([accounts]), got ${describe(raw.accounts)}.`);
    } else if (Object.keys(raw.accounts).length === 0) {
        problems.push('The [accounts] table is empty.');
    } else {
        for (const [key, value] of Object.entries(raw.accounts)) {
            if (!key.trim()) {
                problems.push('An account has an empty name.');
                continue;
            }
            const amount = toNumber(value);
            if (amount === null) {
                problems.push(`Account "${key}" must be a number, got ${describe(value)}${typeof value === 'string' ? ` ("${value}"); remove the quotes` : ''}.`);
            } else if (!Number.isFinite(amount)) {
                problems.push(`Account "${key}" must be a finite number.`);
            } else if (Math.abs(Math.round(amount * CENTS_PER_UNIT) / CENTS_PER_UNIT - amount) > CENT_EPSILON) {
                problems.push(`Account "${key}" has more than 2 decimals (${amount}).`);
            } else {
                accounts.push({ key, amount });
            }
        }
    }

    if (problems.length > 0) {
        throw new Error(`Opening balances file "${label}" has ${problems.length} problem(s):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }
    return { year, asOfDate, accounts };
}

// Reads and validates the file at filePath. A missing file is a clear error.
function loadOpeningBalancesFile(filePath) {
    if (!filePath) throw new Error('--opening-balances needs a file path.');
    if (!fs.existsSync(filePath)) throw new Error(`Opening balances file "${filePath}" not found.`);
    return parseOpeningBalancesText(fs.readFileSync(filePath, 'utf8'), filePath);
}

// Lowercase, with runs of space, "_" and "-" treated as one separator.
function normalizeKey(key) {
    return key.trim().toLowerCase().replace(SEPARATOR_RUNS, ' ');
}

// Finds the Setup category for a key: exact (case-insensitive) first, then
// normalized, which must be unique. Returns { category } or { error }.
function matchCategory(key, uniqueCategories) {
    const exact = uniqueCategories.get(key.trim().toLowerCase());
    if (exact) return { category: exact };
    const wanted = normalizeKey(key);
    const hits = Array.from(uniqueCategories.values()).filter(c => normalizeKey(c.displayName) === wanted);
    if (hits.length === 1) return { category: hits[0] };
    if (hits.length > 1) return { error: `Account "${key}" matches several Setup categories (${hits.map(h => `"${h.displayName}"`).join(', ')}).` };
    return { error: `Account "${key}" does not match any Setup category.` };
}

// Checks the parsed file against the Setup stage (categories, sheet links,
// tax year) and applies it: a linked account's amount becomes its sheet's
// startBalance; the rest are recorded on ctx.openingBalances for buildReports.
// Throws one Error listing every problem. Returns { count, plug }, where plug
// is assets minus liabilities minus equity from the file.
function resolveOpeningBalances(parsed, ctx) {
    const { uniqueCategories, sheetConfigs, taxYear } = ctx;
    const problems = [];
    const warnings = [];

    if (parsed.year !== taxYear - 1) {
        problems.push(`File year ${parsed.year} must be the year before --year=${taxYear} (${taxYear - 1}).`);
    }

    const seen = new Map(); // category key -> file key that claimed it
    const linkedApply = []; // [{ sheet, amount }]
    const unlinked = [];
    let plug = 0;

    for (const { key, amount } of parsed.accounts) {
        const m = matchCategory(key, uniqueCategories);
        if (m.error) {
            problems.push(m.error);
            continue;
        }
        const cat = m.category;
        const catKey = cat.displayName.toLowerCase();
        if (seen.has(catKey)) {
            problems.push(`Accounts "${seen.get(catKey)}" and "${key}" are both "${cat.displayName}".`);
            continue;
        }
        seen.set(catKey, key);
        if (cat.report !== 'Balance Sheet') {
            problems.push(`Account "${key}" is "${cat.displayName}", a ${cat.report} category. Only Balance Sheet accounts carry over; P&L starts at zero each year.`);
            continue;
        }

        const kind = classifyBalanceSheetType(cat.accountType);
        plug += kind === 'asset' ? amount : -amount;
        if (amount < 0) warnings.push(`Opening balance for "${cat.displayName}" is negative (${amount.toFixed(2)}).`);

        const sheets = sheetConfigs.filter(s => s.linkedAccount && s.linkedAccount.toLowerCase() === catKey);
        if (sheets.length > 1) {
            problems.push(`Account "${key}" is linked to several sheets (${sheets.map(s => `"${s.name}"`).join(', ')}); the opening balance is ambiguous.`);
        } else if (sheets.length === 1) {
            const existing = sheets[0].startBalance || 0;
            if (Math.abs(existing) < AMOUNT_EPSILON) {
                linkedApply.push({ sheet: sheets[0], amount });
            } else if (Math.abs(existing - amount) < AMOUNT_EPSILON) {
                warnings.push(`Sheet "${sheets[0].name}" already has StartBalance ${existing.toFixed(2)} in Setup, the same as the opening balances file for "${key}". The file entry is redundant.`);
            } else {
                problems.push(`Account "${key}": sheet "${sheets[0].name}" has StartBalance ${existing.toFixed(2)} in Setup but the file says ${amount.toFixed(2)}. Remove one of them.`);
            }
        } else {
            unlinked.push({ label: cat.displayName, amount, kind });
        }
    }

    if (problems.length > 0) {
        throw new Error(`Opening balances do not fit this workbook (${problems.length} problem(s)):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }

    linkedApply.forEach(({ sheet, amount }) => { sheet.startBalance = amount; });
    ctx.openingBalances = unlinked;
    warnings.forEach(w => console.warn(`[!] WARNING: ${w}`));
    return { count: parsed.accounts.length, plug };
}

module.exports = { parseOpeningBalancesText, loadOpeningBalancesFile, resolveOpeningBalances };

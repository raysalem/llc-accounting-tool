// Parses and validates a book file (<book>.toml) and its private file
// (<book>.private.toml). Every problem is collected, each with a suggested fix,
// and thrown as one Error, so a single run lists them all.
const path = require('path');
const { parse: parseToml } = require('smol-toml');
const { describe, isTable, toNumber } = require('../opening-balances');
const { NORM_VEND } = require('../report/cells');
const { didYouMean } = require('./suggest');
const S = require('./schema');

const YEAR_PATTERN = /^\d{4}$/;
const LISTS_TAB = 'Lists'; // tabs the tool writes itself (lists.js, workbook.js)
const VERSION_TAB = 'VERSION';
const CENTS_PER_UNIT = 100;
const CENT_EPSILON = 1e-9; // tolerance when checking a value has at most 2 decimals
const MAX_LISTED_NAMES = 6; // how many valid names a message lists

// One line of a smol-toml error. Its full message quotes the offending source
// line, which in a private file could hold a tax ID, so only the first line,
// the line number and the column are used.
function tomlFailure(e) {
    const reason = (e.message || 'parse error').split('\n')[0];
    return e.line ? `${reason} (line ${e.line}, column ${e.column})` : reason;
}

function listNames(names) {
    const shown = names.slice(0, MAX_LISTED_NAMES).map(n => `"${n}"`).join(', ');
    return names.length > MAX_LISTED_NAMES ? `${shown}, ...` : shown;
}

function unknownKeyMessage(where, key, allowed) {
    if (key === S.TAX_ID_KEY) {
        return `${where}: "${key}" does not belong in the book file. Fix: put tax IDs in the private file ([company] tin = "..." and [vendors."Name"] tin = "..."), which is never committed.`;
    }
    return `${where}: unknown key "${key}".${didYouMean(key, allowed)} Allowed keys: ${allowed.join(', ')}.`;
}

// Reads one value of the given kind. Returns { value } or { problem }.
function readValue(raw, kind) {
    switch (kind) {
    case 'string':
        return typeof raw === 'string' ? { value: raw.trim() } : { problem: `must be a string in quotes, got ${describe(raw)}` };
    case 'boolean':
        return typeof raw === 'boolean' ? { value: raw } : { problem: `must be true or false (no quotes), got ${describe(raw)}` };
    case 'year': {
        const n = toNumber(raw);
        if (n === null || !Number.isInteger(n) || !YEAR_PATTERN.test(String(n))) return { problem: `must be a 4-digit year such as 2026, got ${JSON.stringify(String(raw))}` };
        return { value: n };
    }
    case 'positive': {
        const n = toNumber(raw);
        if (n === null || !Number.isInteger(n) || n < 1) return { problem: `must be a whole number of 1 or more, got ${JSON.stringify(String(raw))}` };
        return { value: n };
    }
    case 'money': {
        const n = toNumber(raw);
        if (n === null) return { problem: `must be a number${typeof raw === 'string' ? ` (remove the quotes around "${raw}")` : ''}, got ${describe(raw)}` };
        if (!Number.isFinite(n)) return { problem: 'must be a finite number' };
        if (Math.abs(Math.round(n * CENTS_PER_UNIT) / CENTS_PER_UNIT - n) > CENT_EPSILON) return { problem: `has more than 2 decimals (${n}); round to cents` };
        return { value: n };
    }
    case 'strings': {
        if (!Array.isArray(raw)) return { problem: `must be a list of strings such as ["a", "b"], got ${describe(raw)}` };
        const bad = raw.findIndex(v => typeof v !== 'string' || !v.trim());
        if (bad !== -1) return { problem: `item ${bad + 1} must be a non-empty string, got ${describe(raw[bad])}` };
        return { value: raw.map(v => v.trim()) };
    }
    default:
        throw new Error(`Unknown field kind ${kind}`);
    }
}

// Reads a table against a spec { key: { kind, required } }. Pushes a problem for
// every unknown key, wrong type or missing required key. Returns the valid values.
function readTable(table, where, spec, problems) {
    const out = {};
    const allowed = Object.keys(spec);
    for (const [key, raw] of Object.entries(table)) {
        if (!spec[key]) {
            problems.push(unknownKeyMessage(where, key, allowed));
            continue;
        }
        const r = readValue(raw, spec[key].kind);
        if (r.problem) problems.push(`${where}: "${key}" ${r.problem}.`);
        else if (spec[key].required && r.value === '') problems.push(`${where}: "${key}" must not be empty.`);
        else out[key] = r.value;
    }
    for (const [key, field] of Object.entries(spec)) {
        if (field.required && table[key] === undefined) problems.push(`${where}: "${key}" is missing. Fix: add ${key} = ...`);
    }
    return out;
}

// raw[key] as a list of tables. A missing key gives [] (and a problem when required).
function readTables(raw, key, required, problems) {
    const value = raw[key];
    if (value === undefined) {
        if (required) problems.push(`No [[${key}]] entries. Fix: add at least one [[${key}]] table.`);
        return [];
    }
    if (!Array.isArray(value) || value.some(v => !isTable(v))) {
        problems.push(`"${key}" must be written as [[${key}]] entries (one per item), got ${describe(value)}. Fix: use [[${key}]] instead of [${key}].`);
        return [];
    }
    if (value.length === 0 && required) problems.push(`No [[${key}]] entries. Fix: add at least one [[${key}]] table.`);
    return value;
}

function itemLabel(group, index, name) {
    return `[[${group}]] #${index + 1}${name ? ` "${name}"` : ''}`;
}

// Matches a value against a closed list, ignoring case. Returns { value } or { problem }.
function readChoice(raw, choices, what) {
    const hit = choices.find(c => c.toLowerCase() === String(raw).trim().toLowerCase());
    if (hit) return { value: hit };
    return { problem: `${what} "${raw}" is not valid.${didYouMean(raw, choices)} Valid: ${choices.join(', ')}.` };
}

// Flags names that appear twice. key maps a name to the form compared (default: lower case).
function checkDuplicates(items, where, problems, key = n => n.toLowerCase()) {
    const seen = new Map();
    items.forEach((name, i) => {
        const k = key(name);
        if (seen.has(k)) problems.push(`${where}: "${name}" appears twice (#${seen.get(k) + 1} and #${i + 1}). Fix: remove one, or rename it.`);
        else seen.set(k, i);
    });
}

function readBookSection(raw, problems) {
    if (!isTable(raw.book)) {
        problems.push(raw.book === undefined ? 'The [book] table is missing. Fix: add [book] with id and workbook.' : `"book" must be a table ([book]), got ${describe(raw.book)}.`);
        return {};
    }
    const b = readTable(raw.book, '[book]', S.BOOK_KEYS, problems);
    (b.default_flags || []).forEach(flag => {
        if (S.DEFAULT_FLAG_CHOICES.includes(flag)) return;
        if (flag === '--year' || flag.startsWith('--year=')) problems.push(`[book]: default_flags has "${flag}". Fix: set year = 2026 in [book] instead.`);
        else if (flag.startsWith('--book')) problems.push(`[book]: default_flags has "${flag}". Fix: remove it; the book file is the book.`);
        else problems.push(`[book]: default_flags has unknown flag "${flag}".${didYouMean(flag, S.DEFAULT_FLAG_CHOICES)} Valid: ${S.DEFAULT_FLAG_CHOICES.join(', ')}.`);
    });
    return b;
}

function readCategories(raw, problems) {
    const categories = [];
    readTables(raw, 'categories', true, problems).forEach((table, i) => {
        const where = itemLabel('categories', i, typeof table.name === 'string' ? table.name : '');
        const c = readTable(table, where, S.CATEGORY_KEYS, problems);
        if (c.type !== undefined) {
            const t = readChoice(c.type, S.CATEGORY_TYPES, 'type');
            if (t.problem) { problems.push(`${where}: ${t.problem}`); c.type = undefined; } else c.type = t.value;
        }
        if (c.name === undefined || c.type === undefined) return;
        c.subcategories = c.subcategories || [];
        checkDuplicates(c.subcategories, `${where} subcategories`, problems);
        if (c.transfer_account && c.type !== 'transfer') {
            problems.push(`${where}: "transfer_account" only applies to type = "transfer". Fix: remove it, or change type to "transfer".`);
        }
        categories.push({ name: c.name, type: c.type, subcategories: c.subcategories, transferAccount: c.transfer_account || '' });
    });
    checkDuplicates(categories.map(c => c.name), '[[categories]]', problems);
    return categories;
}

function readSheets(raw, categories, problems) {
    const byName = new Map(categories.map(c => [c.name.toLowerCase(), c]));
    const balanceSheetNames = categories.filter(c => S.BALANCE_SHEET_TYPES.includes(c.type)).map(c => c.name);
    const sheets = [];
    readTables(raw, 'sheets', true, problems).forEach((table, i) => {
        const where = itemLabel('sheets', i, typeof table.name === 'string' ? table.name : '');
        const s = readTable(table, where, S.SHEET_KEYS, problems);
        if (s.type !== undefined) {
            const t = readChoice(s.type, S.SHEET_TYPES, 'type');
            if (t.problem) { problems.push(`${where}: ${t.problem}`); s.type = undefined; } else s.type = t.value;
        }
        if (s.name === undefined || s.type === undefined) return;

        let link = null;
        if (s.type === 'ledger') {
            if (!S.LEDGER_SHEET_NAMES.includes(s.name.toLowerCase())) {
                problems.push(`${where}: a ledger sheet must be named "Ledger" or "General Ledger" (the report looks for those names). Fix: rename the worksheet and this entry.`);
            }
            ['link', 'start_balance', 'end_balance'].filter(k => table[k] !== undefined).forEach(k => {
                problems.push(`${where}: a ledger sheet has no "${k}". Fix: remove it.`);
            });
        } else {
            if (!s.link) {
                const hint = balanceSheetNames.length ? ` Balance Sheet categories: ${listNames(balanceSheetNames)}.` : ' Add an asset or liability category first.';
                problems.push(`${where}: "link" is missing. Fix: add link = "<Balance Sheet category this sheet reconciles to>".${hint}`);
            } else {
                const target = byName.get(s.link.toLowerCase());
                if (!target) {
                    problems.push(`${where}: link "${s.link}" is not a category.${didYouMean(s.link, balanceSheetNames)} Fix: add it under [[categories]] with type asset or liability, or correct the name.`);
                } else if (!S.BALANCE_SHEET_TYPES.includes(target.type)) {
                    const report = target.type === 'transfer' ? 'a Transfer' : 'a P&L';
                    problems.push(`${where}: link "${target.name}" is ${report} category (type = "${target.type}"). Fix: link to a Balance Sheet category (asset, liability or equity).${didYouMean(s.link, balanceSheetNames)}`);
                } else {
                    link = target.name;
                }
            }
            // A missing end_balance is not a book error: a year in progress has no December statement yet.
            // The report's reconciliation check prints UNRECONCILED for the sheet instead.
        }
        sheets.push({
            name: s.name,
            type: s.type,
            flip: s.flip === undefined ? null : s.flip,
            headerRow: s.header_row || 0,
            link,
            shortName: s.short_name || '',
            startBalance: s.start_balance === undefined ? 0 : s.start_balance,
            endBalance: s.end_balance === undefined ? null : s.end_balance,
        });
    });
    checkDuplicates(sheets.map(s => s.name), '[[sheets]]', problems, n => n.trim().toLowerCase());
    return sheets;
}

function readVendors(raw, problems) {
    const vendors = [];
    readTables(raw, 'vendors', false, problems).forEach((table, i) => {
        const where = itemLabel('vendors', i, typeof table.name === 'string' ? table.name : '');
        const v = readTable(table, where, S.VENDOR_KEYS, problems);
        if (v.name === undefined) return;
        let form = 'none';
        if (v.form_1099 !== undefined) {
            const f = readChoice(v.form_1099 === '' ? 'none' : v.form_1099, S.FORM_1099_VALUES, 'form_1099');
            if (f.problem) problems.push(`${where}: ${f.problem}`); else form = f.value;
        }
        vendors.push({
            name: v.name, business: v.business_name || '', recipientName: v.recipient_name || '', form1099: form,
            address: v.address || '', city: v.city || '', state: v.state || '', zip: v.zip || '',
            country: v.country || '', email: v.email || '', phone: v.phone || '',
        });
    });
    checkDuplicates(vendors.map(v => v.name), '[[vendors]]', problems, NORM_VEND);
    return vendors;
}

function readCustomers(raw, problems) {
    const customers = [];
    readTables(raw, 'customers', false, problems).forEach((table, i) => {
        const where = itemLabel('customers', i, typeof table.name === 'string' ? table.name : '');
        const c = readTable(table, where, S.CUSTOMER_KEYS, problems);
        if (c.name !== undefined) customers.push(c.name);
    });
    checkDuplicates(customers, '[[customers]]', problems, NORM_VEND);
    return customers;
}

// A vendor or customer must not share a name with a category. (Setup only warns;
// a book is strict because the names are matched without regard to the column.)
function checkNameClashes(categories, vendors, customers, problems) {
    const categoryNames = new Set(categories.map(c => c.name.toLowerCase()));
    vendors.forEach(v => {
        if (categoryNames.has(NORM_VEND(v.name))) problems.push(`[[vendors]] "${v.name}" has the same name as a category. Fix: rename the vendor (e.g. "${v.name} Inc") or the category.`);
    });
    customers.forEach(c => {
        if (categoryNames.has(NORM_VEND(c))) problems.push(`[[customers]] "${c}" has the same name as a category. Fix: rename the customer or the category.`);
    });
}

// Checks a transfer category's transfer_account names an existing Balance Sheet category.
function checkTransferAccounts(categories, problems) {
    const balanceSheet = categories.filter(c => S.BALANCE_SHEET_TYPES.includes(c.type));
    const names = balanceSheet.map(c => c.name);
    categories.filter(c => c.transferAccount).forEach(c => {
        if (!balanceSheet.some(b => b.name.toLowerCase() === c.transferAccount.toLowerCase())) {
            problems.push(`[[categories]] "${c.name}": transfer_account "${c.transferAccount}" is not a Balance Sheet category.${didYouMean(c.transferAccount, names)} Fix: use the name of the account the transfer goes to or from.`);
        }
    });
}

// Parses and validates the text of a book file. file is its absolute path (relative
// paths inside the book resolve against its folder). Returns the normalized book,
// or throws one Error listing every problem.
function parseBookText(text, file) {
    const label = path.basename(file);
    let raw;
    try {
        raw = parseToml(text);
    } catch (e) {
        throw new Error(`Book file "${label}" is not valid TOML:\n  - ${tomlFailure(e)}`);
    }

    const problems = [];
    Object.keys(raw).filter(k => !S.TOP_LEVEL_KEYS.includes(k)).forEach(k => {
        problems.push(`Unknown top-level key or table "${k}".${didYouMean(k, S.TOP_LEVEL_KEYS)} Allowed: ${S.TOP_LEVEL_KEYS.join(', ')}.`);
    });

    const b = readBookSection(raw, problems);
    let company = {};
    if (raw.company !== undefined) {
        if (isTable(raw.company)) company = readTable(raw.company, '[company]', S.COMPANY_KEYS, problems);
        else problems.push(`"company" must be a table ([company]), got ${describe(raw.company)}.`);
    }
    const categories = readCategories(raw, problems);
    const sheets = readSheets(raw, categories, problems);
    const vendors = readVendors(raw, problems);
    const customers = readCustomers(raw, problems);
    checkNameClashes(categories, vendors, customers, problems);
    checkTransferAccounts(categories, problems);

    // Blank tabs must not collide with a book sheet or a tab the tool writes.
    const taken = new Map(sheets.map(s => [s.name.toLowerCase(), `sheet "${s.name}"`]));
    taken.set(LISTS_TAB.toLowerCase(), 'the generated Lists tab');
    taken.set(VERSION_TAB.toLowerCase(), 'the VERSION tab');
    (b.blank_sheets || []).forEach(name => {
        const key = name.trim().toLowerCase();
        if (!key) problems.push('[book]: blank_sheets has an empty name.');
        else if (taken.has(key)) problems.push(`[book]: blank_sheets "${name}" has the same name as ${taken.get(key)}. Fix: rename it.`);
        else taken.set(key, `blank sheet "${name}"`);
    });

    if (problems.length > 0) {
        throw new Error(`Book file "${label}" has ${problems.length} problem(s):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }

    const dir = path.dirname(file);
    const resolveIn = p => (p ? path.resolve(dir, p) : null);
    return {
        file,
        dir,
        id: b.id,
        name: b.name || company.name || b.id,
        workbookPath: resolveIn(b.workbook),
        year: b.year || null,
        openingBalancesPath: resolveIn(b.opening_balances),
        privatePath: resolveIn(b.private),
        defaultFlags: b.default_flags || [],
        blankSheets: b.blank_sheets || [],
        company,
        sheets,
        categories,
        vendors,
        customers,
    };
}

// Parses and validates the text of a private file ([company] tin, [vendors."Name"] tin)
// against the book's vendors. Returns { companyTin, vendorTins: Map(NORM_VEND name -> tin) }.
// Problem messages name keys only: a tax ID is never put in a message.
function parsePrivateText(text, file, book) {
    const label = path.basename(file);
    let raw;
    try {
        raw = parseToml(text);
    } catch (e) {
        throw new Error(`Private file "${label}" is not valid TOML:\n  - ${tomlFailure(e)}`);
    }

    const problems = [];
    Object.keys(raw).filter(k => !S.PRIVATE_TOP_LEVEL_KEYS.includes(k)).forEach(k => {
        problems.push(`Unknown top-level key or table "${k}".${didYouMean(k, S.PRIVATE_TOP_LEVEL_KEYS)} Allowed: ${S.PRIVATE_TOP_LEVEL_KEYS.join(', ')}.`);
    });

    const spec = { [S.TAX_ID_KEY]: { kind: 'string' } };
    const readTin = (table, where) => {
        if (!isTable(table)) {
            problems.push(`${where} must be a table, got ${describe(table)}.`);
            return '';
        }
        const t = readTable(table, where, spec, problems);
        if (t.tin === '') problems.push(`${where}: "tin" must not be empty. Fix: remove the entry, or fill it in.`);
        return t.tin || '';
    };

    let companyTin = '';
    if (raw.company !== undefined) companyTin = readTin(raw.company, '[company]');

    const vendorTins = new Map();
    if (raw.vendors !== undefined) {
        if (!isTable(raw.vendors)) {
            problems.push(`"vendors" must be a table of vendors ([vendors."Name"]), got ${describe(raw.vendors)}.`);
        } else {
            const bookNames = book.vendors.map(v => v.name);
            for (const [name, table] of Object.entries(raw.vendors)) {
                const tin = readTin(table, `[vendors."${name}"]`);
                if (!book.vendors.some(v => NORM_VEND(v.name) === NORM_VEND(name))) {
                    problems.push(`[vendors."${name}"]: no vendor with this name in the book.${didYouMean(name, bookNames)} Fix: use the vendor's name exactly as in [[vendors]].`);
                } else if (tin) {
                    vendorTins.set(NORM_VEND(name), tin);
                }
            }
        }
    }

    if (problems.length > 0) {
        throw new Error(`Private file "${label}" has ${problems.length} problem(s):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }
    return { companyTin, vendorTins };
}

// Checks the book against the workbook's worksheet names. Throws one Error listing
// every sheet that is not in the workbook, with the closest worksheet name.
function checkBookAgainstWorkbook(book, worksheetNames) {
    const problems = [];
    book.sheets.filter(s => !worksheetNames.includes(s.name)).forEach(s => {
        const others = worksheetNames.filter(n => !book.sheets.some(b => b.name === n));
        problems.push(`Sheet "${s.name}" is not in the workbook.${didYouMean(s.name, others.length ? others : worksheetNames)} Fix: rename the worksheet in Excel or the name in [[sheets]]. Worksheets: ${listNames(worksheetNames)}.`);
    });
    if (problems.length > 0) {
        throw new Error(`Book file "${path.basename(book.file)}" does not fit the workbook "${path.basename(book.workbookPath)}" (${problems.length} problem(s)):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }
}

module.exports = { parseBookText, parsePrivateText, checkBookAgainstWorkbook };

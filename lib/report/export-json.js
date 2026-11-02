// --json <path>: the report's results as data, for the local site (site.js) and for
// later tools (the filed-return comparison, #4). Built from the same ctx the console
// output is printed from, so the two cannot disagree.
//
// Tax IDs are never written: a vendor only carries hasTaxId true/false.
const fs = require('fs');
const { get1099Threshold } = require('../accounting');
const { NORM_VEND } = require('./cells');

const EXPORT_VERSION = 1;
const CENT = 0.01;

const sum = rows => rows.reduce((a, r) => a + (r.value || 0), 0);

// One P&L / balance-sheet / vendor / customer row, with its per-sheet and sub-category breakdown.
function row(r) {
    return {
        label: r.label,
        value: r.value || 0,
        opening: r.opening || 0,
        add: r.add || 0,
        sub: r.sub || 0,
        sheets: r.sheets || {},
        subCats: r.subCats || {},
    };
}

// The 1099 review: every vendor with a 1099 form (or payments) and whether it is reportable and complete.
function review1099(ctx) {
    const { reports, vendor1099Map, vendorDetailsMap, taxYear } = ctx;
    return (reports.vendors || []).map(v => {
        const info = vendor1099Map.get(NORM_VEND(v.label)) || { type: '', req: '' };
        const d = vendorDetailsMap.get(NORM_VEND(v.label)) || null;
        const form = info.type || '';
        const threshold = form ? get1099Threshold(form, taxYear) : null;
        const reportable = Boolean(form) && v.value > 0 && v.value >= threshold;
        const missing = [];
        if (reportable) {
            if (!d) missing.push('vendor details');
            else {
                if (!d.ssn) missing.push('tax ID');
                ['address', 'city', 'state', 'zip'].forEach(k => { if (!d[k]) missing.push(k); });
            }
        }
        return {
            vendor: v.label,
            paid: v.value,
            form: form || null,
            threshold,
            reportable,
            hasTaxId: Boolean(d && d.ssn),
            missing,
        };
    });
}

// The checks the site shows, each { id, ok, title, items: [text] }.
function checks(ctx) {
    const { reports } = ctx;
    const out = [];
    const add = (id, title, items) => out.push({ id, title, ok: items.length === 0, items });

    const a = sum(reports.assets || []);
    const le = sum(reports.liabilities || []) + sum(reports.equity || []);
    add('equation', 'Balance sheet balances (A = L + E)',
        Math.abs(a - le) < CENT ? [] : [`Assets ${a.toFixed(2)} vs liabilities + equity ${le.toFixed(2)} (difference ${(a - le).toFixed(2)})`]);

    add('reconciliation', 'Bank and card sheets match their End Balance',
        ((ctx.reconciliation && ctx.reconciliation.details) || [])
            .filter(d => d.expected === null || Math.abs(d.calculated - d.expected) > CENT)
            .map(d => `${d.sheet}: calculated ${d.calculated.toFixed(2)}, statement ${d.expected === null ? 'End Balance missing' : d.expected.toFixed(2)}`));

    add('ledger', 'Ledger debits equal credits',
        Math.abs(ctx.ledgerValidationTotal || 0) > CENT ? [`Ledger off by ${ctx.ledgerValidationTotal.toFixed(2)} (debits minus credits)`] : []);

    add('taxYear', `Every row is dated ${ctx.taxYear}`,
        (ctx.outOfYearRows || []).map(o => `${o.sheet} row ${o.row}: ${o.date instanceof Date ? o.date.toISOString().slice(0, 10) : o.date}`));

    add('uncategorized', 'Every row has a category',
        (ctx.uncategorizedDetails || []).map(x => `${x.sheet} row ${x.row}: no category (${x.desc || ''})`));

    const named = (list, what) => list.map(x => `${x.sheet} row ${x.row}: ${what} "${x.value}"`);
    add('categories', 'Every category is in the book', named(ctx.illegalCategories || [], 'unknown category'));
    add('subCategories', 'Every sub-category belongs to its category', named(ctx.illegalSubCategories || [], 'unknown sub-category'));
    add('vendors', 'Every vendor is known', named(ctx.illegalVendors || [], 'unknown vendor'));
    add('customers', 'Every customer is known', named(ctx.illegalCustomers || [], 'unknown customer'));

    const assetsNegative = (reports.assets || []).filter(r => r.value < -CENT).map(r => `${r.label} is ${r.value.toFixed(2)}`);
    add('negativeAssets', 'No asset is negative', assetsNegative);

    add('vendors1099', '1099 vendors have complete details',
        review1099(ctx).filter(r => r.missing.length).map(r => `${r.vendor}: missing ${r.missing.join(', ')}`));
    return out;
}

function buildExport(ctx, logLines = []) {
    const { reports } = ctx;
    const list = checks(ctx);
    return {
        exportVersion: EXPORT_VERSION,
        book: ctx.book ? { id: ctx.book.id, name: ctx.book.name || ctx.book.id, file: ctx.book.file } : null,
        workbook: ctx.filename,
        taxYear: ctx.taxYear,
        stamp: ctx.stamp,
        sheets: (ctx.reportSheetList || []).map(name => ({ name, short: (ctx.sheetNameMap || {})[name] || name })),
        netIncome: ctx.netIncome || 0,
        pl: (reports.pl || []).map(row),
        assets: (reports.assets || []).map(row),
        liabilities: (reports.liabilities || []).map(row),
        equity: (reports.equity || []).map(row),
        vendors: (reports.vendors || []).map(row),
        customers: (reports.customers || []).map(row),
        review1099: review1099(ctx),
        checks: list,
        ok: list.every(c => c.ok) && !ctx.hasErrors,
        log: logLines.map(l => String(l)),
    };
}

function writeExport(ctx, filePath, logLines) {
    fs.writeFileSync(filePath, JSON.stringify(buildExport(ctx, logLines), null, 1));
}

module.exports = { buildExport, writeExport, checks, review1099 };

// Fills the report context from a book, with exactly the structures that
// readSetup (lib/report/setup.js) and readSheetConfigs (lib/report/sheet-config.js)
// fill from the Setup sheet, so every later phase is unchanged.
const { NORM_VEND } = require('../report/cells');
const { warnNameOverlaps } = require('../report/name-overlaps');

// Book category type -> the { report, accountType } the rest of the code reads.
// The report is one of 'P&L', 'Balance Sheet', 'Transfer'; the accountType is the
// word Setup's Type column holds (build-reports and linkage test it for asset/bank/cash).
const CATEGORY_TYPE_MAP = {
    income: { report: 'P&L', accountType: 'Income' },
    expense: { report: 'P&L', accountType: 'Expense' },
    asset: { report: 'Balance Sheet', accountType: 'Asset' },
    liability: { report: 'Balance Sheet', accountType: 'Liability' },
    equity: { report: 'Balance Sheet', accountType: 'Equity' },
    transfer: { report: 'Transfer', accountType: 'Transfer' },
};

// Sheet types that are credit-card or expense statements: they show charges as
// positive, so polarity is flipped (the same rule readSheetConfigs applies to the
// Setup sheet types 'cc', 'expense' and anything containing 'credit' or 'liability').
const FLIPPED_SHEET_TYPES = ['cc', 'expense'];

function applyCategories(ctx, book) {
    book.categories.forEach(c => {
        const lower = c.name.toLowerCase();
        ctx.validCategories.add(lower);
        c.subcategories.forEach(sub => ctx.validSubCategories.add(sub.toLowerCase()));
        const mapped = CATEGORY_TYPE_MAP[c.type];
        ctx.uniqueCategories.set(lower, {
            report: mapped.report,
            accountType: mapped.accountType,
            // Setup keeps the last row's sub-category here; only the sheet-linkage guess reads it.
            subCategory: c.subcategories.length ? c.subcategories[c.subcategories.length - 1] : '',
            displayName: c.name,
            transferAccount: c.transferAccount,
        });
    });
}

function applyVendors(ctx, book) {
    book.vendors.forEach(v => {
        const key = NORM_VEND(v.name);
        ctx.validVendors.set(key, v.name);
        if (v.form1099 !== 'none') ctx.vendor1099Map.set(key, { type: v.form1099, req: 'YES' });
        const details = {
            business: v.business,
            name: v.recipientName || v.name, // Setup falls back to the vendor name too
            ssn: '', // tax IDs come from the private file (applyPrivate)
            address: v.address,
            email: v.email,
            phone: v.phone,
        };
        // Setup has no columns for these; they are only set when the book gives them.
        ['city', 'state', 'zip', 'country'].filter(k => v[k]).forEach(k => { details[k] = v[k]; });
        ctx.vendorDetailsMap.set(key, details);
    });
    book.customers.forEach(c => ctx.validCustomers.set(NORM_VEND(c), c));
}

// Payer info uses the keys readSetup builds from the Setup key/value list (lower case, letters and digits only).
function applyCompany(ctx, book) {
    const c = book.company;
    const keys = { name: 'companyname', address: 'address', city: 'city', state: 'state', zip: 'zip', country: 'country', email: 'email', phone: 'phone' };
    Object.entries(keys).filter(([k]) => c[k]).forEach(([k, payerKey]) => { ctx.payerInfo[payerKey] = c[k]; });
}

function applySheets(ctx, book) {
    book.sheets.forEach(s => {
        const flip = s.flip === null ? FLIPPED_SHEET_TYPES.includes(s.type) : s.flip;
        ctx.sheetConfigs.push({
            name: s.name,
            shortName: s.shortName || s.name,
            type: s.type, // 'bank', 'cc', 'income', 'expense', 'ledger' (Setup's lower-cased Sheet Type)
            flip,
            offset: s.headerRow, // 0 = find the header row, as for a Setup row without one
            startBalance: s.startBalance,
            endBalance: s.endBalance,
            linkedAccount: s.type === 'ledger' ? null : s.link,
        });
    });
}

// Replaces readSetup + readSheetConfigs.
function applyBook(ctx, book) {
    applyCategories(ctx, book);
    applyVendors(ctx, book);
    applyCompany(ctx, book);
    applySheets(ctx, book);
    if (ctx.showDebug) console.log(`[Book] ${book.categories.length} categories, ${book.sheets.length} sheets, ${book.vendors.length} vendors, ${book.customers.length} customers from ${book.file}`);
    warnNameOverlaps(ctx);
}

// Puts the tax IDs from the private file into the payer info and vendor details.
// Called after vendor.xlsx is loaded so the private file wins. The IDs live only in
// ctx (for the 1099 Data sheet); print-1099.js never prints them in book mode.
function applyPrivate(ctx, priv) {
    if (!priv) return;
    if (priv.companyTin) ctx.payerInfo.tin = priv.companyTin;
    for (const [key, tin] of priv.vendorTins) {
        const existing = ctx.vendorDetailsMap.get(key);
        if (existing) existing.ssn = tin;
    }
}

module.exports = { applyBook, applyPrivate, CATEGORY_TYPE_MAP };

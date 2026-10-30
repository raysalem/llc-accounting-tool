// The generated "Lists" worksheet: the book's categories, sub-categories, vendors
// and customers, laid out for the dropdowns on the transaction sheets.
// load_transactions.js --book writes it; nothing reads it back (the book is the source).
const path = require('path');
const { CATEGORY_TYPE_MAP } = require('./apply');

const LISTS_SHEET = 'Lists';
const FIRST_ROW = 2; // row 1 holds the headers
const LAST_ROW = 2000; // dropdowns point at rows 2..2000, so new entries appear after a refresh with no change to the transaction sheets
const NOTE_COLUMN = 'G';
const COLUMN_WIDTH = 28;
const NOTE_WIDTH = 70;
const COLUMNS = ['A', 'B', 'C', 'D', 'E']; // Categories, SubCategories, Vendors, Customers, Report Type
const HEADERS = ['Categories', 'SubCategories', 'Vendors', 'Customers', 'Report Type (for Categories)'];
// Sub-categories per category, for the Sub-Category dropdown that follows the row's Category:
// row 1 holds one category per column from column J, its sub-categories below.
const BY_CATEGORY_FIRST_COL = 10; // column J
const MAX_SUBCATEGORIES = 200; // rows read per category column
const DROPDOWN_EXTRA_ROWS = 1000; // empty rows below the data that also get dropdowns
const HEADER_SCAN_ROWS = 20; // rows searched for the header when the book gives no header row

// The five lists, each an array of strings in book order. SubCategories are unique across categories.
function listsFromBook(book) {
    const seen = new Set();
    const subCategories = [];
    book.categories.forEach(c => c.subcategories.forEach(s => {
        if (!seen.has(s.toLowerCase())) { seen.add(s.toLowerCase()); subCategories.push(s); }
    }));
    return {
        categories: book.categories.map(c => c.name),
        subCategories,
        vendors: book.vendors.map(v => v.name),
        customers: book.customers.slice(),
        reportTypes: book.categories.map(c => CATEGORY_TYPE_MAP[c.type].report),
    };
}

// Cell references for the transaction sheets' dropdowns and Report Type formula.
function listsRefs() {
    const range = col => `${LISTS_SHEET}!$${col}$${FIRST_ROW}:$${col}$${LAST_ROW}`;
    return {
        category: range('A'),
        subCategory: range('B'),
        vendor: range('C'),
        customer: range('D'),
        reportType: cell => `IFERROR(INDEX(${range('E')},MATCH(${cell},${range('A')},0)),"")`,
    };
}

// Writes (or replaces) the Lists sheet, marks it as generated and protects it
// (no password: Review > Unprotect Sheet opens it). Throws if a list is too long for the dropdown range.
async function writeListsSheet(workbook, book, now = new Date()) {
    const lists = listsFromBook(book);
    const capacity = LAST_ROW - FIRST_ROW + 1;
    Object.entries(lists).forEach(([name, values]) => {
        if (values.length > capacity) throw new Error(`The ${name} list has ${values.length} entries; the Lists sheet holds ${capacity}. Fix: raise LAST_ROW in lib/book/lists.js.`);
    });

    const existing = workbook.getWorksheet(LISTS_SHEET);
    if (existing) workbook.removeWorksheet(existing.id);
    const ws = workbook.addWorksheet(LISTS_SHEET);

    HEADERS.forEach((h, i) => {
        const cell = ws.getCell(`${COLUMNS[i]}1`);
        cell.value = h;
        cell.font = { bold: true };
        ws.getColumn(i + 1).width = COLUMN_WIDTH;
    });
    [lists.categories, lists.subCategories, lists.vendors, lists.customers, lists.reportTypes].forEach((values, i) => {
        values.forEach((v, r) => { ws.getCell(`${COLUMNS[i]}${FIRST_ROW + r}`).value = v; });
    });

    book.categories.forEach((c, i) => {
        const col = BY_CATEGORY_FIRST_COL + i;
        ws.getCell(1, col).value = c.name;
        ws.getCell(1, col).font = { bold: true };
        c.subcategories.forEach((s, r) => { ws.getCell(FIRST_ROW + r, col).value = s; });
    });

    const note = ws.getCell(`${NOTE_COLUMN}1`);
    note.value = `GENERATED from ${path.basename(book.file)} on ${now.toISOString().slice(0, 10)}. Do not edit: changes are overwritten. `
        + `Edit the book file, then run: node load_transactions.js --book ${book.id} --refresh-lists`;
    note.font = { bold: true, color: { argb: 'FFC00000' } };
    ws.getColumn(NOTE_COLUMN).width = NOTE_WIDTH;

    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true });
    return lists;
}

// Column letter(s) for a 1-based column number (1 -> A, 27 -> AA).
function colLetter(n) {
    let s = '';
    for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
    return s;
}

// A Lists column range exactly as long as its list (at least one row), so a dropdown shows no trailing blanks.
// The ranges are rewritten on every import and --refresh-lists, so they follow the book.
function exactRange(col, count) {
    return `${LISTS_SHEET}!$${col}$${FIRST_ROW}:$${col}$${FIRST_ROW + Math.max(count, 1) - 1}`;
}

// Sub-Category dropdown source for one row: the sub-categories of the Category in categoryCell,
// or every sub-category (allSubCategories, a range) while Category is empty.
function subCategoryFormula(categoryCell, categoryCount, allSubCategories = listsRefs().subCategory) {
    const first = colLetter(BY_CATEGORY_FIRST_COL);
    const last = colLetter(BY_CATEGORY_FIRST_COL + Math.max(categoryCount, 1) - 1);
    const start = `${LISTS_SHEET}!$${first}$${FIRST_ROW}`;
    const offset = `MATCH(${categoryCell},${LISTS_SHEET}!$${first}$1:$${last}$1,0)-1`;
    const count = `MAX(1,COUNTA(OFFSET(${start},0,${offset},${MAX_SUBCATEGORIES},1)))`;
    return `IF(${categoryCell}="",${allSubCategories},OFFSET(${start},0,${offset},${count},1))`;
}

// Header row and column numbers of a sheet's Category, Sub-Category, Vendor and Customer columns.
// Sub-Category is matched first: "sub-category" also contains "category".
function findDropdownColumns(ws, headerRow) {
    const rows = headerRow ? [headerRow] : Array.from({ length: HEADER_SCAN_ROWS }, (_, i) => i + 1);
    for (const r of rows) {
        const cols = {};
        ws.getRow(r).eachCell((cell, c) => {
            const v = String(cell.value || '').trim().toLowerCase();
            if (['sub-category', 'subcategory', 'sub category'].includes(v)) cols.subCategory = c;
            else if (v === 'category') cols.category = c;
            else if (v === 'vendor') cols.vendor = c;
            else if (v === 'customer') cols.customer = c;
        });
        if (cols.category) return { headerRow: r, cols };
    }
    return null;
}

// Deletes the dropdown rules of the given columns below headerRow. ExcelJS splits a range rule into one
// rule per cell when it reads a file, each keeping the top cell's formula text; left in place, every
// row's Sub-Category list would follow the top row's Category.
function clearDropdowns(ws, headerRow, columnNumbers) {
    const letters = new Set(columnNumbers.map(colLetter));
    Object.keys(ws.dataValidations.model).forEach(address => {
        const m = /^([A-Z]+)(\d+)(?::[A-Z]+\d+)?$/.exec(address);
        if (m && letters.has(m[1]) && Number(m[2]) > headerRow) delete ws.dataValidations.model[address];
    });
}

// Puts the dropdowns on every book sheet in the workbook (transaction sheets and the ledger), from the
// row under the header to DROPDOWN_EXTRA_ROWS below the last row, as one range rule per column. A range
// rule creates no rows, so the next import still appends right under the data. The Sub-Category rule's
// row-relative reference ($D4) is adjusted by Excel for each row of the range.
// Returns the names of the sheets done.
function applyDropdowns(workbook, book) {
    const lists = listsFromBook(book);
    const refs = {
        category: exactRange('A', lists.categories.length),
        subCategory: exactRange('B', lists.subCategories.length),
        vendor: exactRange('C', lists.vendors.length),
        customer: exactRange('D', lists.customers.length),
    };
    const list = formula => ({ type: 'list', allowBlank: true, formulae: [formula] });
    const done = [];
    for (const s of book.sheets) {
        const ws = workbook.getWorksheet(s.name);
        if (!ws) continue;
        const found = findDropdownColumns(ws, s.headerRow);
        if (!found) continue;
        const { headerRow, cols } = found;
        const first = headerRow + 1;
        const last = Math.max(ws.rowCount, headerRow) + DROPDOWN_EXTRA_ROWS;
        clearDropdowns(ws, headerRow, Object.values(cols));
        const range = c => `${colLetter(c)}${first}:${colLetter(c)}${last}`;
        ws.dataValidations.add(range(cols.category), list(refs.category));
        if (cols.subCategory) {
            ws.dataValidations.add(range(cols.subCategory), list(subCategoryFormula(`$${colLetter(cols.category)}${first}`, book.categories.length, refs.subCategory)));
        }
        if (cols.vendor) ws.dataValidations.add(range(cols.vendor), list(refs.vendor));
        if (cols.customer) ws.dataValidations.add(range(cols.customer), list(refs.customer));
        done.push(s.name);
    }
    return done;
}

module.exports = { writeListsSheet, listsFromBook, listsRefs, applyDropdowns, subCategoryFormula, LISTS_SHEET };

// Builds a starting workbook for a book: one transaction sheet per [[sheets]]
// entry (header on row 3, TOTAL/SUBTOTAL above), a Ledger, VERSION and the
// generated Lists sheet. There is no Setup sheet: the book holds the configuration.
const { prepareTargetSheet } = require('../loader/target-sheet');
const { writeListsSheet, applyDropdowns } = require('./lists');

// Entry ID is a label for the user to group the lines of one entry; the report does not read it.
const LEDGER_HEADERS = ['Entry ID', 'Date', 'Description', 'Category', 'Sub-Category', 'Debit', 'Credit', 'Vendor', 'Customer'];
const LEDGER_COLUMN_WIDTHS = [10, 12, 30, 20, 20, 15, 15, 20, 20];

async function buildBookWorkbook(ExcelJS, book, versionRows) {
    const workbook = new ExcelJS.Workbook();
    for (const s of book.sheets) {
        if (s.type === 'ledger') continue;
        prepareTargetSheet(workbook, s.type === 'cc' ? 'cc' : 'bank', false, { sheetName: s.name });
    }

    const declaredLedger = book.sheets.find(s => s.type === 'ledger');
    const ledger = workbook.addWorksheet(declaredLedger ? declaredLedger.name : 'Ledger');
    // Header on the row the book declares (row 1 when it gives none), so the report finds it there.
    const ledgerHeaderRow = (declaredLedger && declaredLedger.headerRow) || 1;
    LEDGER_HEADERS.forEach((header, i) => {
        ledger.getColumn(i + 1).width = LEDGER_COLUMN_WIDTHS[i];
        ledger.getCell(ledgerHeaderRow, i + 1).value = header;
    });
    ledger.getRow(ledgerHeaderRow).font = { bold: true };

    // Blank tabs from the book (e.g. a scratch "Worksheet"), after the ledger; the report ignores them.
    book.blankSheets.forEach(name => workbook.addWorksheet(name));

    const version = workbook.addWorksheet('VERSION');
    version.columns = [{ header: 'Property', width: 20 }, { header: 'Value', width: 50 }];
    versionRows.forEach(r => version.addRow(r));

    await writeListsSheet(workbook, book);
    applyDropdowns(workbook, book);
    return workbook;
}

module.exports = { buildBookWorkbook };

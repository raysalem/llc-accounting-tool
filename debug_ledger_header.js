const ExcelJS = require('exceljs');
async function inspectLedgerHeaders() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\valencia\\2025-Valencia.xlsx";
    console.log(`Loading ${filename}...`);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const ledger = workbook.getWorksheet('Ledger');
    if (!ledger) { console.log('No Ledger sheet found.'); return; }

    const row1 = ledger.getRow(1).values;
    console.log('Ledger Header Row:', JSON.stringify(row1));
}
inspectLedgerHeaders().catch(console.error);

const ExcelJS = require('exceljs');
async function debugClear() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\rmp proprietor\\2025-RMP Prop.xlsx";
    console.log(`Loading ${filename}...`);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const sheet = workbook.getWorksheet('Credit Card Transactions');
    if (!sheet) {
        console.log('Sheet not found!');
        return;
    }

    console.log(`Sheet Row Count: ${sheet.rowCount}`);
    console.log(`Sheet Last Row Number: ${sheet.lastRow ? sheet.lastRow.number : 'N/A'}`);

    const r1 = sheet.getRow(1).values;
    const r2 = sheet.getRow(2).values;
    const r3 = sheet.getRow(3).values;

    console.log('Row 1:', JSON.stringify(r1));
    console.log('Row 2:', JSON.stringify(r2));
    console.log('Row 3:', JSON.stringify(r3));
}
debugClear().catch(console.error);

const ExcelJS = require('exceljs');

async function inspectFile() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('ax 3rd.xlsx');
    const sheet = workbook.getWorksheet('Transaction Details');
    if (!sheet) { console.log('Sheet not found'); return; }

    console.log('--- Transaction Details Structure (Rows 5-8) ---');
    for (let i = 5; i <= 8; i++) {
        const row = sheet.getRow(i);
        const vals = (row.values || []).slice(1).map(v => v ? v.toString().trim() : '');
        console.log(`Row ${i}: [${vals.join('] [')}]`);
    }
}
inspectFile();

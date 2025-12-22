const ExcelJS = require('exceljs');
const fs = require('fs');

async function listSheets() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\RMP DEVELOPMENT\\2025-RMP Development.xlsx";
    if (!fs.existsSync(filename)) {
        console.log("File not found");
        return;
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    console.log("Sheets in workbook:");
    workbook.worksheets.forEach(sheet => {
        console.log(`- "${sheet.name}"`);
    });
}

listSheets().catch(console.error);

const ExcelJS = require('exceljs');
const fs = require('fs');

async function dumpSetupConfig() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\3rd\\2025-3rd-Accounting_.xlsx";
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const setup = workbook.getWorksheet('Setup');

    console.log("\n--- Setup: Sheet Configuration (Columns I-L) ---");
    setup.eachRow((row, r) => {
        const name = row.getCell(9).value;
        const type = row.getCell(10).value;
        const flip = row.getCell(11).value;
        const offset = row.getCell(12).value;
        if (name || type) {
            console.log(`Row ${r}: Name="${name}" | Type="${type}" | Flip="${flip}" | Offset="${offset}"`);
        }
    });

    console.log("\nAll Worksheets in file:");
    workbook.worksheets.forEach(s => console.log(` - "${s.name}"`));
}

dumpSetupConfig().catch(console.error);

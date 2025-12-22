const ExcelJS = require('exceljs');
const fs = require('fs');

async function debugFile() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\3rd\\2025-3rd-Accounting_.xlsx";
    console.log(`Checking: ${filename}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    console.log("\nActual sheets in workbook:");
    workbook.worksheets.forEach(s => console.log(` - "${s.name}"`));

    const setupSheet = workbook.getWorksheet('Setup');
    if (!setupSheet) {
        console.log("No Setup sheet found!");
        return;
    }

    console.log("\nSetup Sheet Config (Rows 2-10):");
    for (let i = 2; i <= 10; i++) {
        const row = setupSheet.getRow(i);
        const name = row.getCell(9).value;
        const type = row.getCell(10).value;
        const offset = row.getCell(12).value;
        if (name || type) {
            console.log(`Row ${i}: Name="${name}", Type="${type}", Offset="${offset}"`);
        }
    }

    // Look at Jan 13 category issue
    const targetDate = new Date('2025-01-13');
    // We'll check all sheets for this date
    workbook.worksheets.forEach(sheet => {
        if (sheet.name === 'Setup' || sheet.name === 'Summary' || sheet.name === 'Notes') return;
        console.log(`\nSearching "${sheet.name}" for Date: 2025-01-13`);
        sheet.eachRow((row, r) => {
            let dateVal = row.getCell(1).value;
            if (dateVal && typeof dateVal === 'object' && dateVal.result !== undefined) dateVal = dateVal.result;

            if (dateVal instanceof Date) {
                if (dateVal.getFullYear() === 2025 && dateVal.getMonth() === 0 && dateVal.getDate() === 13) {
                    console.log(`Row ${r}: Date=${dateVal.toISOString()}, Desc="${row.getCell(2).value}", Cat="${row.getCell(4).value}", Vendor="${row.getCell(7).value}" (BANK INDEXES)`);
                    console.log(`Row ${r}: Date=${dateVal.toISOString()}, Desc="${row.getCell(3).value}", Cat="${row.getCell(5).value}", Vendor="${row.getCell(8).value}" (CC INDEXES)`);
                }
            }
        });
    });
}

debugFile().catch(console.error);

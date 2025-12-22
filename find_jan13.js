const ExcelJS = require('exceljs');
const fs = require('fs');

async function findJan13() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\3rd\\2025-3rd-Accounting_.xlsx";
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    workbook.worksheets.forEach(sheet => {
        console.log(`\nScanning sheet: "${sheet.name}"`);
        sheet.eachRow((row, r) => {
            let dateVal = row.getCell(1).value;
            if (dateVal && typeof dateVal === 'object' && dateVal.result !== undefined) dateVal = dateVal.result;

            if (dateVal instanceof Date) {
                const year = dateVal.getFullYear();
                const month = dateVal.getMonth() + 1; // 1-indexed
                const day = dateVal.getDate();
                if (year === 2025 && month === 1 && day === 13) {
                    const rowVals = row.values.slice(1).map(v => (v ? v.toString().trim() : ''));
                    console.log(`Row ${r}: ${rowVals.join(' | ')}`);
                }
            }
        });
    });
}

findJan13().catch(console.error);

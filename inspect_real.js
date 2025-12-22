const ExcelJS = require('exceljs');
const fs = require('fs');

async function inspectRealFile() {
    const filename = '2025-3rd-accounting.xlsx';
    console.log(`Inspecting ${filename}...`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const setupSheet = workbook.getWorksheet('Setup');
    if (setupSheet) {
        console.log('\n--- Setup Sheet Config ---');
        setupSheet.eachRow((row, r) => {
            if (r === 1 || (row.getCell(9).value)) {
                console.log(`Row ${r}: Col I: ${row.getCell(9).value} | Col J: ${row.getCell(10).value} | Col K: ${row.getCell(11).value}`);
            }
        });
    } else {
        console.log('Setup sheet missing.');
    }

    const sheets = workbook.worksheets.map(s => s.name);
    console.log('\nAvailable Sheets:', sheets);

    // Let's also check first few rows of Bank and CC transactions if they match expected layouts
    for (const sheetName of ['Bank Transactions', 'Credit Card Transactions']) {
        const sheet = workbook.getWorksheet(sheetName);
        if (sheet) {
            console.log(`\n--- First 3 rows of ${sheetName} ---`);
            sheet.eachRow((row, r) => {
                if (r <= 3) {
                    const values = [];
                    row.eachCell((cell, col) => values.push(`${col}:${cell.value}`));
                    console.log(`Row ${r}: ${values.join(' | ')}`);
                }
            });
        }
    }
}

inspectRealFile().catch(console.error);

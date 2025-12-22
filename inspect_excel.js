const ExcelJS = require('exceljs');

async function inspectFile() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('ax 3rd.xlsx');

        workbook.worksheets.forEach(sheet => {
            console.log(`\n--- Sheet: ${sheet.name} ---`);
            // Inspect first 20 rows to find headers
            for (let i = 1; i <= 20; i++) {
                const row = sheet.getRow(i);
                // row.values is 1-indexed (index 0 is null/empty usually), so slice(1)
                const values = (row.values || []).slice(1).map(v => v ? v.toString().trim() : '');

                // Only print if row has content
                if (values.some(v => v !== '')) {
                    console.log(`Row ${i}: ${values.join(' | ')}`);
                }
            }
        });
    } catch (error) {
        console.error('Error reading file:', error);
    }
}

inspectFile();

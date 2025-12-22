const ExcelJS = require('exceljs');
const fs = require('fs');

async function inspectFile() {
    const filename = process.argv[2] || 'LLC_Accounting_Template.xlsx';
    if (!fs.existsSync(filename)) {
        console.error(`Error: File '${filename}' not found.`);
        console.log('Usage: node inspect.js <filename.xlsx>');
        return;
    }

    console.log(`\n=== Inspecting: ${filename} ===`);
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(filename);
    } catch (e) {
        console.error(`Error reading file: ${e.message}`);
        return;
    }

    const setupSheet = workbook.getWorksheet('Setup');
    if (setupSheet) {
        console.log('\n--- Setup: Sheet Configurations ---');
        setupSheet.eachRow((row, r) => {
            if (r === 1) return;
            const name = row.getCell(9).value;
            const type = row.getCell(10).value;
            const flip = row.getCell(11).value;
            const offset = row.getCell(12).value;
            if (name) {
                console.log(`Sheet: ${name.toString().padEnd(25)} | Type: ${type.toString().padEnd(10)} | Flip: ${flip.toString().padEnd(5)} | Offset: ${offset}`);
            }
        });
    }

    workbook.worksheets.forEach(sheet => {
        if (sheet.name === 'Setup') return;
        console.log(`\n--- Sheet: ${sheet.name} (${sheet.rowCount} rows) ---`);
        for (let i = 1; i <= Math.min(sheet.rowCount, 5); i++) {
            const values = (sheet.getRow(i).values || []).slice(1).map(v => {
                if (v && typeof v === 'object' && v.result !== undefined) return v.result;
                return v ? v.toString().trim() : '';
            });
            if (values.length > 0) {
                console.log(`Row ${i}: ${values.join(' | ')}`);
            }
        }
    });
}

inspectFile().catch(console.error);

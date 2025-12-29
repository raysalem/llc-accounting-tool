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

        const headerRow = setupSheet.getRow(1);
        const colMap = {};
        headerRow.eachCell((cell, colNumber) => {
            const val = cell.value ? cell.value.toString().toLowerCase().trim() : '';
            if (val.includes('sheet name') && !val.includes('config')) colMap.name = colNumber; // 'Sheet Name' vs 'Sheet Name (Config)'
            if (val.includes('sheet name (config)')) colMap.name = colNumber;
            if (val.includes('account type') || val.includes('sheet type')) colMap.type = colNumber;
            if (val.includes('flip')) colMap.flip = colNumber;
            if (val.includes('header') || val.includes('offset')) colMap.offset = colNumber;
        });

        setupSheet.eachRow((row, r) => {
            if (r === 1) return;
            const name = colMap.name ? row.getCell(colMap.name).value : null;
            const type = colMap.type ? row.getCell(colMap.type).value : null;
            const flip = colMap.flip ? row.getCell(colMap.flip).value : null;
            const offset = colMap.offset ? row.getCell(colMap.offset).value : null;
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

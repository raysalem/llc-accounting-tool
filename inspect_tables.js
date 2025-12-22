const ExcelJS = require('exceljs');
const fs = require('fs');

async function inspectTables() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\RMP DEVELOPMENT\\2025-RMP Development.xlsx";
    if (!fs.existsSync(filename)) {
        console.log("File not found");
        return;
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    workbook.worksheets.forEach(sheet => {
        console.log(`Sheet: ${sheet.name}`);
        // exceljs stores tables in an internal list
        if (sheet.tables) {
            console.log(`  Found ${Object.keys(sheet.tables).length} tables:`);
            for (let name in sheet.tables) {
                const table = sheet.tables[name];
                console.log(`    Table Name: ${name}`);
                console.log(`    Table Range: ${table.tableRange}`);
                console.log(`    Columns: ${table.columns.map(c => c.name).join(', ')}`);
            }
        } else {
            console.log("  No tables found in metadata (or not supported by current exceljs version for read).");
        }
    });
}

inspectTables().catch(console.error);

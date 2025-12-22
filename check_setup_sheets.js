const ExcelJS = require('exceljs');
const fs = require('fs');

async function checkSetup() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\RMP DEVELOPMENT\\2025-RMP Development.xlsx";
    if (!fs.existsSync(filename)) {
        console.log("File not found");
        return;
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const setup = workbook.getWorksheet('Setup');
    if (setup) {
        console.log("Setup Sheet Rows (Col I-L):");
        setup.eachRow((row, r) => {
            if (r === 1) return;
            const name = row.getCell(9).value; // I
            const type = row.getCell(10).value; // J
            if (name) {
                console.log(`- Row ${r}: Name="${name}", Type="${type}"`);
            }
        });
    } else {
        console.log("Setup sheet NOT found.");
    }
}

checkSetup().catch(console.error);

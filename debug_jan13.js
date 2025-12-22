const ExcelJS = require('exceljs');
const fs = require('fs');

async function debugUpdate() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\3rd\\2025-3rd-Accounting_.xlsx";
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const setup = workbook.getWorksheet('Setup');
    const validCats = new Set();
    setup.eachRow((row, r) => {
        if (r === 1) return;
        let v = row.getCell(1).value;
        if (v && typeof v === 'object' && v.result !== undefined) v = v.result;
        if (v) validCats.add(v.toString().trim());
    });

    console.log("Valid Categories from Setup:", Array.from(validCats));

    const cc = workbook.getWorksheet('Credit Card Transactions');
    const headerRow = cc.getRow(3); // Offset 3 means headers at 3? Or headers at 2?
    // User said offset is 3. Usually that means skip 1,2,3. So headers are likely at 3.
    console.log("\nCC Headers (Row 3):");
    const headers = [];
    headerRow.eachCell((c, i) => headers.push(`${i}: ${c.value}`));
    console.log(headers.join(' | '));

    console.log("\nJan 13 Row (Row 12):");
    const row12 = cc.getRow(12);
    row12.eachCell({ includeEmpty: true }, (c, i) => {
        let v = c.value;
        let type = typeof v;
        if (v && typeof v === 'object') {
            if (v.result !== undefined) v = `Result: ${v.result}`;
            else if (v.formula) v = `Formula: ${v.formula}`;
        }
        console.log(`Col ${i}: ${v} (${type})`);
    });
}

debugUpdate().catch(console.error);

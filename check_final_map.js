const ExcelJS = require('exceljs');
const fs = require('fs');

async function checkFinalMap() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\3rd\\2025-3rd-Accounting_.xlsx";
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const setup = workbook.getWorksheet('Setup');
    const configs = [];
    setup.eachRow((row, r) => {
        if (r === 1) return;
        if (row.getCell(9).value) {
            configs.push({
                name: row.getCell(9).value.toString().trim(),
                offset: parseInt(row.getCell(12).value) || 0
            });
        }
    });

    for (const conf of configs) {
        const sheet = workbook.getWorksheet(conf.name);
        if (!sheet) continue;
        console.log(`\nSheet: ${sheet.name}`);
        const headerRow = sheet.getRow(conf.offset || 1);
        const map = {};
        headerRow.eachCell((cell, colNumber) => {
            const val = cell.value ? cell.value.toString().toLowerCase() : '';
            if (val === 'date') map.date = colNumber;
            else if (val === 'description' || val === 'desc') map.desc = colNumber;
            else if (val === 'amount') map.amount = colNumber;
            else if (val === 'category' || val === 'cat') map.category = colNumber;
            else if (val === 'vendor' || val === 'vend') map.vendor = colNumber;
            else if (val === 'customer' || val === 'cust') map.customer = colNumber;
        });
        console.log(`  Discovered Map: ${JSON.stringify(map)}`);

        // Show row 12 in CC
        if (sheet.name === 'Credit Card Transactions') {
            const row12 = sheet.getRow(12);
            console.log(`  Row 12 Data (using Map):`);
            for (let k in map) {
                console.log(`    ${k}: ${row12.getCell(map[k]).value}`);
            }
        }
    }
}

checkFinalMap().catch(console.error);

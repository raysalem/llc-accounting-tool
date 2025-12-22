const ExcelJS = require('exceljs');
const fs = require('fs');

async function dumpHeaders() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\3rd\\2025-3rd-Accounting_.xlsx";
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const sheets = ['Bank Transactions', 'Credit Card Transactions', 'Ledger', 'Setup'];

    sheets.forEach(name => {
        const sheet = workbook.getWorksheet(name);
        if (!sheet) {
            console.log(`\nSheet "${name}" not found.`);
            return;
        }
        console.log(`\n--- Sheet: ${name} ---`);
        for (let i = 1; i <= 10; i++) {
            const row = sheet.getRow(i);
            const vals = [];
            for (let j = 1; j <= 12; j++) {
                let v = row.getCell(j).value;
                if (v && typeof v === 'object' && v.result !== undefined) v = v.result;
                if (v && typeof v === 'object' && v.richText) v = v.richText.map(t => t.text).join('');
                vals.push(v === null || v === undefined ? '' : v.toString().trim());
            }
            console.log(`Row ${i.toString().padStart(2)}: ${vals.join(' | ')}`);
        }
    });
}

dumpHeaders().catch(console.error);

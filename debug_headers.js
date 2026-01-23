const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const filename = '\\\\192.168.1.90\\Documents Private\\taxes\\2025\\rmp ventures\\2025-RMP Ventrues.xlsx';

async function checkHeaders() {
    const workbook = new ExcelJS.Workbook();
    console.log(`Loading workbook: ${filename}`);
    await workbook.xlsx.readFile(filename);

    const sheetName = 'Credit Card Transactions';
    const sheet = workbook.getWorksheet(sheetName);

    if (!sheet) {
        // Try fuzzy find
        const s = workbook.worksheets.find(ws => ws.name.toLowerCase().includes('card') || ws.name.toLowerCase().includes('cc'));
        if (s) {
            console.log(`Found similar sheet: "${s.name}"`);
            printHeaders(s);
        } else {
            console.log(`Sheet "${sheetName}" not found! Available sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);
        }
    } else {
        console.log(`Found sheet: "${sheetName}"`);
        printHeaders(sheet);
    }
}

function printHeaders(sheet) {
    // Check first 5 rows for something that looks like a header
    for (let r = 1; r <= 5; r++) {
        const row = sheet.getRow(r);
        const vals = row.values;
        if (Array.isArray(vals)) {
            // exceljs row.values is 1-based, index 0 is undefined usually?
            // Actually it returns [ <empty>, val1, val2... ]
            const cleanVals = vals.map(v => v ? v.toString().trim() : '').filter(v => v !== '');
            if (cleanVals.length > 0) {
                console.log(`Row ${r}: ${cleanVals.join(' | ')}`);
            }
        }
    }
}

checkHeaders().catch(console.error);

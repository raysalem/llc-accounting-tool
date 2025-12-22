const ExcelJS = require('exceljs');

async function probe() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('2025-3rd-accounting.xlsx');

    const setupSheet = workbook.getWorksheet('Setup');
    const categories = [];
    setupSheet.eachRow((row, r) => {
        if (r > 1) categories.push(row.getCell(1).value ? row.getCell(1).value.toString() : '');
    });
    console.log('Categories in Setup (first 10):', categories.slice(0, 10).map(c => `[${c}]`));

    const bankSheet = workbook.getWorksheet('Bank Transactions');
    if (bankSheet) {
        console.log('\nScanning Bank Transactions...');
        let found = 0;
        bankSheet.eachRow((row, r) => {
            if (r > 1) {
                const cat = row.getCell(4).value;
                const amt = row.getCell(3).value;
                if (cat) {
                    const catStr = cat.toString();
                    if (found < 5) console.log(`Row ${r}: Cat=[${catStr}], Amt=${amt}`);
                    found++;
                }
            }
        });
        console.log(`Found ${found} rows with categories in Bank Sheet.`);
    }
}
probe().catch(console.error);

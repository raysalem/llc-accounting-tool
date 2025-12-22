const ExcelJS = require('exceljs');

async function checkRent() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('2025-3rd-accounting.xlsx');
    const bankSheet = workbook.getWorksheet('Bank Transactions');

    console.log('--- Rent Rows in Bank Transactions ---');
    bankSheet.eachRow((row, r) => {
        if (r > 1 && row.getCell(4).value === 'Rent') {
            console.log(`Row ${r}: Desc=[${row.getCell(2).value}], Amt=[${row.getCell(3).value}]`);
        }
    });
}
checkRent().catch(console.error);

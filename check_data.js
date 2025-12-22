const ExcelJS = require('exceljs');

async function checkBankSheet() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('LLC_Accounting_Example_With_Data.xlsx');
        const sheet = workbook.getWorksheet('Bank Transactions');
        console.log(`Sheet: ${sheet.name}, Row Count: ${sheet.rowCount}`);
        if (sheet.rowCount > 0) {
            console.log('Headers:', sheet.getRow(1).values);
            const firstRow = sheet.getRow(2).values; // Row 2 is data
            console.log('Row 2 Data:', firstRow);
        } else {
            console.log('Sheet is empty');
        }
    } catch (e) {
        console.error(e);
    }
}
checkBankSheet();

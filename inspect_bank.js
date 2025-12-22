const ExcelJS = require('exceljs');

async function inspectBank() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('2025-3rd-accounting.xlsx');
    const bankSheet = workbook.getWorksheet('Bank Transactions');
    if (bankSheet) {
        console.log('\n--- Bank Transactions Header ---');
        const row1 = bankSheet.getRow(1);
        row1.eachCell((cell, col) => console.log(`${col}: ${cell.value}`));

        console.log('\n--- Bank Transactions Row 2 ---');
        const row2 = bankSheet.getRow(2);
        row2.eachCell((cell, col) => console.log(`${col}: ${cell.value}`));
    }
}
inspectBank().catch(console.error);

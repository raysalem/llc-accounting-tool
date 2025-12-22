const ExcelJS = require('exceljs');
async function checkSetupRow() {
    const filename = "\\\\192.168.1.90\\Documents Private\\taxes\\2025\\RMP DEVELOPMENT\\2025-RMP Development.xlsx";
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const sheet = workbook.getWorksheet('Setup');
    sheet.eachRow((row, r) => {
        if (r === 1) return;
        const name = row.getCell(9).value;
        const type = row.getCell(10).value;
        const flip = row.getCell(11).value;
        console.log(`Row ${r}: Name='${name}', Type='${type}', Flip='${flip}'`);
    });
}
checkSetupRow().catch(console.error);

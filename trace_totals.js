const ExcelJS = require('exceljs');

async function trace() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('2025-3rd-accounting.xlsx');

    const catStats = {};
    const bankSheet = workbook.getWorksheet('Bank Transactions');
    bankSheet.eachRow((row, r) => {
        if (r > 1) {
            const cat = row.getCell(4).value;
            let amt = row.getCell(3).value || 0;
            if (amt && typeof amt === 'object') amt = amt.result || 0;
            if (cat) {
                const s = cat.toString();
                catStats[s] = (catStats[s] || 0) + amt;
            }
        }
    });
    console.log('--- Bank Totals ---');
    console.log(catStats);

    const ccStats = {};
    const ccSheet = workbook.getWorksheet('Credit Card Transactions');
    ccSheet.eachRow((row, r) => {
        if (r > 1) {
            const cat = row.getCell(5).value;
            let amt = row.getCell(4).value || 0;
            if (amt && typeof amt === 'object') amt = amt.result || 0;
            if (cat) {
                const s = cat.toString();
                ccStats[s] = (ccStats[s] || 0) + (amt * -1); // Flipped
            }
        }
    });
    console.log('\n--- CC Totals (Flipped) ---');
    console.log(ccStats);
}
trace().catch(console.error);

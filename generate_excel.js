const ExcelJS = require('exceljs');
const { execSync } = require('child_process');
const pkg = require('./package.json');

async function createTemplate() {
    const filename = 'LLC_Accounting_Template.xlsx';
    const workbook = new ExcelJS.Workbook();

    const setupSheet = workbook.addWorksheet('Setup');
    setupSheet.columns = [
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Sub-Category', key: 'subcategory', width: 25 },
        { header: 'Type', key: 'type', width: 15 },
        { header: 'Report', key: 'report', width: 15 },
        { header: '', key: 'spacer1', width: 5 },
        { header: 'Vendors', key: 'vendors', width: 25 },
        { header: 'Customers', key: 'customers', width: 25 },
        { header: '', key: 'spacer2', width: 5 },
        { header: 'Sheet Name (Config)', key: 'sheetname', width: 30 },
        { header: 'Account Type', key: 'sheettype', width: 15 },
        { header: 'Flip Polarity? (Yes/No)', key: 'flip', width: 20 },
        { header: 'Header Row', key: 'offset', width: 15 },
    ];

    const categories = [
        ['Sales', 'General', 'Income', 'P&L'],
        ['Rent', 'Office', 'Expense', 'P&L'],
        ['Checking Account', 'Bank', 'Asset', 'Balance Sheet'],
        ['Credit Card', 'Liability', 'Liability', 'Balance Sheet'],
        ['AX CC', 'Liability', 'Liability', 'Balance Sheet'],
    ];
    setupSheet.addRows(categories);

    setupSheet.getCell('I2').value = 'Bank Transactions';
    setupSheet.getCell('J2').value = 'Bank';
    setupSheet.getCell('K2').value = 'No';
    setupSheet.getCell('L2').value = 1;

    setupSheet.getCell('I3').value = 'Credit Card Transactions';
    setupSheet.getCell('J3').value = 'CC';
    setupSheet.getCell('K3').value = 'Yes'; // Default to Yes for CC
    setupSheet.getCell('L3').value = 1;

    // Add some sample vendors/customers
    setupSheet.getCell('F2').value = 'Starbucks';
    setupSheet.getCell('F3').value = 'Amazon';
    setupSheet.getCell('F4').value = 'AWS';

    const bankSheet = workbook.addWorksheet('Bank Transactions');
    bankSheet.columns = [
        { header: 'Date', key: 'date', width: 12 }, { header: 'Description', key: 'desc', width: 35 },
        { header: 'Amount', key: 'amount', width: 15 }, { header: 'Category', key: 'category', width: 20 },
        { header: 'Sub-Category', key: 'subcategory', width: 20 }, { header: 'Extended', width: 30 },
        { header: 'Vendor', width: 20 }, { header: 'Customer', width: 20 }, { header: 'Type (Auto)', width: 15 }
    ];
    bankSheet.autoFilter = { from: 'A1', to: 'I1' };

    const ccSheet = workbook.addWorksheet('Credit Card Transactions');
    ccSheet.columns = [
        { header: 'Date', key: 'date', width: 12 }, { header: 'Member', width: 15 }, { header: 'Description', key: 'desc', width: 35 },
        { header: 'Amount', key: 'amount', width: 15 }, { header: 'Category', key: 'category', width: 20 },
        { header: 'Sub-Category', key: 'subcategory', width: 20 }, { header: 'Extended', width: 30 },
        { header: 'Vendor', width: 20 }, { header: 'Customer', width: 20 }, { header: 'Acct #', width: 15 },
        { header: 'Receipt', width: 10 }, { header: 'Type (Auto)', width: 15 }
    ];
    ccSheet.autoFilter = { from: 'A1', to: 'L1' };

    const ledgerSheet = workbook.addWorksheet('Ledger');
    ledgerSheet.columns = [
        { header: 'Date', width: 12 }, { header: 'Description', width: 30 }, { header: 'Category', width: 20 },
        { header: 'Debit', width: 15 }, { header: 'Credit', width: 15 }, { header: 'Type (Auto)', width: 15 }
    ];

    workbook.addWorksheet('Summary');

    const versionSheet = workbook.addWorksheet('VERSION');
    versionSheet.columns = [{ header: 'Property', width: 20 }, { header: 'Value', width: 50 }];

    let gitSha = 'N/A';
    try { gitSha = execSync('git rev-parse HEAD').toString().trim(); } catch (e) { }

    versionSheet.addRow(['Version ID', pkg.version]);
    versionSheet.addRow(['Git SHA', gitSha]);
    versionSheet.addRow(['Generated At', new Date().toLocaleString()]);

    await workbook.xlsx.writeFile(filename);
    console.log(`Template updated: ${filename}`);
}
createTemplate();

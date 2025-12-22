const ExcelJS = require('exceljs');

async function createExampleData() {
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
    ];

    const categories = [
        ['Sales', 'General', 'Income', 'P&L'],
        ['Rent', 'Office', 'Expense', 'P&L'],
        ['Checking Account', 'Bank', 'Asset', 'Balance Sheet'],
        ['AX CC', 'Liability', 'Liability', 'Balance Sheet'],
    ];
    setupSheet.addRows(categories);

    setupSheet.getCell('I2').value = 'Bank Transactions';
    setupSheet.getCell('J2').value = 'Bank';
    setupSheet.getCell('K2').value = 'No';
    setupSheet.getCell('I3').value = 'Credit Card Transactions';
    setupSheet.getCell('J3').value = 'CC';
    setupSheet.getCell('K3').value = 'Yes'; // Most CC statements use positive for charges

    const bankSheet = workbook.addWorksheet('Bank Transactions');
    bankSheet.columns = [{ header: 'Date', width: 12 }, { header: 'Description', width: 35 }, { header: 'Amount', width: 15 }, { header: 'Category', width: 20 }, { header: 'Sub-Category', width: 20 }, { header: 'Extended', width: 30 }, { header: 'Vendor', width: 20 }, { header: 'Customer', width: 20 }, { header: 'Type (Auto)', width: 15 }];

    // In Flow
    bankSheet.addRow([new Date(), 'Rent Payment', -1000, 'Rent', '', '', '', '', 'P&L']);
    bankSheet.addRow([new Date(), 'Client XYZ Deposit', 5000, 'Sales', '', '', '', 'Client XYZ', 'P&L']);

    const ccSheet = workbook.addWorksheet('Credit Card Transactions');
    ccSheet.columns = [
        { header: 'Date', width: 12 }, { header: 'Member', width: 15 }, { header: 'Description', width: 35 },
        { header: 'Amount', width: 15 }, { header: 'Category', width: 20 },
        { header: 'Sub-Category', width: 20 }, { header: 'Extended', width: 30 },
        { header: 'Vendor', width: 20 }, { header: 'Customer', width: 20 }, { header: 'Acct #', width: 15 },
        { header: 'Receipt', width: 10 }, { header: 'Type (Auto)', width: 15 }
    ];
    // CC Statement: Purchases are usually positive numbers in raw exports
    ccSheet.addRow([new Date(), 'John Doe', 'Amazon.com', 45.99, 'Rent', 'Office', '', 'Amazon', '', '1234', '', 'P&L']);

    workbook.addWorksheet('Ledger');
    workbook.addWorksheet('Summary');

    await workbook.xlsx.writeFile('LLC_Accounting_Example_With_Data.xlsx');
    console.log(`Example data updated.`);
}
createExampleData();

const ExcelJS = require('exceljs');
const { execSync } = require('child_process');
const fs = require('fs');
const pkg = require('./package.json');
const { parseOptionValue, stripOption } = require('./lib/cli');
const book = require('./lib/book');

// Help Support
if (process.argv.includes('--help')) {
    console.log(`
Usage: node generate_excel.js [outputFilename]
       node generate_excel.js --book <id|path> [outputFilename]

Description:
  Generates the core LLC Accounting Template Excel file.
  Includes standard tabs: Setup, Ledger, Bank Transactions, Credit Card Transactions, and Summary.

Arguments:
  [outputFilename]   (Optional) Name of the file to generate. Defaults to "LLC_Accounting_Template.xlsx".

  --book <id|path>   Make a starting workbook for a book file instead (see docs/setup.md): one sheet per
                     [[sheets]] entry, a Ledger and a generated Lists sheet, and NO Setup sheet. Written to the
                     book's "workbook" path unless a file name is given. Refuses to overwrite an existing file.

Example:
  node generate_excel.js "My_Consulting_Books_2025.xlsx"
  node generate_excel.js --book examples/example.book.toml
    `);
    process.exit(0);
}

// --book: a workbook with the book's sheets and Lists, and no Setup sheet.
async function createBookWorkbook(args) {
    const ref = parseOptionValue(args, '--book');
    if (!ref) throw new Error('--book needs a book id or the path of a book file (.toml).');
    const config = book.loadBook(ref);
    const filename = stripOption(args, '--book').find(a => !a.startsWith('--')) || config.workbookPath;
    if (fs.existsSync(filename)) throw new Error(`"${filename}" already exists; not overwriting it. Fix: give another file name, or delete it first.`);
    let gitSha = 'N/A';
    try {
        gitSha = execSync('git rev-parse HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
    } catch (e) {
        // Not a git checkout; keep 'N/A'
    }
    const workbook = await book.buildBookWorkbook(ExcelJS, config, [
        ['Version ID', pkg.version], ['Git SHA', gitSha], ['Generated At', new Date().toLocaleString()], ['Book', config.id],
    ]);
    await workbook.xlsx.writeFile(filename);
    console.log(`Workbook created for book "${config.id}": ${filename}`);
}

async function createTemplate() {
    const args = process.argv.slice(2);
    if (args.some(a => a === '--book' || a.startsWith('--book='))) {
        try {
            await createBookWorkbook(args);
        } catch (e) {
            console.error(`Error: ${e.message}`);
            process.exitCode = 1;
        }
        return;
    }
    const filename = args.find(a => !a.startsWith('--')) || 'LLC_Accounting_Template.xlsx';
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
        { header: 'Sheet Type', key: 'sheettype', width: 15 },
        { header: 'Flip Polarity? (Yes/No)', key: 'flip', width: 20 },
        { header: 'Header Row', key: 'offset', width: 15 },
        { header: 'End Balance', key: 'endbalance', width: 15 },
    ];

    const categories = [
        ['Sales', 'General', 'Income', 'P&L'],
        ['Rent', 'Office', 'Expense', 'P&L'],
        ['Checking Account', 'Bank', 'Asset', 'Balance Sheet'],
        ['Credit Card', 'CC', 'Liability', 'Balance Sheet'],
        ['AX CC', 'Liability', 'Liability', 'Balance Sheet'],
    ];
    setupSheet.addRows(categories);

    // Default Sheet Configurations
    // Row 4 corresponds to 'Checking Account'
    setupSheet.getCell('I4').value = 'Bank Transactions';
    setupSheet.getCell('J4').value = 'Bank';
    setupSheet.getCell('K4').value = 'No';
    setupSheet.getCell('L4').value = 3; // load_transactions.js puts TOTAL/SUBTOTAL in rows 1-2
    // End Balance is required for every bank/card sheet. 0 is right for the empty template;
    // after loading a statement, enter that statement's ending balance.
    setupSheet.getCell('M4').value = 0;

    // Row 5 corresponds to 'Credit Card'
    setupSheet.getCell('I5').value = 'Credit Card Transactions';
    setupSheet.getCell('J5').value = 'CC';
    setupSheet.getCell('K5').value = 'Yes'; // Default to Yes for CC
    setupSheet.getCell('L5').value = 3;
    setupSheet.getCell('M5').value = 0;

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
    try {
        gitSha = execSync('git rev-parse HEAD', { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
    } catch (e) {
        // Not a git checkout; keep 'N/A'
    }

    versionSheet.addRow(['Version ID', pkg.version]);
    versionSheet.addRow(['Git SHA', gitSha]);
    versionSheet.addRow(['Generated At', new Date().toLocaleString()]);

    await workbook.xlsx.writeFile(filename);
    console.log(`Template updated: ${filename}`);
}
createTemplate();

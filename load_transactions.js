// Imports a bank or credit-card export (CSV or Excel) into the accounting workbook.
// Run with --help for usage.
const ExcelJS = require('exceljs');
const fs = require('fs');
const { prepareTargetSheet } = require('./lib/loader/target-sheet');
const { readSourceRecords } = require('./lib/loader/read-source');
const { writeRecords } = require('./lib/loader/write-rows');
const { recordImportHistory } = require('./lib/loader/history');
const { parseOptionValue, stripOption } = require('./lib/cli');
const book = require('./lib/book');
const { didYouMean } = require('./lib/book/suggest');

const MIN_POSITIONALS = 3; // <inputFile> <accountType> <targetTemplate>
const MIN_BOOK_POSITIONALS = 2; // <inputFile> <sheet>

// --book mode: the book names the workbook and the sheets; the dropdowns point at the generated Lists sheet.
// Checks that the workbook is not open in Excel before the export is read.
async function loadWithBook(args) {
    const clearFlag = args.includes('--clear');
    const refreshOnly = args.includes('--refresh-lists');
    const ref = parseOptionValue(args, '--book');
    if (!ref) throw new Error('--book needs a book id or the path of a book file (.toml).');
    const positionals = stripOption(args, '--book').filter(a => !a.startsWith('--'));

    const config = book.loadBook(ref);
    const targetFile = config.workbookPath;
    book.assertWorkbookUnlocked(targetFile);

    let sheet = null;
    if (!refreshOnly) {
        if (positionals.length < MIN_BOOK_POSITIONALS) {
            throw new Error('With --book, give <inputFile> and <sheet>, e.g. node load_transactions.js --book acme "chase_jan.csv" "Chase Checking". Use --refresh-lists alone to rewrite the Lists sheet.');
        }
        const [inputFile, sheetName] = positionals;
        sheet = config.sheets.find(s => s.name === sheetName) || config.sheets.find(s => s.name.toLowerCase() === sheetName.trim().toLowerCase());
        if (!sheet) {
            const names = config.sheets.filter(s => s.type !== 'ledger').map(s => s.name);
            throw new Error(`Sheet "${sheetName}" is not in the book.${didYouMean(sheetName, names)} Sheets: ${names.map(n => `"${n}"`).join(', ')}.`);
        }
        if (sheet.type === 'ledger') throw new Error(`"${sheet.name}" is the ledger; transactions are loaded into bank, card, income or expense sheets.`);
        if (!fs.existsSync(inputFile)) throw new Error(`Input file not found: ${inputFile}`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(targetFile);

    if (sheet) {
        const inputFile = positionals[0];
        const accountType = sheet.type === 'cc' ? 'cc' : 'bank'; // the layout: card sheets have extra columns
        console.log(`Loading ${accountType.toUpperCase()} transactions from ${inputFile} to ${targetFile} (sheet "${sheet.name}")...`);
        if (clearFlag) console.log('  (Option --clear active: Existing data will be removed)');
        const { targetSheet, targetSheetName, headerRowIdx } = prepareTargetSheet(workbook, accountType, clearFlag, { sheetName: sheet.name });
        const { records, globalAccountNum } = await readSourceRecords(inputFile);
        writeRecords(targetSheet, records, { accountType, headerRowIdx, globalAccountNum, targetSheetName, refs: book.listsRefs() });
        recordImportHistory(workbook, { args, inputFile, targetSheetName });
    } else {
        console.log(`Refreshing the Lists sheet in ${targetFile}...`);
    }

    const lists = await book.writeListsSheet(workbook, config);
    console.log(`Lists sheet written: ${lists.categories.length} categories, ${lists.subCategories.length} sub-categories, ${lists.vendors.length} vendors, ${lists.customers.length} customers.`);
    const withDropdowns = book.applyDropdowns(workbook, config);
    console.log(`Dropdowns (Sub-Category follows Category) on: ${withDropdowns.map(n => `"${n}"`).join(', ') || 'no sheets'}.`);

    try {
        await workbook.xlsx.writeFile(targetFile);
        console.log(`Saved changes to ${targetFile}.`);
    } catch (saveError) {
        console.error(`Error saving file: ${saveError.message}`);
        process.exitCode = 1;
    }
}

async function loadTransactions() {
    const args = process.argv.slice(2);
    // Parse flags vs positionals
    const clearFlag = args.includes('--clear');
    const helpFlag = args.includes('--help');
    const positionals = args.filter(a => !a.startsWith('--'));

    if (helpFlag) {
        console.log(`
Usage: node load_transactions.js <inputFile> <accountType> <targetTemplate> [--clear]
       node load_transactions.js --book <id|path> <inputFile> <sheet> [--clear]
       node load_transactions.js --book <id|path> --refresh-lists

Description:
  Imports transactions from a CSV or Excel file into the main accounting workbook.
  It automatically detects columns, formats headers, and appends data as an Excel Table.

Arguments:
  <inputFile>       Path to the source file (CSV or Excel).
  <accountType>     Type of account to load: 'bank' or 'cc' (Credit Card).
                    - 'bank': Mapped to 'Bank Transactions' (or sheet name configured in Setup).
                    - 'cc':   Mapped to 'Credit Card Transactions' (or sheet name configured in Setup).
  <targetTemplate>  Path to the main accounting Excel file (e.g., "My_Books_2025.xlsx").
  <sheet>           (--book) Name of a sheet in the book file, e.g. "Chase Checking".

Flags:
  --help            Show this help message.
  --clear           [WARNING] Clears ALL existing data rows in the target sheet before importing.
                    Use this for fresh imports or re-runs.
  --book <id|path>  Use a book file (see docs/setup.md): the book names the workbook and the sheets.
                    Checks first that the workbook is not open in Excel, then writes the rows and
                    rewrites the generated "Lists" sheet that the dropdowns use (the Setup sheet is not read).
  --refresh-lists   (--book) Only rewrite the "Lists" sheet from the book. Use after editing the book's
                    categories, vendors or customers.

Example:
  node load_transactions.js "e_statements/jan_bank.csv" bank "Books_2025.xlsx"
  node load_transactions.js "e_statements/jan_cc.xlsx" cc "Books_2025.xlsx" --clear
  node load_transactions.js --book acme "e_statements/jan_bank.csv" "Chase Checking"
        `);
        return;
    }

    if (args.some(a => a === '--book' || a.startsWith('--book='))) {
        try {
            await loadWithBook(args);
        } catch (e) {
            console.error(`Error: ${e.message}`);
            process.exitCode = 1;
        }
        return;
    }

    if (positionals.length < MIN_POSITIONALS) {
        console.log('Error: Missing required arguments. Use --help for usage information.');
        return;
    }

    const inputFile = positionals[0];
    const accountType = positionals[1].toLowerCase();
    const targetFile = positionals[2];

    if (!fs.existsSync(inputFile)) { console.error(`Input file not found: ${inputFile}`); return; }
    if (!fs.existsSync(targetFile)) { console.error(`Target file not found: ${targetFile}`); return; }

    console.log(`Loading ${accountType.toUpperCase()} transactions from ${inputFile} to ${targetFile}...`);
    if (clearFlag) console.log('  (Option --clear active: Existing data will be removed)');

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(targetFile);
    } catch (e) {
        if (e.code === 'EBUSY') { console.error(`Error: ${targetFile} is open in Excel. Please close it.`); return; }
        throw e;
    }

    const { targetSheet, targetSheetName, headerRowIdx } = prepareTargetSheet(workbook, accountType, clearFlag);
    const { records, globalAccountNum } = await readSourceRecords(inputFile);
    writeRecords(targetSheet, records, { accountType, headerRowIdx, globalAccountNum, targetSheetName });
    recordImportHistory(workbook, { args, inputFile, targetSheetName });

    try {
        await workbook.xlsx.writeFile(targetFile);
        console.log(`Saved changes to ${targetFile}.`);
    } catch (saveError) {
        console.error(`Error saving file: ${saveError.message}`);
    }
}

loadTransactions();

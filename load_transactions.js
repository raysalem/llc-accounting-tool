// Imports a bank or credit-card export (CSV or Excel) into the accounting workbook.
// Run with --help for usage.
const ExcelJS = require('exceljs');
const fs = require('fs');
const { prepareTargetSheet } = require('./lib/loader/target-sheet');
const { readSourceRecords } = require('./lib/loader/read-source');
const { writeRecords } = require('./lib/loader/write-rows');
const { recordImportHistory } = require('./lib/loader/history');

async function loadTransactions() {
    const args = process.argv.slice(2);
    // Parse flags vs positionals
    const clearFlag = args.includes('--clear');
    const helpFlag = args.includes('--help');
    const positionals = args.filter(a => !a.startsWith('--'));

    if (helpFlag) {
        console.log(`
Usage: node load_transactions.js <inputFile> <accountType> <targetTemplate> [--clear]

Description:
  Imports transactions from a CSV or Excel file into the main accounting workbook.
  It automatically detects columns, formats headers, and appends data as an Excel Table.

Arguments:
  <inputFile>       Path to the source file (CSV or Excel).
  <accountType>     Type of account to load: 'bank' or 'cc' (Credit Card).
                    - 'bank': Mapped to 'Bank Transactions' (or sheet name configured in Setup).
                    - 'cc':   Mapped to 'Credit Card Transactions' (or sheet name configured in Setup).
  <targetTemplate>  Path to the main accounting Excel file (e.g., "My_Books_2025.xlsx").

Flags:
  --help            Show this help message.
  --clear           [WARNING] Clears ALL existing data rows in the target sheet before importing.
                    Use this for fresh imports or re-runs.

Example:
  node load_transactions.js "e_statements/jan_bank.csv" bank "Books_2025.xlsx"
  node load_transactions.js "e_statements/jan_cc.xlsx" cc "Books_2025.xlsx" --clear
        `);
        return;
    }

    if (positionals.length < 3) {
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

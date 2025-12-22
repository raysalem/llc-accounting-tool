const ExcelJS = require('exceljs');
const fs = require('fs');
const readline = require('readline');

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

    // --- Find Target Sheet Name from Setup (Resiliency) ---
    let targetSheetName = accountType === 'cc' ? 'Credit Card Transactions' : 'Bank Transactions';
    const setupSheet = workbook.getWorksheet('Setup');
    if (setupSheet) {
        setupSheet.eachRow((row, r) => {
            if (r === 1) return;
            const sName = row.getCell(9).value; // Col I
            const sType = (row.getCell(10).value || '').toString().toLowerCase(); // Col J
            if (sName && (sType === accountType || (accountType === 'bank' && sType.includes('bank')) || (accountType === 'cc' && sType.includes('cc')))) {
                targetSheetName = sName.toString();
            }
        });
    }

    let targetSheet = workbook.getWorksheet(targetSheetName);

    if (!targetSheet) {
        console.log(`  Target sheet '${targetSheetName}' not found. Creating it...`);
        targetSheet = workbook.addWorksheet(targetSheetName);
        // Initialize headers if adding new sheet
        if (accountType === 'cc') {
            targetSheet.columns = [
                { header: 'Date', key: 'date', width: 12 },
                { header: 'Member', key: 'member', width: 15 },
                { header: 'Description', key: 'desc', width: 35 },
                { header: 'Amount', key: 'amount', width: 15 },
                { header: 'Category', key: 'category', width: 20 },
                { header: 'Sub-Category', key: 'subcategory', width: 20 },
                { header: 'Extended Details', key: 'extended', width: 30 },
                { header: 'Vendor', key: 'vendor', width: 20 },
                { header: 'Customer', key: 'customer', width: 20 },
                { header: 'Account #', key: 'account', width: 15 },
                { header: 'Receipt', key: 'receipt', width: 10 },
                { header: 'Report Type (Auto)', key: 'report_type', width: 15 },
            ];
        } else {
            targetSheet.columns = [
                { header: 'Date', key: 'date', width: 12 },
                { header: 'Description', key: 'desc', width: 35 },
                { header: 'Amount', key: 'amount', width: 15 },
                { header: 'Category', key: 'category', width: 20 },
                { header: 'Sub-Category', key: 'subcategory', width: 20 },
                { header: 'Extended Details', key: 'extended', width: 30 },
                { header: 'Vendor', key: 'vendor', width: 20 },
                { header: 'Customer', key: 'customer', width: 20 },
                { header: 'Report Type (Auto)', key: 'report_type', width: 15 },
            ];
        }
    }

    // Clear Logic (Metadata-preserving)
    if (clearFlag) {
        console.log(`  Clearing existing data in '${targetSheetName}'...`);
        // Remove all rows except the header
        if (targetSheet.rowCount > 1) {
            targetSheet.spliceRows(2, targetSheet.rowCount - 1);
        }
    }

    // --- Adjust Layout for Top Totals ---
    // Check if header is at Row 1 (standard template) or Row 3 (already adjusted)
    const firstRowVals = targetSheet.getRow(1).values;
    const isHeaderAtTop = firstRowVals.includes('Date') || (firstRowVals[1] && firstRowVals[1].includes('Date'));

    let headerRowIdx = isHeaderAtTop ? 1 : 3;

    if (isHeaderAtTop) {
        console.log('  Adjusting layout: Inserting 2 rows at top for Totals...');
        targetSheet.spliceRows(1, 0, [], []);
        headerRowIdx = 3;
    }

    // Set Formulas
    const amtCol = accountType === 'cc' ? 'D' : 'C'; // Amount Column Letter
    const startRow = headerRowIdx + 1;
    const maxRow = 10000; // Arbitrary large number for range

    targetSheet.getCell('A1').value = 'TOTAL';
    targetSheet.getCell('A1').font = { bold: true };
    targetSheet.getCell(`${amtCol}1`).value = { formula: `SUM(${amtCol}${startRow}:${amtCol}${maxRow})` };
    targetSheet.getCell(`${amtCol}1`).font = { bold: true };

    targetSheet.getCell('A2').value = 'SUBTOTAL (Filtered)';
    targetSheet.getCell('A2').font = { bold: true, italic: true };
    targetSheet.getCell(`${amtCol}2`).value = { formula: `SUBTOTAL(109, ${amtCol}${startRow}:${amtCol}${maxRow})` };
    targetSheet.getCell(`${amtCol}2`).font = { bold: true, italic: true };

    // Explicitly disable worksheet-level autofilter to avoid conflicts with Table-level filters
    targetSheet.autoFilter = null;

    const records = [];

    // --- CSV Parsing ---
    if (inputFile.toLowerCase().endsWith('.csv')) {
        const fileStream = fs.createReadStream(inputFile);
        const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });

        let headers = [];
        let isFirstLine = true;

        for await (const line of rl) {
            const cleanValues = (line.match(/(?:^|,)(\"(?:[^\"]+|\"\")*\"|[^,]*)/g) || []).map(v => {
                v = v.replace(/^,/, '');
                if (v.startsWith('"') && v.endsWith('"')) return v.slice(1, -1);
                return v;
            });

            if (isFirstLine) {
                headers = cleanValues.map(h => h.trim().toLowerCase());
                isFirstLine = false;
                continue;
            }

            const record = {};
            const dateIdx = headers.indexOf('date');
            const nameIdx = headers.indexOf('name') !== -1 ? headers.indexOf('name') : headers.indexOf('description');
            // Fallback for description if 'name' not found (common in bank csvs vs quickbooks)
            const memoIdx = headers.indexOf('memo');
            const amtIdx = headers.indexOf('amount');

            if (dateIdx !== -1) record.date = cleanValues[dateIdx];
            if (nameIdx !== -1) record.desc = cleanValues[nameIdx];
            if (memoIdx !== -1) record.extended = cleanValues[memoIdx];
            if (amtIdx !== -1) {
                let amtStr = cleanValues[amtIdx];
                if (amtStr) amtStr = amtStr.replace(/[$,]/g, '');
                record.amount = amtStr;
            }
            if (record.date) records.push(record);
        }

    } else {
        // --- Excel Parsing ---
        const inputWorkbook = new ExcelJS.Workbook();
        await inputWorkbook.xlsx.readFile(inputFile);
        const inputSheet = inputWorkbook.worksheets[0];

        let colMap = {};
        let headerRowIndex = 1;

        inputSheet.eachRow((row, rowNumber) => {
            if (Object.keys(colMap).length > 0) return;
            const values = (row.values || []).map(v => v ? v.toString().trim().toLowerCase() : '');
            if (values.includes('date') && (values.includes('amount') || values.includes('description'))) {
                headerRowIndex = rowNumber;
                row.eachCell((cell, colNumber) => {
                    const v = cell.value ? cell.value.toString().trim().toLowerCase() : '';
                    colMap[v] = colNumber;
                });
            }
        });

        if (Object.keys(colMap).length === 0 && accountType === 'cc') {
            headerRowIndex = 7;
            colMap = {
                'date': 1, 'receipt': 2, 'description': 3, 'card member': 4,
                'account #': 5, 'amount': 6, 'extended details': 7
            };
        }

        inputSheet.eachRow((row, rowNumber) => {
            if (rowNumber <= headerRowIndex) return;

            const getVal = (key) => {
                let idx = colMap[key];
                if (!idx) idx = colMap[Object.keys(colMap).find(k => k.includes(key))];
                if (!idx) return '';
                return row.getCell(idx).value;
            };

            const rec = {
                date: getVal('date'),
                desc: getVal('description'),
                amount: getVal('amount'),
                member: getVal('card member') || getVal('member'),
                extended: getVal('extended details'),
                receipt: getVal('receipt'),
                account: getVal('account') || getVal('account #')
            };

            // Enhanced Junk Filter
            if (!rec.date) return;
            // Filter out obviously empty rows (desc & amount missing)
            if ((!rec.desc || rec.desc.trim() === '') && (!rec.amount || parseFloat(rec.amount) === 0)) return;
            // Filter out summary/total rows often found in exports
            if (rec.desc && /total|balance|sum/i.test(rec.desc)) return;

            if (rec.date) records.push(rec);
        });
    }

    // --- Prepare Data Rows ---
    const rowsToAdd = [];
    records.forEach(rec => {
        let dateVal = rec.date;
        if (typeof dateVal === 'string') dateVal = new Date(dateVal);

        let newRow = [];
        if (accountType === 'cc') {
            // [1]Date [2]Member [3]Desc [4]Amount [5]Cat [6]Sub [7]Ext [8]Vend [9]Cust [10]Acct [11]Rec
            newRow = [
                dateVal,
                rec.member || '',
                rec.desc || '',
                parseFloat(rec.amount) || 0,
                '', '', // Cat, Sub
                rec.extended || '',
                rec.vendor || '', rec.customer || '', // Vend, Cust
                rec.account || '',
                rec.receipt || ''
            ];
        } else {
            // [1]Date [2]Desc [3]Amount [4]Cat [5]Sub [6]Ext [7]Vend [8]Cust
            newRow = [
                dateVal,
                rec.desc || '',
                parseFloat(rec.amount) || 0,
                '', '', // Cat, Sub
                rec.extended || '',
                rec.vendor || '', rec.customer || '', // Vendor, Cust
            ];
        }
        rowsToAdd.push(newRow);
    });

    // --- Table Management ---
    // Check if a table already exists to avoid corruption
    const hasTable = targetSheet.tables && Object.keys(targetSheet.tables).length > 0;

    if (hasTable) {
        console.log('  Existing Excel Table detected. Appending rows to it...');
        // Just add rows normally, Excel should expand the table
        targetSheet.addRows(rowsToAdd);
    } else {
        console.log('  No Excel Table found. Creating one...');
        // Create table with data
        // Define columns matching the sheet structure request
        let tableCols = [];
        if (accountType === 'cc') {
            tableCols = [
                { name: 'Date' }, { name: 'Member' }, { name: 'Description' }, { name: 'Amount' },
                { name: 'Category' }, { name: 'Sub-Category' }, { name: 'Extended Details' },
                { name: 'Vendor' }, { name: 'Customer' }, { name: 'Account #' }, { name: 'Receipt' },
                { name: 'Report Type (Auto)' }
            ];
        } else {
            tableCols = [
                { name: 'Date' }, { name: 'Description' }, { name: 'Amount' },
                { name: 'Category' }, { name: 'Sub-Category' }, { name: 'Extended Details' },
                { name: 'Vendor' }, { name: 'Customer' }, { name: 'Report Type (Auto)' }
            ];
        }

        // We need to write the data differently for addTable
        // addTable expects 'rows' as array of arrays, and it writes everything relative to 'ref'
        // Since we already have the Header at Row 3 (headerRowIdx), we put the table there.
        // However, addTable writes the header too. We should overwrite the existing header to be safe or ensure it matches.

        targetSheet.addTable({
            name: `Table_${targetSheetName.replace(/\s/g, '')}_${Date.now()}`,
            ref: `A${headerRowIdx}`,
            headerRow: true,
            totalsRow: false,
            style: {
                theme: 'TableStyleMedium2',
                showRowStripes: true,
            },
            columns: tableCols,
            rows: rowsToAdd
        });
    }

    console.log(`Successfully processed ${rowsToAdd.length} transactions in ${targetSheetName}.`);

    // Apply Validation & Formulas (Post-Insert)
    // Adjust start row for loop: header is at 3, data starts at 4
    const validationStartRow = headerRowIdx + 1;
    const finalMaxRow = targetSheet.rowCount;

    for (let i = validationStartRow; i <= finalMaxRow; i++) {
        if (accountType === 'cc') {
            targetSheet.getCell(`E${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$A$2:$A$100'] };
            targetSheet.getCell(`F${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$B$2:$B$100'] };
            targetSheet.getCell(`H${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$F$2:$F$100'] };
            targetSheet.getCell(`I${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$G$2:$G$100'] };
            targetSheet.getCell(`L${i}`).value = { formula: `IFERROR(VLOOKUP(E${i},Setup!A:D,4,FALSE), "")` };
        } else {
            targetSheet.getCell(`D${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$A$2:$A$100'] };
            targetSheet.getCell(`E${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$B$2:$B$100'] };
            targetSheet.getCell(`G${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$F$2:$F$100'] };
            targetSheet.getCell(`H${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$G$2:$G$100'] };
            targetSheet.getCell(`I${i}`).value = { formula: `IFERROR(VLOOKUP(D${i},Setup!A:D,4,FALSE), "")` };
        }
    }

    // --- Record History in VERSION tab ---
    let versionSheet = workbook.getWorksheet('VERSION');
    if (!versionSheet) {
        versionSheet = workbook.addWorksheet('VERSION');
    }

    // Add a marker if not present
    let historyHeaderFound = false;
    versionSheet.eachRow(row => {
        if (row.getCell(1).value === '--- Import History ---') historyHeaderFound = true;
    });

    if (!historyHeaderFound) {
        versionSheet.addRow([]);
        versionSheet.addRow(['--- Import History ---', '']);
    }

    // Log as a multi-line value in the second column to preserve 2-column layout
    const timestamp = new Date().toLocaleString();
    const historyDetail = [
        `Command: node load_transactions.js ${args.join(' ')}`,
        `Input: ${inputFile}`,
        `Target Sheet: ${targetSheetName}`
    ].join('\n');

    versionSheet.addRow([`Import at ${timestamp}`, historyDetail]);

    // Auto-fit height for the new row if possible (or just let Excel handle it)
    const lastRow = versionSheet.lastRow;
    if (lastRow) {
        lastRow.alignment = { wrapText: true, vertical: 'top' };
    }

    try {
        await workbook.xlsx.writeFile(targetFile);
        console.log(`Saved changes to ${targetFile}.`);
    } catch (saveError) {
        console.error(`Error saving file: ${saveError.message}`);
    }
}

loadTransactions();

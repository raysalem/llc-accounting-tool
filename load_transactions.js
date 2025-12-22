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

    // --- Strict Clear Logic: Delete & Recreate ---
    // This is the "Nuclear Option" to guarantee a 100% clean state (no lingering categories or metadata).
    if (clearFlag) {
        const existingSheet = workbook.getWorksheet(targetSheetName);
        if (existingSheet) {
            console.log(`  [CLEAR] 'Nuclear' option active: Deleting sheet '${targetSheetName}' to ensure a clean start.`);
            workbook.removeWorksheet(existingSheet.id);
        }
    }

    // --- Target Sheet Setup ---
    let targetSheet = workbook.getWorksheet(targetSheetName);

    if (!targetSheet) {
        console.log(`  Creating sheet '${targetSheetName}'...`);
        targetSheet = workbook.addWorksheet(targetSheetName);
        // Define Headers only for new sheets
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
                { header: 'Account Number', key: 'account', width: 20 },
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
    // targetSheet.autoFilter = null; // Removed to prevent corruption

    // --- Global Account Number Detection (Source File) ---
    // 1. Try Filename
    let globalAccountNum = '';
    const filenameMatch = inputFile.match(/[-_ ](\d{4,})[.]/); // e.g., "- 81002."
    if (filenameMatch) globalAccountNum = filenameMatch[1];

    const records = [];

    // --- CSV Parsing ---
    if (inputFile.toLowerCase().endsWith('.csv')) {
        const fileStream = fs.createReadStream(inputFile);
        const rl = readline.createInterface({ input: fileStream, crlfDelay: Infinity });

        let headers = [];
        let isFirstLine = true;

        for await (const line of rl) {
            // Check top lines for Account Number if not found yet
            if (!globalAccountNum && /account\s*(?:number|#)/i.test(line)) {
                const match = line.match(/(?:number|#)[:\s]*([\d-]+)/i);
                if (match) globalAccountNum = match[1];
            }

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
            const memoIdx = headers.indexOf('memo');
            const amtIdx = headers.indexOf('amount');
            const acctIdx = headers.indexOf('account') !== -1 ? headers.indexOf('account') : headers.indexOf('account number');

            if (dateIdx !== -1) record.date = cleanValues[dateIdx];
            if (nameIdx !== -1) record.desc = cleanValues[nameIdx];
            if (memoIdx !== -1) record.extended = cleanValues[memoIdx];

            // Prefer row-level, fallback to global
            if (acctIdx !== -1 && cleanValues[acctIdx]) record.account = cleanValues[acctIdx];
            else if (globalAccountNum) record.account = globalAccountNum;

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

        // Scan top rows for "Account Number" label in Source
        if (!globalAccountNum) {
            for (let r = 1; r <= 10; r++) {
                const row = inputSheet.getRow(r);
                row.eachCell((cell) => {
                    const val = (cell.value || '').toString();
                    if (/account\s*(?:number|#)/i.test(val)) {
                        // Check this cell or next cell for digits
                        const numMatch = val.match(/(?:number|#)[:\s]*([\d-]+)/i);
                        if (numMatch) globalAccountNum = numMatch[1];
                        else {
                            // Valid next cell?
                            const nextVal = (row.getCell(cell.col + 1).value || '').toString();
                            if (/[\d-]+/.test(nextVal)) globalAccountNum = nextVal;
                        }
                    }
                });
                if (globalAccountNum) break;
            }
        }

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
                'account #': 5, 'amount': 6, 'extended details': 7,
                'account number': 5
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
                account: getVal('account') || getVal('account #') || getVal('account number') || globalAccountNum
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

    // --- Insert Data ---
    targetSheet.addRows(rowsToAdd);

    // --- Visual Table & Formatting (Avoid XML Corruption) ---
    // Instead of addTable (which conflicts with updates), we apply AutoFilter and Styling manually.
    const finalLastRow = targetSheet.rowCount;
    if (finalLastRow >= headerRowIdx) {
        // Apply AutoFilter to the range
        // Note: targetSheet.columnCount might be excessive, restrict to data width
        const lastCol = accountType === 'cc' ? 12 : 9;

        targetSheet.autoFilter = {
            from: { row: headerRowIdx, column: 1 },
            to: { row: finalLastRow, column: lastCol }
        };

        // Header Styling
        const headerRow = targetSheet.getRow(headerRowIdx);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4F81BD' } // Standard Excel Blue
        };

        // Display Global Account Number in Top Row (J1) if found
        if (globalAccountNum && accountType === 'cc') {
            const acctCell = targetSheet.getCell('J1');
            acctCell.value = `Acct: ${globalAccountNum}`;
            acctCell.font = { bold: true, color: { argb: 'FF000000' } };
        }

        // Border Styling for Data
        for (let r = headerRowIdx + 1; r <= finalLastRow; r++) {
            const row = targetSheet.getRow(r);
            // row.border = { bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } } }; // Light grey border
        }
    }

    console.log(`Successfully processed ${rowsToAdd.length} transactions in ${targetSheetName}.`);
    console.log(`  (Note: Applied AutoFilter and Sytling. Formal Excel 'Tables' disabled to prevent file corruption on update.)`);

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

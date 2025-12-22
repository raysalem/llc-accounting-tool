const ExcelJS = require('exceljs');
const fs = require('fs');
const readline = require('readline');

async function loadTransactions() {
    const args = process.argv.slice(2);
    // Parse flags vs positionals
    const clearFlag = args.includes('--clear');
    const positionals = args.filter(a => !a.startsWith('--'));

    if (positionals.length < 3) {
        console.log('Usage: node load_transactions.js <inputFile> <accountType> <targetTemplate> [--clear]');
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

    const targetSheetName = accountType === 'cc' ? 'Credit Card Transactions' : 'Bank Transactions';
    let targetSheet = workbook.getWorksheet(targetSheetName);

    if (!targetSheet) { console.error(`Target sheet '${targetSheetName}' not found.`); return; }

    // Clear Logic (Metadata-preserving)
    if (clearFlag) {
        console.log(`  Clearing existing data in '${targetSheetName}'...`);
        // Remove all rows except the header
        if (targetSheet.rowCount > 1) {
            targetSheet.spliceRows(2, targetSheet.rowCount - 1);
        }
    }

    if (accountType === 'cc') {
        // [1]Date [2]Member [3]Desc [4]Amount [5]Cat [6]Sub [7]Ext [8]Vend [9]Cust [10]Acct [11]Rec [12]Report
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
        // [1]Date [2]Desc [3]Amount [4]Cat [5]Sub [6]Ext [7]Vend [8]Cust [9]Report
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
            const nameIdx = headers.indexOf('name');
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
            if (rec.date) records.push(rec);
        });
    }

    // Append Records
    let addedCount = 0;
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
                '', '', // Vendor, Cust
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
                '', '', // Vendor, Cust
            ];
        }

        targetSheet.addRow(newRow);
        addedCount++;
    });

    console.log(`Successfully added ${addedCount} transactions to ${targetSheetName}.`);

    // Apply Validation & Formulas (Post-Insert)
    const maxRow = Math.max(500, targetSheet.rowCount);
    for (let i = 2; i <= maxRow; i++) {
        if (accountType === 'cc') {
            targetSheet.getCell(`E${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$A$2:$A$100'] };
            targetSheet.getCell(`F${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$B$2:$B$100'] }; // Sub-Cat
            targetSheet.getCell(`H${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$F$2:$F$100'] };
            targetSheet.getCell(`I${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$G$2:$G$100'] };
            targetSheet.getCell(`L${i}`).value = { formula: `IFERROR(VLOOKUP(E${i},Setup!A:D,4,FALSE), "")` };
        } else {
            // Bank: Cat[4], Sub[5], Vend[7], Cust[8], Formula[9]
            targetSheet.getCell(`D${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$A$2:$A$100'] };
            targetSheet.getCell(`E${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['Setup!$B$2:$B$100'] }; // Sub-Cat
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

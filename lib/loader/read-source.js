// Reads transactions from a bank/credit-card export (CSV or Excel) into plain
// records: { date, desc, amount, member, extended, receipt, account }.
const ExcelJS = require('exceljs');
const fs = require('fs');
const readline = require('readline');
const { parseAmount } = require('../accounting');

// Returns { records, globalAccountNum }. The account number comes from the
// file name (e.g. "Checking - 81002.csv") or an "Account Number" line in the file.
async function readSourceRecords(inputFile) {
    // --- Global Account Number Detection (Source File) ---
    // 1. Try Filename
    let globalAccountNum = '';
    const filenameMatch = inputFile.match(/[-_ ](\d{4,})[.]/); // e.g., "- 81002."
    if (filenameMatch) globalAccountNum = filenameMatch[1];

    if (inputFile.toLowerCase().endsWith('.csv')) {
        return readCsv(inputFile, globalAccountNum);
    }
    return readExcel(inputFile, globalAccountNum);
}

async function readCsv(inputFile, initialAccountNum) {
    let globalAccountNum = initialAccountNum;
    const records = [];

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

    return { records, globalAccountNum };
}

async function readExcel(inputFile, initialAccountNum) {
    let globalAccountNum = initialAccountNum;
    const records = [];

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

    // --- Constants for Column Headers ---
    // (Shared logic with update_financials for consistency)
    const HEADERS = {
        DATE: ['date', 'txn date', 'transaction date'],
        DESC: ['description', 'desc', 'payee', 'merchant', 'name'],
        AMOUNT: ['amount', 'amt', 'value'],
        ACCOUNT: ['account', 'account #', 'account number'],
        MEMBER: ['card member', 'member'],
        EXTENDED: ['extended details', 'memo', 'extended'],
        RECEIPT: ['receipt']
    };

    const findCol = (cellVal, headerList) => {
        if (!cellVal) return false;
        const v = cellVal.toString().toLowerCase().trim();
        return headerList.some(h => v === h || v.includes(h));
    };

    let colMap = {};
    let headerRowIndex = 1;

    inputSheet.eachRow((row, rowNumber) => {
        if (Object.keys(colMap).length > 0) return;
        const values = (row.values || []).map(v => v ? v.toString().trim().toLowerCase() : '');

        // Robust Header Detection
        const hasDate = values.some(v => findCol(v, HEADERS.DATE));
        const hasAmount = values.some(v => findCol(v, HEADERS.AMOUNT));
        const hasDesc = values.some(v => findCol(v, HEADERS.DESC));

        if (hasDate && (hasAmount || hasDesc)) {
            headerRowIndex = rowNumber;
            row.eachCell((cell, colNumber) => {
                const v = cell.value ? cell.value.toString().trim().toLowerCase() : '';
                colMap[v] = colNumber;
            });
        }
    });

    if (Object.keys(colMap).length === 0) {
        console.log('Warning: No valid header row found. Please ensure the file has "Date" and "Amount" columns.');
    }

    inputSheet.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRowIndex) return;

        const getVal = (headerKeys) => {
            // Try strict match first, then constants match
            for (const key of headerKeys) {
                if (colMap[key]) return row.getCell(colMap[key]).value;
                // Look for fuzzy match in colMap keys
                const foundKey = Object.keys(colMap).find(k => k.includes(key) || key.includes(k));
                if (foundKey) return row.getCell(colMap[foundKey]).value;
            }
            return '';
        };

        // Helper wrapper for single key or list
        const getField = (headerList) => {
            // Find column index that matches headerList
            const foundColName = Object.keys(colMap).find(k => findCol(k, headerList));
            if (foundColName) return row.getCell(colMap[foundColName]).value;
            return '';
        };

        const rec = {
            date: getField(HEADERS.DATE),
            desc: getField(HEADERS.DESC),
            amount: getField(HEADERS.AMOUNT),
            member: getField(HEADERS.MEMBER),
            extended: getField(HEADERS.EXTENDED),
            receipt: getField(HEADERS.RECEIPT),
            account: getField(HEADERS.ACCOUNT) || globalAccountNum
        };

        // Enhanced Junk Filter
        if (!rec.date) return;
        // Filter out obviously empty rows (desc & amount missing)
        if ((!rec.desc || rec.desc.trim() === '') && (!rec.amount || parseAmount(rec.amount) === 0)) return;
        // Filter out summary/total rows often found in exports
        if (rec.desc && /total|balance|sum/i.test(rec.desc)) return;

        if (rec.date) records.push(rec);
    });

    return { records, globalAccountNum };
}

module.exports = { readSourceRecords };

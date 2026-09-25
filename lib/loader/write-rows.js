// Appends imported records to the target sheet and applies the header styling,
// AutoFilter, category/vendor dropdowns and the Report Type lookup formula.
const { parseAmount, parseDateUTC } = require('../accounting');

// Returns the number of rows written.
function writeRecords(targetSheet, records, { accountType, headerRowIdx, globalAccountNum, targetSheetName }) {
    // --- Prepare Data Rows ---
    const rowsToAdd = [];
    records.forEach(rec => {
        let dateVal = rec.date;
        if (typeof dateVal === 'string') dateVal = parseDateUTC(dateVal);

        let newRow = [];
        if (accountType === 'cc') {
            // [1]Date [2]Member [3]Desc [4]Amount [5]Cat [6]Sub [7]Ext [8]Vend [9]Cust [10]Acct [11]Rec
            newRow = [
                dateVal,
                rec.member || '',
                rec.desc || '',
                parseAmount(rec.amount) || 0,
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
                parseAmount(rec.amount) || 0,
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

    return rowsToAdd.length;
}

module.exports = { writeRecords };

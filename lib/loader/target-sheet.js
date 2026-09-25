// Finds (or creates) the sheet that imported transactions go into, applies
// --clear, and puts TOTAL / SUBTOTAL formulas in rows 1-2 above the header.

// Returns { targetSheet, targetSheetName, headerRowIdx }.
function prepareTargetSheet(workbook, accountType, clearFlag) {
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


    return { targetSheet, targetSheetName, headerRowIdx };
}

module.exports = { prepareTargetSheet };

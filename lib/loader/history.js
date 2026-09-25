// Records each import (command, input file, target sheet) in the VERSION sheet.

function recordImportHistory(workbook, { args, inputFile, targetSheetName }) {
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
}

module.exports = { recordImportHistory };

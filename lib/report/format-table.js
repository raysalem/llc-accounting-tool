// Prints a report table with one column per sheet: ending balance, then
// additions and subtractions by sheet, with optional sub-category lines.
function printDetailedTable(title, rows, sheetList, sheetNameMap = {}, label = "Category", forceOpening = false) {
    console.log(`\n--- ${title} ---`);
    if (!rows.length) { console.log('(No Data)'); return; }

    const LABEL_WIDTH = 30;
    const COL_WIDTH = 12;

    // Header 1
    const SECTION_WIDTH = (sheetList.length + 1) * COL_WIDTH;

    const hasOpening = forceOpening || rows.some(r => typeof r.opening === 'number' && Math.abs(r.opening) > 0);

    const h1 = " ".repeat(LABEL_WIDTH) +
        (hasOpening ? "Opening".padStart(COL_WIDTH) + " " : "") +
        "Ending".padStart(COL_WIDTH) + " | " +
        "Additions".padStart(SECTION_WIDTH) + " | " +
        "Subtractions".padStart(SECTION_WIDTH);
    console.log(h1);

    // Header 2
    let h2 = label.padEnd(LABEL_WIDTH);

    // Opening column
    if (hasOpening) h2 += "Balance".padStart(COL_WIDTH) + " ";
    // Net (Grand Total)
    h2 += "Balance".padStart(COL_WIDTH) + " | ";

    // Additions Columns - use short names
    sheetList.forEach(s => {
        const displayName = sheetNameMap[s] || s;
        h2 += displayName.substring(0, COL_WIDTH - 1).padStart(COL_WIDTH);
    });
    h2 += "Total".padStart(COL_WIDTH) + " | ";
    // Subtractions Columns
    sheetList.forEach(s => {
        const displayName = sheetNameMap[s] || s;
        h2 += displayName.substring(0, COL_WIDTH - 1).padStart(COL_WIDTH);
    });
    h2 += "Total".padStart(COL_WIDTH);

    console.log(h2);
    console.log("-".repeat(h2.length));

    rows.forEach(r => {
        let line = r.label.substring(0, LABEL_WIDTH - 1).padEnd(LABEL_WIDTH);
        const sheets = r.sheets || {};

        // Opening
        if (hasOpening) {
            const openVal = r.opening || 0;
            line += openVal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) + " ";
        }

        // Net / Ending
        const signedNet = r.value;
        line += signedNet.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) + " | ";

        // Additions
        let rowAddTotal = 0;
        sheetList.forEach(s => {
            const val = (sheets[s] ? sheets[s].add : 0) || 0;
            rowAddTotal += Math.abs(val);
            // Use absolute for display in additions column
            const disp = Math.abs(val);
            line += disp === 0 ? " ".padStart(COL_WIDTH) : disp.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH);
        });
        line += Math.abs(rowAddTotal).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) + " | ";

        // Subtractions
        let rowSubTotal = 0;
        sheetList.forEach(s => {
            const val = (sheets[s] ? sheets[s].sub : 0) || 0;
            rowSubTotal += Math.abs(val);
            // Use absolute for display in subtractions column
            const disp = Math.abs(val);
            line += disp === 0 ? " ".padStart(COL_WIDTH) : disp.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH);
        });
        line += Math.abs(rowSubTotal).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH);

        console.log(line);

        // Print subcategory breakdowns if available
        if (r.subCats && Object.keys(r.subCats).length > 0) {
            Object.entries(r.subCats).forEach(([subName, subTotal]) => {
                // Show "(No Sub-Cat)" only when other sub-categories exist, so the lines add up to the total
                if (subName === '(No Sub-Cat)' && Object.keys(r.subCats).length === 1) return;
                const subLine = `  > ${subName}`.substring(0, LABEL_WIDTH - 1).padEnd(LABEL_WIDTH) +
                    subTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) +
                    " ".repeat(h2.length - LABEL_WIDTH - COL_WIDTH); // Fill rest with spaces
                console.log(subLine);
            });
        }
    });
}

module.exports = { printDetailedTable };

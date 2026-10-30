// Finds the header row of a sheet and maps its column names. Used for the Setup sheet and vendor.xlsx.
const { getVal } = require('./cells');

// Finds the header row (the row in the first 20 with the most known header names)
// and maps each normalized header name to its column numbers. Also used for vendor.xlsx.
function getHeaderMap(sheet, showDebug) {
    let bestRowIdx = 1;
    let bestMap = new Map();
    let maxFound = -1;
    const lookups = [
        'category', 'subcategory', 'vendors', 'vendor', 'sheetname', 'sheetnameconfig', 'report', 'linkcategory', 'linkcat',
        'name', 'fullname', 'firstname', 'lastname', 'address', 'city', 'state', 'zip', 'tin', 'ssn', 'ein', 'taxid',
        '1099', '1099type', '1099required', '1099nec', 'businessname', 'business', 'flippolarity'
    ];

    for (let ri = 1; ri <= 20; ri++) { // Scan more rows
        const currentMap = new Map();
        const row = sheet.getRow(ri);
        let foundCount = 0;
        row.eachCell((cell, colNumber) => {
            const rawVal = getVal(cell).toString().trim();
            const cleanVal = rawVal.toLowerCase().replace(/[^a-z0-9]/g, '');
            if (cleanVal) {
                if (!currentMap.has(cleanVal)) currentMap.set(cleanVal, []);
                currentMap.get(cleanVal).push(colNumber);
                if (lookups.includes(cleanVal)) foundCount++;
            }
        });
        if (foundCount > maxFound) {
            maxFound = foundCount;
            bestRowIdx = ri;
            bestMap = currentMap;
        }
        if (foundCount >= 4) break; // Higher threshold for "Header Row"
    }
    if (showDebug) {
        console.log(`[DEBUG] Header scanning on "${sheet.name}" complete. Best row: ${bestRowIdx}. Found headers: ${Array.from(bestMap.keys()).join(', ')}`);
    }
    return { map: bestMap, headerRow: bestRowIdx };
}

module.exports = { getHeaderMap };

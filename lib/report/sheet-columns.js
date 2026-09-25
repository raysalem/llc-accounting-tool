// Finds the header row of a transaction sheet and maps its columns (date,
// description, amount, category, sub-category, vendor, customer).
const { HEADERS, bankMapDefault, ccMapDefault, findCol, HEAD_MATCH } = require('./columns');
const { getVal } = require('./cells');

// Returns the column map: { date, desc, amount, category, subCat, vendor, customer }.
function detectSheetColumns(sheet, config, isCC, { showDebug, showChecker }) {
    // Dynamic Map detection
    let headerRowIndex = 1; // Default
    const configOffset = parseInt(config.offset);

    if (!isNaN(configOffset) && configOffset > 0) {
        // Golden Rule: User Input is Mandatory
        headerRowIndex = configOffset;
        if (showChecker) console.log(`[Config] Sheet "${config.name}" using specified Header Row: ${headerRowIndex}`);
    } else {
        // Fallback: Scan rows 1-5 only if no offset provided
        for (let r = 1; r <= 5; r++) {
            const rowVals = sheet.getRow(r).values;
            if (Array.isArray(rowVals)) {
                const rowStr = rowVals.map(v => v ? v.toString().toLowerCase() : '').join(' ');
                if (HEAD_MATCH(rowStr)) {
                    headerRowIndex = r;
                    config.offset = r;
                    break;
                }
            }
        }
    }

    const headerRow = sheet.getRow(headerRowIndex);
    const map = isCC ? { ...ccMapDefault } : { ...bankMapDefault };

    const headerNames = {}; // Store mapped header names for debugging

    headerRow.eachCell((cell, colNumber) => {
        const val = getVal(cell);
        const vLower = val.toString().toLowerCase().trim();

        if (showDebug && sheet.name.includes('Credit Card')) {
            console.log(`[DEBUG CC] Col ${colNumber}: "${val}" (Clean: "${vLower}")`);
        }

        if (findCol(val, HEADERS.DATE)) { map.date = colNumber; headerNames.date = val; }
        else if (findCol(val, HEADERS.DESC)) { map.desc = colNumber; headerNames.desc = val; }
        else if (findCol(val, HEADERS.AMOUNT)) { map.amount = colNumber; headerNames.amount = val; }
        else if (findCol(val, HEADERS.SUBCAT)) { map.subCat = colNumber; headerNames.subCat = val; }
        else if (findCol(val, HEADERS.CATEGORY)) { map.category = colNumber; headerNames.category = val; }
        else if (findCol(val, HEADERS.VENDOR)) {
            // Harden against partial matches like "Vendor Address" or "Vendor Phone"
            if (!vLower.includes('address') && !vLower.includes('phone') && !vLower.includes('email') && !vLower.includes(' id') && !vLower.includes('zip') && !vLower.includes('state') && !vLower.includes('city')) {
                map.vendor = colNumber;
                headerNames.vendor = val;
            }
        }
        else if (findCol(val, HEADERS.CUSTOMER)) {
            if (!vLower.includes('address') && !vLower.includes('phone') && !vLower.includes('email') && !vLower.includes(' id')) {
                map.customer = colNumber;
                headerNames.customer = val;
            }
        }
    });

    if (showDebug) console.log(`[DEBUG] Sheet "${config.name}" (Header Row: ${headerRowIndex}). Mapped Columns:`, JSON.stringify(headerNames));

    if (showChecker) {
        if (!map.vendor) console.warn(`  [!] WARNING: No Vendor column found for sheet "${sheet.name}". 'Unknown Vendor' checks will be skipped for this sheet.`);
    }

    if (showChecker) {
        console.log(`\nProcessing "${config.shortName}" (${config.name}):`);
        console.log(`  Header Row: ${headerRowIndex}`);

        if (!map.date) console.warn(`  [!] CRITICAL: 'Date' column not found in "${config.shortName}"`);
        if (!map.amount) console.warn(`  [!] CRITICAL: 'Amount' column not found in "${config.shortName}"`);

        console.log(`  Mapping: ${Object.entries(map).filter(([k, v]) => v).map(([k, v]) => k + ':' + v).join(', ')}`);
        if (config.flip) console.log(`  [Polarity] Flip enabled for this sheet.`);
        console.log(`  Linked Account: ${config.linkedAccount || 'NONE (Error if not Ledger)'}`);
    }

    return map;
}

module.exports = { detectSheetColumns };

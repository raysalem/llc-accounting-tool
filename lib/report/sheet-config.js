// Reads the sheet configuration (which tabs to process, their type, polarity,
// header row and linked balance-sheet account) from the Setup sheet.
const { parseBalance, getVal } = require('./cells');

async function readSheetConfigs(ctx) {
    const {
        setupSheet, sheetConfigs, showDebug, uniqueCategories,
    } = ctx;

    // --- Pass 2: Read Sheet Configurations ---
    const sheetInfoTable = setupSheet.getTable('SheetInfo');
    const configRows = [];

    if (sheetInfoTable && sheetInfoTable.table && sheetInfoTable.table.tableRef) {
        if (showDebug) console.log('[DEBUG] Reading SheetInfo from Excel Table...');

        const tableRef = sheetInfoTable.table.tableRef;
        const match = tableRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
            const startRow = parseInt(match[2]);
            const headerRow = setupSheet.getRow(startRow);
            const headerCols = {};
            headerRow.eachCell((cell, colNumber) => {
                const val = (cell.value || '').toString().toLowerCase().trim().replace(/[^a-z]/g, '');
                headerCols[val] = colNumber;
            });

            const colN = headerCols['sheetname'] || headerCols['sheetnameconfig'] || 1;
            const colT = headerCols['sheettype'] || headerCols['type'] || 2;
            const colC = headerCols['linkasset'] || headerCols['linkcategory'] || headerCols['linkcat'] || headerCols['category'] || 0;
            const colO = headerCols['headerrow'] || headerCols['offset'] || 4;
            const colS = headerCols['shortnames'] || headerCols['shortname'] || 5;

            const colStart = headerCols['start'] || headerCols['startbalance'] || 0;
            const colEnd = headerCols['end'] || headerCols['endbalance'] || headerCols['endingbalance'] || 0;

            const endRow = parseInt(match[4]);
            for (let r = startRow + 1; r <= endRow; r++) {
                const row = setupSheet.getRow(r);
                configRows.push({
                    name: getVal(row.getCell(colN)),
                    type: getVal(row.getCell(colT)),
                    cat: colC ? getVal(row.getCell(colC)) : '',
                    // flip is legacy/ignored now, type is king
                    offset: getVal(row.getCell(colO)),
                    shortName: getVal(row.getCell(colS)),
                    startBalance: colStart ? (parseBalance(getVal(row.getCell(colStart))) || 0) : 0,
                    endBalance: colEnd ? parseBalance(getVal(row.getCell(colEnd))) : null
                });
            }
        }
    } else {
        if (showDebug) console.log('[DEBUG] SheetInfo Table missing. Scanning for "Sheet Name" header block...');

        // Scan for header row containing "Sheet Name"
        let blockHeaderRow = -1;
        const blockMap = {};

        setupSheet.eachRow((row, r) => {
            if (blockHeaderRow !== -1) return; // Found already
            let hasName = false;
            row.eachCell((cell, c) => {
                const v = getVal(cell).toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                if (v.includes('sheetname')) hasName = true;
            });

            if (hasName) {
                blockHeaderRow = r;
                // Map this row
                row.eachCell((cell, c) => {
                    const v = getVal(cell).toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                    blockMap[v] = c;
                    // Handle aliases
                    if (v === 'category') blockMap['linkasset'] = c;
                });
            }
        });

        if (blockHeaderRow !== -1) {
            const colN = blockMap['sheetname'] || blockMap['sheetnameconfig'];
            // The user requested "Link Asset", but we support aliases. 
            // Crucially, we will validate that the target IS a Balance Sheet asset.
            const colL = blockMap['linkasset'] || blockMap['linkcategory'] || blockMap['linkcat'] || blockMap['category'];
            const colStart = blockMap['start'] || blockMap['startbalance'];
            const colEnd = blockMap['end'] || blockMap['endbalance'] || blockMap['endingbalance'];
            const colO = blockMap['headerrow'] || blockMap['offset'];
            const colT = blockMap['type'] || blockMap['sheettype']; // Optional now
            const colS = blockMap['shortname'] || blockMap['shortnames'];
            const colFlip = blockMap['flip'] || blockMap['flippolarity'];

            if (showDebug) console.log(`[Setup] Found SheetInfo headers on Row ${blockHeaderRow}. Link Col: ${colL}, Start Col: ${colStart}, End Col: ${colEnd}`);

            setupSheet.eachRow((row, r) => {
                if (r <= blockHeaderRow) return;
                const nameRaw = colN ? getVal(row.getCell(colN)) : null;
                const name = nameRaw ? nameRaw.toString().trim() : '';
                if (name && name.toLowerCase() !== 'sheet name') {
                    configRows.push({
                        name: name,
                        cat: colL ? getVal(row.getCell(colL)) : '',
                        shortName: colS ? getVal(row.getCell(colS)) : '',
                        // Type is now the golden truth for polarity
                        type: colT ? getVal(row.getCell(colT)) : '',
                        offset: colO ? getVal(row.getCell(colO)) : '',
                        startBalance: colStart ? (parseBalance(getVal(row.getCell(colStart))) || 0) : 0,
                        endBalance: colEnd ? parseBalance(getVal(row.getCell(colEnd))) : null,
                        flip: colFlip ? getVal(row.getCell(colFlip)) : ''
                    });
                }
            });
        } else {
            console.warn("[!] WARNING: Could not find 'Sheet Name' header row. Sheet configuration may fail.");
        }
    }

    // Process Valid Config Rows
    if (showDebug) console.log(`[DEBUG] Found ${configRows.length} Config Rows from Setup: ${configRows.map(r => r.name).join(', ')}`);

    for (const conf of configRows) {
        const confSheetName = conf.name;
        if (confSheetName) {
            const cType = (conf.type || '').toString().trim().toLowerCase();
            const confOffset = conf.offset; // Keep confOffset as it's still used

            // Polarity Logic based on STRICT User Type
            let doFlip = false;
            let doLink = true;
            let link = null; // Initialize link here

            if (cType === 'expense' || cType === 'cc' || cType.includes('credit') || cType.includes('liability')) {
                doFlip = true;
            } else if (cType === 'income') {
                doFlip = false;
            } else if (cType === 'ledger') {
                doFlip = false;
                doLink = false; // Ledger never links automatically
            }

            // Explicit Override from 'Flip' column
            if (conf.flip) {
                const fVal = conf.flip.toString().toLowerCase();
                if (fVal === 'yes' || fVal === 'true' || fVal === 'y' || fVal === '1') doFlip = true;
                if (fVal === 'no' || fVal === 'false' || fVal === 'n' || fVal === '0') doFlip = false;
            }

            // Linkage Priority 0: Explicit 'Category' column in SheetInfo
            if (doLink && conf.cat) {
                const explicitCat = conf.cat.toString().trim();
                const lowerEC = explicitCat.toLowerCase();
                // Check if this is a direct match to a display name or key
                if (uniqueCategories.has(lowerEC)) {
                    link = uniqueCategories.get(lowerEC).displayName;
                } else {
                    // Search for exact display name match
                    for (const catData of uniqueCategories.values()) {
                        if (catData.displayName && catData.displayName.toLowerCase() === lowerEC) {
                            link = catData.displayName;
                            break;
                        }
                    }
                }
            }

            if (doLink && !link && cType) {
                const targetType = cType.toLowerCase();
                // Linkage Priority 1: Exact Name Match with Balance Sheet Categories
                for (const catData of uniqueCategories.values()) {
                    if (catData.report !== 'Balance Sheet') continue;
                    if (catData.displayName && catData.displayName.toLowerCase() === targetType) {
                        link = catData.displayName;
                        break;
                    }
                }

                // Linkage Priority 2: Sub-Category or Account Type Match
                if (!link) {
                    for (const catData of uniqueCategories.values()) {
                        if (catData.report !== 'Balance Sheet') continue;

                        // Check 'Sub-Category' (Bank/General)
                        if (catData.subCategory && catData.subCategory.toLowerCase() === targetType) {
                            link = catData.displayName;
                            break;
                        }
                        // Check 'Type' (Asset/Liability)
                        if (catData.accountType && catData.accountType.toLowerCase() === targetType) {
                            link = catData.displayName;
                            break;
                        }
                    }
                }
            }

            // STAGE 2: Validate that 'link' is actually a Balance Sheet account
            if (doLink && link) {
                const lowerLink = link.toLowerCase();
                const catData = uniqueCategories.get(lowerLink);
                if (catData) {
                    if (catData.report !== 'Balance Sheet') {
                        console.warn(`[!] WARNING: Sheet "${confSheetName}" is trying to link to "${link}", but that category is type "${catData.report}". It MUST be a "Balance Sheet" account.`);
                        console.warn(`    Linkage Rejected. Please update your Categories table.`);
                        link = null;
                    }
                } else {
                    // Category not found (yet? or typo)
                    // converting raw name to display name might have failed or it's a raw string
                    // We'll allow it if it looks like a valid string, but warn?
                    // Actually, if it's not in uniqueCategories, we can't be sure type is BS.
                    // But we might be early in the process. uniqueCategories IS populated by now.
                    console.warn(`[!] WARNING: Linked asset "${link}" for sheet "${confSheetName}" not found in Categories table.`);
                }
            }

            if (showDebug) {
                console.log(`[Linkage Result] Sheet "${confSheetName}" (Type: "${cType}") -> Linked to: "${(doLink && link) || 'NONE'}"`);
            }

            if (doLink && !link) {
                console.error(`\n[CRITICAL ERROR] Sheet "${confSheetName}" (Type: "${cType}") failed to link to a Balance Sheet account!`);
                console.error(`  Linkage is REQUIRED for all Bank/Credit Card sheets to track balances.`);
                console.error(`  Reason: No matching "Balance Sheet" category found for Type "${cType}" or explicit Link "${conf.cat || ''}".`);
                const bsExamples = Array.from(uniqueCategories.values())
                    .filter(c => c.report === 'Balance Sheet')
                    .map(c => `"${c.displayName}"`)
                    .slice(0, 5);
                console.error(`  Available 'Balance Sheet' categories: ${bsExamples.join(', ')}...`);
                console.error(`  Fix: update 'Type' to match a BS Asset/Liability, or use 'Link Asset' column in SheetInfo.`);
                ctx.hasErrors = true; // Mark global error
            }

            sheetConfigs.push({
                name: confSheetName.toString().trim(),
                shortName: (conf.shortName && conf.shortName.toString().trim()) || confSheetName.toString().trim(),
                type: cType, // 'income', 'expense', 'ledger'
                flip: doFlip,
                offset: parseInt(confOffset) || 0,
                startBalance: parseFloat(conf.startBalance) || 0,
                endBalance: conf.endBalance !== null ? parseFloat(conf.endBalance) : null,
                linkedAccount: doLink ? link : null
            });
        }
    }

    if (showDebug) {
        console.log(`\n--- CONSUMED SHEETINFO TABLE ---`);
        const header = `Sheet Name`.padEnd(30) + `Type`.padEnd(10) + `Linked Account`.padEnd(30) + `Flip`.padEnd(6) + `Offset`.padEnd(8) + `Start`.padStart(12) + `End`.padStart(12);
        console.log(header);
        console.log("-".repeat(header.length + 5));
        sheetConfigs.forEach(s => {
            const endDisp = (s.endBalance !== null) ? s.endBalance.toFixed(2) : "N/A";
            console.log(`${s.name.padEnd(30)}${s.type.padEnd(10)}${(s.linkedAccount || 'NONE').padEnd(30)}${(s.flip ? 'YES' : 'NO').padEnd(6)}${s.offset.toString().padEnd(8)}${(s.startBalance ? s.startBalance.toFixed(2) : '0.00').padStart(12)}${endDisp.padStart(12)}`);
        });
        console.log("");
    }

    if (sheetConfigs.length === 0) {
        // Fallback defaults if no config found in Setup
        console.warn('[!] No sheet configurations found in Setup. Using defaults.');
        sheetConfigs.push({ name: 'Bank Transactions', shortName: 'Bank Transactions', type: 'Bank', flip: false, offset: 1, linkedAccount: null });
        sheetConfigs.push({ name: 'Credit Card Transactions', shortName: 'Credit Card Transactions', type: 'CC', flip: true, offset: 1, linkedAccount: null });
    }
}

module.exports = { readSheetConfigs };

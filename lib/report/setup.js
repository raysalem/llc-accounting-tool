// Reads the Setup sheet: categories, vendors, customers, 1099 settings and payer info,
// and warns about names shared between categories, vendors and customers.
const { NORM_VEND, getVal } = require('./cells');

async function readSetup(ctx) {
    const {
        duplicateCategories, payerInfo, setupSheet, showDebug, uniqueCategories, validCategories,
        validCustomers, validSubCategories, validVendors, vendor1099Map, vendorDetailsMap,
    } = ctx;

    // --- 1. Validate Required Named Tables in Setup Sheet ---
    const REQUIRED_TABLES = ['CompanyInfo', 'Categories', 'Vendor', 'Customer', 'SheetInfo'];
    const foundTables = [];
    const missingTables = [];



    // Check which tables exist
    REQUIRED_TABLES.forEach(tableName => {
        try {
            const table = setupSheet.getTable(tableName);
            if (table && table.table) {
                foundTables.push(tableName);
            } else {
                missingTables.push(tableName);
            }
        } catch (e) {
            // Table doesn't exist
            missingTables.push(tableName);
        }
    });

    // ERROR if any tables are missing
    // WARN if tables are missing (but allow fallback to row-scanning)
    if (missingTables.length > 0) {
        console.warn(`[WARNING] Setup sheet is missing formal Excel table(s): ${missingTables.join(', ')}`);
        console.warn(`Falling back to row-scanning for Setup configuration.`);
        // Do not exit, as we have fallback logic below
    }

    if (showDebug) {
        console.log(`[Setup] Found all required tables: ${foundTables.join(', ')}`);
    }

    // TODO: Replace this with table-based reading once tables are created
    // Temporary fallback: Keep old header-map logic for backward compatibility
    function getHeaderMap(sheet) {
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
    const { map: setupHeaders, headerRow: setupHeaderRow } = getHeaderMap(setupSheet);
    const getCol = (key, preferred = 'first') => {
        const indices = setupHeaders.get(key);
        if (!indices || indices.length === 0) return null;
        return preferred === 'first' ? indices[0] : indices[indices.length - 1];
    };

    // Table 1: Category Info
    const colCategory = getCol('category');
    const colSubCategory = getCol('subcategory');
    const colType = getCol('accounttype') || getCol('type');
    const colReport = getCol('report') || getCol('pnlbs') || getCol('statement') || getCol('bspl');
    const colTransferAccount = getCol('transferaccount') || getCol('transfer');

    // Table 2: Vendors
    const colVendor = getCol('vendors') || getCol('vendor');

    // Additional Vendor Columns for 1099
    const colBusiness = getCol('businessname') || getCol('business');
    const colName = getCol('name') || getCol('fullname');
    const colSSN = getCol('ssn') || getCol('ein') || getCol('taxid') || getCol('tin') || getCol('ssnein');
    const colAddress = getCol('address');
    const colEmail = getCol('email');
    const colPhone = getCol('phone');

    const finalCol1099 = getCol('1099') || getCol('1099nec') || getCol('nec');
    // New Split Columns: "1099 Type" -> "1099type", "1099 Required" -> "1099required"
    const col1099Type = getCol('1099type');
    const col1099Req = getCol('1099required');
    if (showDebug) console.log(`[DEBUG] 1099 Column Detection: Type=${col1099Type}, Req=${col1099Req}, Legacy=${finalCol1099}`);

    // Table 3: Customers
    const colCustomer = getCol('customers') || getCol('customer');

    // Table 4: Sheet Info - CRITICAL: We must ensure we don't accidentally pick columns from Table 1.
    // Since SheetInfo is typically to the right or below, we prefer the 'last' occurrences of these names.
    const colSheetName = getCol('sheetnameconfig') || getCol('sheetname', 'last');
    const colSheetType = getCol('sheettype') || getCol('type', 'last');

    // Table 5: Payer Info (Vertical Key-Value)
    const colPayerKey = getCol('companyinfo', 'last') || getCol('payerinfo', 'last') || getCol('companyname', 'last');
    const colPayerValue = colPayerKey ? colPayerKey + 1 : null;
    // If strict "Key" / "Value" headers exist, use them? No, user said "two columns". We assume Key Col -> Value Col.

    if (showDebug) {
        console.log(`\n--- SETUP HEADER DETECTION ---`);
        console.log(`Detected header row: ${setupHeaderRow}`);
        console.log(`Col Category: ${colCategory || 'NOT FOUND'}`);
        console.log(`Col Sheet Name: ${colSheetName || 'NOT FOUND'}`);
        console.log(`Col Sheet Type: ${colSheetType || 'NOT FOUND'}`);
        console.log(`Col Report: ${colReport || 'NOT FOUND'}`);
    }

    // --- Pass 1: Read Reference Tables (Categories, Vendors, Customers) ---
    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber <= setupHeaderRow) return;

        // 1. Process Category Table
        const catName = colCategory ? getVal(row.getCell(colCategory)) : null;
        if (catName) {
            const trimmed = catName.toString().trim();
            const lower = trimmed.toLowerCase();
            const typeVal = colType ? getVal(row.getCell(colType)) : '';
            const subCatVal = colSubCategory ? getVal(row.getCell(colSubCategory)) : '';
            const report = colReport ? getVal(row.getCell(colReport)) : '';
            const transferAccountVal = colTransferAccount ? getVal(row.getCell(colTransferAccount)).toString().trim() : '';
            validCategories.add(lower);

            // Track valid subcategories
            if (subCatVal) {
                const subCatStr = subCatVal.toString().trim();
                if (subCatStr) {
                    validSubCategories.add(subCatStr.toLowerCase());
                }
            }

            // Detect CONFLICTING category definitions (same category, different Report type)
            if (uniqueCategories.has(lower)) {
                const existing = uniqueCategories.get(lower);
                // Only warn if Report type conflicts (different subcategories with same category is VALID)
                if (existing.report !== report) {
                    duplicateCategories.push({
                        name: trimmed,
                        row: rowNumber,
                        newReport: report,
                        existingReport: existing.report
                    });
                }
            }

            let rType = 'P&L';
            const rUpper = report.toString().trim().toUpperCase();
            const tLower = typeVal.toString().trim().toLowerCase();

            if (rUpper.includes('BALANCE') || rUpper.includes('BS')) {
                rType = 'Balance Sheet';
            } else if (rUpper.includes('P&L') || rUpper.includes('PROFIT')) {
                rType = 'P&L';
            } else if (rUpper.includes('TRANSFER') || rUpper.includes('IGNORE')) {
                rType = 'Transfer';
            } else if (tLower.includes('transfer')) {
                rType = 'Transfer';
            } else if (tLower.includes('asset') || tLower.includes('liability') || tLower.includes('bank') || tLower.includes('credit') || tLower.includes('cc')) {
                // Fallback for missing/ambiguous Report column
                rType = 'Balance Sheet';
            }
            // STRICT VALIDATION: Check for Report/Type mismatch
            const isPL = rType === 'P&L';
            const isBS = rType === 'Balance Sheet';
            const tSanitized = tLower.replace(/[^a-z]/g, '');

            if (isPL) {
                if (!tSanitized.includes('income') && !tSanitized.includes('expense')) {
                    console.warn(`[!] CRITICAL CONFIG ERROR (Row ${rowNumber}): Category "${trimmed}" is set to "P&L" but Type is "${typeVal}". P&L requires Income or Expense.`);
                }
            } else if (isBS) {
                if (!tSanitized.includes('asset') && !tSanitized.includes('liability') && !tSanitized.includes('equity')) {
                    console.warn(`[!] CRITICAL CONFIG ERROR (Row ${rowNumber}): Category "${trimmed}" is set to "Balance Sheet" but Type is "${typeVal}". BS requires Asset, Liability, or Equity.`);
                }
            }

            uniqueCategories.set(lower, {
                report: rType,
                accountType: typeVal,
                subCategory: subCatVal,
                displayName: trimmed,
                transferAccount: transferAccountVal
            });
        }

        // 2. Process Vendor Table
        const vendor = colVendor ? getVal(row.getCell(colVendor)) : null;
        if (vendor) {
            const vRaw = vendor.toString().trim();
            const lowerV = NORM_VEND(vRaw);
            validVendors.set(lowerV, vRaw);
            if (showDebug && validVendors.size % 50 === 0) console.log(`... Loaded ${validVendors.size} vendors so far ...`);


            // 1099 Logic: Prioritize "Type" + "Required", Fallback to old "1099" column
            let type = '';
            let req = '';

            if (col1099Type) type = getVal(row.getCell(col1099Type)).toString().trim().toUpperCase();
            if (col1099Req) req = getVal(row.getCell(col1099Req)).toString().trim().toUpperCase();

            if (type || req) {
                // console.log(`[DEBUG] Vendor ${lowerV}: Type='${type}', Req='${req}'`);
            }

            // Legacy/Combined Column logic fallback
            if (!type && !req && finalCol1099) {
                const unknownVal = getVal(row.getCell(finalCol1099)).toString().trim().toUpperCase();
                if (unknownVal === 'NEC' || unknownVal === 'INT') {
                    type = unknownVal;
                    req = 'YES'; // Legacy column with type implies required
                } else if (unknownVal === 'YES' || unknownVal === 'Y') {
                    type = 'NEC'; // Default to NEC
                    req = 'YES';
                }
            }

            // Determine final status
            // If "Required" is NO, ignored.
            // If "Required" is YES (or blank/implied by Type presence), use Type.
            const isExplicitNo = (req === 'NO' || req === 'N' || req === 'FALSE');

            if (type && !isExplicitNo) {
                if (type === 'NEC' || type === 'INT' || type === 'MISC') {
                    vendor1099Map.set(lowerV, { type, req });
                    if (showDebug) console.log(`  > 1099 Recognized: ${vRaw} (${type})`);
                }
            } else if (!type && (req === 'YES' || req === 'Y') && !isExplicitNo) {
                // Required but no type? Default NEC
                vendor1099Map.set(lowerV, { type: 'NEC', req });
                if (showDebug) console.log(`  > 1099 Recognized: ${vRaw} (NEC - Default)`);
            }

            // Capture Details
            vendorDetailsMap.set(lowerV, {
                business: colBusiness ? getVal(row.getCell(colBusiness)) : '',
                name: colName ? getVal(row.getCell(colName)) : vRaw, // Fallback to raw vendor name
                ssn: colSSN ? getVal(row.getCell(colSSN)) : '',
                address: colAddress ? getVal(row.getCell(colAddress)) : '',
                email: colEmail ? getVal(row.getCell(colEmail)) : '',
                phone: colPhone ? getVal(row.getCell(colPhone)) : ''
            });

        }

        // 3. Process Customer Table
        const customer = colCustomer ? getVal(row.getCell(colCustomer)) : null;
        if (customer) {
            const cRaw = customer.toString().trim();
            validCustomers.set(NORM_VEND(cRaw), cRaw);
        }

        // 5. Process Payer Info (Vertical Table)
        if (colPayerKey && colPayerValue) {
            const pKey = getVal(row.getCell(colPayerKey));
            const pVal = getVal(row.getCell(colPayerValue));
            if (pKey) {
                const kStr = pKey.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                payerInfo[kStr] = pVal.toString().trim();
                // console.log(`[DEBUG] Payer Info: '${kStr}' -> '${payerInfo[kStr]}'`);
            }
        }
    });
    if (showDebug) console.log(`[Setup] Loaded ${validVendors.size} Valid Vendors, ${validCustomers.size} Customers, ${validCategories.size} Categories.`);



    // --- Check for Customer/Vendor Overlap ---
    const overlapSet = new Set();
    for (const [vLower, vName] of validVendors) {
        if (validCustomers.has(vLower)) {
            overlapSet.add(vName); // or validCustomers.get(vLower)
        }
    }

    if (overlapSet.size > 0) {
        console.warn(`\n[!] CRITICAL ACCOUNTING WARNING: Overlap detected between Vendors and Customers.`);
        console.warn(`    The following names appear in BOTH Vendor and Customer tables:`);
        console.warn(`    ${Array.from(overlapSet).join(', ')}`);
        console.warn(`    Principle: An entity should be EITHER a Vendor (Expense side) OR a Customer (Income side), never both.`);
        console.warn(`    Fix: Rename one (e.g., "Client X (Vendor)" vs "Client X").\n`);
    }

    // --- Check for Sub-Category / Customer / Vendor Overlap ---
    // User Request: "ensure category, sub-catgory and vendor customer names, never are onlky used for the correct column"
    const subCatOverlaps = { cust: [], vend: [] };

    // Check SubCategories vs Customers
    validSubCategories.forEach(sub => {
        if (validCustomers.has(sub)) subCatOverlaps.cust.push(validCustomers.get(sub));
        if (validVendors.has(sub)) subCatOverlaps.vend.push(validVendors.get(sub));
    });

    if (subCatOverlaps.cust.length > 0) {
        console.warn(`\n[!] AMBIGUITY WARNING: Sub-Category vs Customer Overlap.`);
        console.warn(`    The following names are defined as BOTH a Sub-Category AND a Customer:`);
        console.warn(`    ${subCatOverlaps.cust.join(', ')}`);
        console.warn(`    Risk: If "HOLDCO" is a Customer, it should generally NOT be a Sub-Category.`);
        console.warn(`    This causes confusion in reports. Please rename one (e.g., "HOLDCO Sub" vs "HOLDCO Cust").\n`);
    }

    if (subCatOverlaps.vend.length > 0) {
        console.warn(`\n[!] AMBIGUITY WARNING: Sub-Category vs Vendor Overlap.`);
        console.warn(`    The following names are defined as BOTH a Sub-Category AND a Vendor:`);
        console.warn(`    ${subCatOverlaps.vend.join(', ')}`);
        console.warn(`    Fix: Rename the Sub-Category or the Vendor to be distinct.\n`);
    }

    // Check Categories vs Customers/Vendors (Less common, but possible)
    const catOverlaps = { cust: [], vend: [] };
    validCategories.forEach(cat => {
        if (validCustomers.has(cat)) catOverlaps.cust.push(validCustomers.get(cat));
        if (validVendors.has(cat)) catOverlaps.vend.push(validVendors.get(cat));
    });

    if (catOverlaps.cust.length > 0 || catOverlaps.vend.length > 0) {
        console.warn(`\n[!] AMBIGUITY WARNING: Category Name Overlaps.`);
        if (catOverlaps.cust.length) console.warn(`    Category == Customer: ${catOverlaps.cust.join(', ')}`);
        if (catOverlaps.vend.length) console.warn(`    Category == Vendor:   ${catOverlaps.vend.join(', ')}`);
        console.warn(`    Ideally, structural Categories (e.g. Sales, Rent) should not share names with entities.\n`);
    }

    Object.assign(ctx, { getHeaderMap });
}

module.exports = { readSetup };

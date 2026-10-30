// Reads the Setup sheet: categories, vendors, customers, 1099 settings and payer info,
// and warns about names shared between categories, vendors and customers.
const { NORM_VEND, getVal } = require('./cells');
const { getHeaderMap } = require('./header-map');
const { warnNameOverlaps } = require('./name-overlaps');

async function readSetup(ctx) {
    const {
        duplicateCategories, payerInfo, setupSheet, showDebug, uniqueCategories, validCategories,
        validCustomers, validSubCategories, validVendors, vendor1099Map, vendorDetailsMap,
    } = ctx;

    // --- Check for the formal Excel tables in the Setup sheet ---
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

    const { map: setupHeaders, headerRow: setupHeaderRow } = getHeaderMap(setupSheet, showDebug);
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



    warnNameOverlaps(ctx);

    ctx.getHeaderMap = sheet => getHeaderMap(sheet, showDebug);
}

module.exports = { readSetup };

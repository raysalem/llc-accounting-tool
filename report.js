const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function resolveShortcut(filePath) {
    try {
        const ext = path.extname(filePath).toLowerCase();
        if (ext === '.lnk') {
            const escapedPath = filePath.replace(/'/g, "''");
            const command = `powershell -NoProfile -Command "(New-Object -ComObject WScript.Shell).CreateShortcut('${escapedPath}').TargetPath"`;
            const target = execSync(command).toString().trim();
            if (target) return target;
        } else if (ext === '.url') {
            const content = fs.readFileSync(filePath, 'utf8');
            const match = content.match(/^URL=(.*)$/m);
            if (match && match[1]) {
                let target = match[1].trim();
                if (target.startsWith('file:///')) target = target.replace('file:///', '');
                else if (target.startsWith('file://')) target = target.replace('file://', '');
                return decodeURIComponent(target);
            }
        }
    } catch (e) {
        console.error(`Warning: Failed to resolve shortcut '${filePath}': ${e.message}`);
    }
    return filePath;
}

function isTruthy(val) {
    if (val === null || val === undefined) return false;
    if (typeof val === 'boolean') return val;
    const s = val.toString().trim().toLowerCase();
    return s === 'yes' || s === 'true' || s === 'y' || s === 'x' || s === '1';
}

async function updateFinancials() {
    const args = process.argv.slice(2);
    const saveFlag = args.includes('--save');
    const showPL = args.includes('--pl');
    const showBS = args.includes('--bs');
    const showVendor = args.includes('--vendor');
    const showVendorSub = args.includes('--vendor-sub');
    const showCustomer = args.includes('--customer');
    const showCustomerSub = args.includes('--customer-sub');
    const showPLSub = args.includes('--pl-sub');
    const showBSSub = args.includes('--bs-sub');
    const showChecker = args.includes('--checker');
    const showDebug = args.includes('--debug');
    const show1099 = args.includes('--1099')

    // Parse --details <Category>
    const detailsIndex = args.indexOf('--details');
    const targetDetailsCategory = detailsIndex !== -1 && args[detailsIndex + 1] ? args[detailsIndex + 1].toLowerCase().trim() : null;
    const showDetails = !!targetDetailsCategory;

    // Help Menu
    if (args.includes('--help')) {
        console.log(`
Usage: node report.js [filename] [flags]

Description:
  Updates the financial accounting spreadsheet. It reads the Setup, Ledger, and Transaction sheets,
  categorizes transactions, balances the ledger, and generates P&L / Balance Sheet reports in standard Output format.
  Supports .lnk and .url shortcut files as input.

Arguments:
  [filename]      Path to the Excel file or shortcut (default: LLC_Accounting_Template.xlsx)

Flags:
  --help          Show this help message.
  --save          Save changes to the Excel file (Summary tab and formatting).
                  (Default behavior is print-only, which does not modify the file).
  --pl            Print the Profit & Loss statement to the console.
  --bs            Print the Balance Sheet to the console.
  --checker       Run the Data Integrity Checker and verify row-by-row categorization issues.
  --debug         Enable verbose debug output for troubleshooting.
  --pl-sub        (Optional) Print detailed P&L with sub-category breakdowns.
  --bs-sub        (Optional) Print detailed Balance Sheet with sub-category breakdowns.
  --vendor        (Optional) Print spending statistics by Vendor.
  --vendor-sub    (Optional) Print detailed Vendor Spending with sheet-level breakdowns.
  --customer      (Optional) Print income statistics by Customer.
  --customer-sub  (Optional) Print detailed Customer Income with sheet-level breakdowns.
  --1099          (Optional) Generate 1099-NEC and 1099-INT reports for enabled vendors.
  --details "Cat" (Optional) List all transactions for a specific Category (e.g., --details "Office Supplies").

Example:
  node report.js "My_Books_2025.xlsx" --pl --checker --save
        `);
        return;
    }

    const knownFlags = [
        '--save', '--pl', '--bs', '--vendor', '--vendor-sub', '--customer', '--customer-sub', '--pl-sub', '--bs-sub', '--checker', '--debug', '--details', '--help', '--1099', '--1099-nec'
    ];

    // Check for unknown arguments
    const unknownArgs = args.filter(a => a.startsWith('--') && !knownFlags.includes(a));
    if (unknownArgs.length > 0) {
        console.error(`Error: Unknown argument(s): ${unknownArgs.join(', ')}`);
        console.error('Run with --help to see available options.');
        process.exit(1);
    }

    const specificFilter = showPL || showBS || showVendor || showCustomer || showCustomerSub || showPLSub || showBSSub || showChecker || showDetails || show1099;
    const showAll = !specificFilter; // Default to showing standard report if no specific filter is set

    let filename = args.find(a => !a.startsWith('--')) || 'LLC_Accounting_Template.xlsx';

    // Resolve shortcut if needed
    if (fs.existsSync(filename)) {
        const resolved = resolveShortcut(filename);
        if (resolved !== filename) {
            console.log(`Resolved shortcut '${filename}' -> '${resolved}'`);
            filename = resolved;
        }
    }

    if (!fs.existsSync(filename)) {
        console.error(`Error: File '${filename}' not found.`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    try {
        if (showChecker || saveFlag) console.log(`Loading workbook: ${filename}...`);
        await workbook.xlsx.readFile(filename);
    } catch (e) {
        console.error(`Error reading file: ${e.message}`);
        return;
    }

    const setupSheet = workbook.getWorksheet('Setup');
    const ledgerSheet = workbook.getWorksheet('Ledger');
    let summarySheet = workbook.getWorksheet('Summary');

    if (!setupSheet || !ledgerSheet) {
        console.error('Error: Mandatory sheets (Setup or Ledger) missing.');
        return;
    }

    // --- State ---
    const validCategories = new Set(); // Stores lowercase for validation
    const validVendors = new Map();    // Maps lower -> Display Name
    const vendor1099Map = new Map();   // Maps lower -> { type: 'NEC'|'INT', req: 'YES'|'NO'|'' }
    const vendorDetailsMap = new Map(); // Maps lower -> strict Object of details
    let payerInfo = {}; // Payer/Company Info Map
    const validCustomers = new Map();  // Maps lower -> Display Name
    const uniqueCategories = new Map(); // Maps lower -> { report, accountType, displayName }
    const validSubCategories = new Set(); // Set of all valid subcategories from Setup
    const sheetConfigs = [];

    const catStats = {};
    const vendorStats = {};
    const vendor1099Stats = { NEC: {}, INT: {} };
    const customerStats = {};
    let bankTotal = 0;
    let ccTotal = 0;
    let uncategorizedBank = 0;
    let uncategorizedCC = 0;

    const illegalCategories = [];
    const illegalVendors = [];
    const illegalCustomers = [];
    const illegalSubCategories = [];
    const uncategorizedDetails = [];
    const detailsRows = [];
    const offsetWarnings = [];
    const duplicateCategories = []; // Track duplicate category definitions

    // Track which report types (P&L vs BS) vendors/customers are used in
    // Structure: Map<vendor, Map<reportType, Array<{sheet, row, category, date}>>>
    const vendorReportUsage = new Map();
    const customerReportUsage = new Map();

    // --- Helper ---
    function getVal(cell) {
        if (!cell) return '';
        let v = cell.value;
        if (v && typeof v === 'object' && v.result !== undefined) v = v.result;
        if (v && v.richText) return v.richText.map(t => t.text).join('').trim();
        if (typeof v === 'number') return v;
        if (v instanceof Date) return v;
        return (v === null || v === undefined) ? '' : v.toString().trim();
    }

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

    if (showChecker) {
        console.log(`[Setup] Found all required tables: ${foundTables.join(', ')}`);
    }

    // TODO: Replace this with table-based reading once tables are created
    // Temporary fallback: Keep old header-map logic for backward compatibility
    function getHeaderMap(sheet) {
        let bestRowIdx = 1;
        let bestMap = new Map();
        let maxFound = -1;
        const lookups = ['category', 'subcategory', 'vendors', 'vendor', 'sheetname', 'sheetnameconfig', 'report', 'linkcategory', 'linkcat'];

        for (let ri = 1; ri <= 10; ri++) {
            const currentMap = new Map();
            const row = sheet.getRow(ri);
            let foundCount = 0;
            row.eachCell((cell, colNumber) => {
                const val = getVal(cell).toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                if (val) {
                    if (!currentMap.has(val)) currentMap.set(val, []);
                    currentMap.get(val).push(colNumber);
                    if (lookups.includes(val)) foundCount++;
                }
            });
            if (foundCount > maxFound) {
                maxFound = foundCount;
                bestRowIdx = ri;
                bestMap = currentMap;
            }
            if (foundCount >= 3 || (foundCount > 0 && ri > 5)) break;
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

    // Table 2: Vendors
    const colVendor = getCol('vendors') || getCol('vendor');

    // Additional Vendor Columns for 1099
    const colBusiness = getCol('businessname') || getCol('business');
    const colName = getCol('name') || getCol('fullname');
    const colSSN = getCol('ssn') || getCol('ein') || getCol('taxid') || getCol('tin') || getCol('ssnein');
    const colAddress = getCol('address');
    const colEmail = getCol('email');
    const colPhone = getCol('phone');

    const finalCol1099 = getCol('1099') || getCol('1099nec');
    // New Split Columns: "1099 Type" -> "1099type", "1099 Required" -> "1099required"
    const col1099Type = setupHeaders.get('1099type');
    const col1099Req = setupHeaders.get('1099required');
    // console.log(`[DEBUG] 1099 Column Detection: Type=${col1099Type}, Req=${col1099Req}, Legacy=${finalCol1099}`);

    // Table 3: Customers
    const colCustomer = getCol('customers') || getCol('customer');

    // Table 4: Sheet Info - CRITICAL: We must ensure we don't accidentally pick columns from Table 1.
    // Since SheetInfo is typically to the right or below, we prefer the 'last' occurrences of these names.
    const colSheetName = getCol('sheetnameconfig') || getCol('sheetname', 'last');
    const colSheetType = getCol('sheettype') || getCol('type', 'last');
    // The user renamed 'Category' to 'Link Category' for the asset account linkage
    const colSheetCat = getCol('linkcategory') || getCol('linkcat') || getCol('category', 'last');
    const colShortName = getCol('shortname', 'last') || getCol('short', 'last');
    const colFlip = getCol('flippolarityyesno', 'last') || getCol('flippolarity', 'last') || getCol('flip', 'last');
    const colOffset = getCol('headerrow', 'last') || getCol('offset', 'last');

    // Table 5: Payer Info (Vertical Key-Value)
    let colPayerKey = getCol('companyinfo', 'last') || getCol('payerinfo', 'last') || getCol('companyname', 'last');
    let colPayerValue = colPayerKey ? colPayerKey + 1 : null;
    // If strict "Key" / "Value" headers exist, use them? No, user said "two columns". We assume Key Col -> Value Col.

    if (showChecker) {
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
            } else if (tLower.includes('asset') || tLower.includes('liability') || tLower.includes('bank') || tLower.includes('credit') || tLower.includes('cc')) {
                // Fallback for missing/ambiguous Report column
                rType = 'Balance Sheet';
            }

            uniqueCategories.set(lower, {
                report: rType,
                accountType: typeVal,
                subCategory: subCatVal,
                displayName: trimmed
            });
        }

        // 2. Process Vendor Table
        const vendor = colVendor ? getVal(row.getCell(colVendor)) : null;
        if (vendor) {
            const vRaw = vendor.toString().trim();
            const lowerV = vRaw.toLowerCase();
            validVendors.set(lowerV, vRaw);
            if (showChecker && validVendors.size % 50 === 0) console.log(`... Loaded ${validVendors.size} vendors so far ...`);


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
                if (type === 'NEC' || type === 'INT') {
                    vendor1099Map.set(lowerV, { type, req });
                    // if (showChecker) console.log(`  > 1099 Detected: ${lowerV} (${type})`);
                }
            } else if (!type && (req === 'YES' || req === 'Y') && !isExplicitNo) {
                // Required but no type? Default NEC
                vendor1099Map.set(lowerV, { type: 'NEC', req });
                if (showChecker) console.log(`  > 1099 Detected: ${lowerV} (NEC - Default)`);
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
            validCustomers.set(cRaw.toLowerCase(), cRaw);
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
    if (showChecker) console.log(`[Setup] Loaded ${validVendors.size} Valid Vendors, ${validCustomers.size} Customers, ${validCategories.size} Categories.`);



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
        console.warn(`    Risk: If "QOZB" is a Customer, it should generally NOT be a Sub-Category.`);
        console.warn(`    This causes confusion in reports. Please rename one (e.g., "QOZB Sub" vs "QOZB Cust").\n`);
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

    // --- Pass 2: Read Sheet Configurations ---
    const sheetInfoTable = setupSheet.getTable('SheetInfo');
    let configRows = [];

    if (sheetInfoTable && sheetInfoTable.table && sheetInfoTable.table.tableRef) {
        if (showChecker) console.log('[DEBUG] Reading SheetInfo from Excel Table...');

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
            const colC = headerCols['linkcategory'] || headerCols['linkcat'] || headerCols['category'] || 0;
            const colF = headerCols['flippolarityyesno'] || headerCols['flippolarity'] || headerCols['flip'] || 3;
            const colO = headerCols['headerrow'] || headerCols['offset'] || 4;
            const colS = headerCols['shortnames'] || headerCols['shortname'] || 5;

            const endRow = parseInt(match[4]);
            for (let r = startRow + 1; r <= endRow; r++) {
                const row = setupSheet.getRow(r);
                configRows.push({
                    name: getVal(row.getCell(colN)),
                    type: getVal(row.getCell(colT)),
                    cat: colC ? getVal(row.getCell(colC)) : '',
                    flip: getVal(row.getCell(colF)),
                    offset: getVal(row.getCell(colO)),
                    shortName: getVal(row.getCell(colS))
                });
            }
        }
    } else {
        if (showChecker) console.log('[DEBUG] SheetInfo Table missing. Scanning rows using detected headers...');
        // Fallback: Scan entire sheet using found column indices
        if (colSheetName) {
            setupSheet.eachRow((row, r) => {
                if (r === 1) return;
                const name = getVal(row.getCell(colSheetName));
                if (name && name.toString().trim()) {
                    configRows.push({
                        name: name,
                        type: colSheetType ? getVal(row.getCell(colSheetType)) : '',
                        cat: colSheetCat ? getVal(row.getCell(colSheetCat)) : '',
                        flip: colFlip ? getVal(row.getCell(colFlip)) : '',
                        offset: colOffset ? getVal(row.getCell(colOffset)) : '',
                        shortName: colShortName ? getVal(row.getCell(colShortName)) : ''
                    });
                }
            });
        }
    }

    // Process Valid Config Rows
    if (showDebug) console.log(`[DEBUG] Found ${configRows.length} Config Rows from Setup: ${configRows.map(r => r.name).join(', ')}`);

    for (const conf of configRows) {
        const confSheetName = conf.name;
        if (confSheetName) {
            const cType = conf.type ? conf.type.toString().trim() : '';
            const confFlip = conf.flip;
            const confOffset = conf.offset;

            let link = null;

            // Linkage Priority 0: Explicit 'Category' column in SheetInfo
            if (conf.cat) {
                const explicitCat = conf.cat.toString().trim();
                const lowerEC = explicitCat.toLowerCase();
                // Check if this is a direct match to a display name or key
                if (uniqueCategories.has(lowerEC)) {
                    link = uniqueCategories.get(lowerEC).displayName;
                } else {
                    // Search for exact display name match
                    for (const [catRaw, catData] of uniqueCategories.entries()) {
                        if (catData.displayName && catData.displayName.toLowerCase() === lowerEC) {
                            link = catData.displayName;
                            break;
                        }
                    }
                }
            }

            if (!link && cType) {
                const targetType = cType.toLowerCase();
                // Linkage Priority 1: Exact Name Match with Balance Sheet Categories
                for (const [catRaw, catData] of uniqueCategories.entries()) {
                    if (catData.report !== 'Balance Sheet') continue;
                    if (catData.displayName && catData.displayName.toLowerCase() === targetType) {
                        link = catData.displayName;
                        break;
                    }
                }

                // Linkage Priority 2: Sub-Category or Account Type Match
                if (!link) {
                    for (const [catRaw, catData] of uniqueCategories.entries()) {
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

            if (showChecker) {
                console.log(`[Linkage Result] Sheet "${confSheetName}" (Type: "${cType}") -> Linked to: "${link || 'NONE'}"`);
                if (!link) {
                    const bsExamples = Array.from(uniqueCategories.values())
                        .filter(c => c.report === 'Balance Sheet')
                        .map(c => `"${c.displayName}" (${c.accountType || 'no type'})`)
                        .slice(0, 5);
                    console.log(`  > Link Failed: No category found with Report="Balance Sheet" matching "${targetType}".`);
                    if (bsExamples.length > 0) {
                        console.log(`  > Available BS accounts: ${bsExamples.join(', ')}${bsExamples.length === 5 ? '...' : ''}`);
                    } else {
                        console.log(`  > WARNING: No Balance Sheet categories loaded! Check your 'Report' column.`);
                    }
                }
            }

            sheetConfigs.push({
                name: confSheetName.toString().trim(),
                shortName: (conf.shortName && conf.shortName.toString().trim()) || confSheetName.toString().trim(),
                type: cType,
                flip: isTruthy(confFlip),
                offset: parseInt(confOffset) || 0,
                linkedAccount: link
            });
        }
    }

    if (showChecker) {
        console.log(`\n--- CONSUMED SHEETINFO TABLE ---`);
        const header = `Sheet Name`.padEnd(30) + `Type`.padEnd(15) + `Linked Account`.padEnd(30) + `Flip`.padEnd(10) + `Offset`;
        console.log(header);
        console.log("-".repeat(header.length));
        sheetConfigs.forEach(s => {
            console.log(`${s.name.padEnd(30)}${s.type.padEnd(15)}${(s.linkedAccount || 'NONE').padEnd(30)}${(s.flip ? 'YES' : 'NO').padEnd(10)}${s.offset}`);
        });
        console.log("");
    }

    if (sheetConfigs.length === 0) {
        // Fallback defaults if no config found in Setup
        console.warn('[!] No sheet configurations found in Setup. Using defaults.');
        sheetConfigs.push({ name: 'Bank Transactions', type: 'Bank', flip: false, offset: 1, linkedAccount: null });
        sheetConfigs.push({ name: 'Credit Card Transactions', type: 'CC', flip: true, offset: 1, linkedAccount: null });
    }

    // --- 1.5 Load External Vendors (Override/Enrich Setup) ---
    async function loadExternalVendors() {
        const dir = path.dirname(filename);
        const candidates = ['vendor.xlsx', 'vendor.csv'];

        for (const fName of candidates) {
            const fPath = path.join(dir, fName);
            if (fs.existsSync(fPath)) {
                if (showChecker) console.log(`[Validation] Found External Vendor File: ${fPath}`);

                const vWb = new ExcelJS.Workbook();
                if (fName.endsWith('.csv')) await vWb.csv.readFile(fPath);
                else await vWb.xlsx.readFile(fPath);

                const vSheet = vWb.worksheets[0];
                if (!vSheet) continue;

                // Simple Header Search
                const vHeaders = getHeaderMap(vSheet, 1);
                // Map common headers
                const cName = vHeaders.get('name') || vHeaders.get('vendor') || vHeaders.get('full name');
                const cBiz = vHeaders.get('business name') || vHeaders.get('business');
                const cSSN = vHeaders.get('ssn') || vHeaders.get('tax id') || vHeaders.get('tin');
                const cAddr = vHeaders.get('address');
                const cEmail = vHeaders.get('email');
                const cPhone = vHeaders.get('phone');

                vSheet.eachRow((row, r) => {
                    if (r === 1) return;
                    // Primary Key is Name or Business
                    let name = cName ? getVal(row.getCell(cName)) : '';
                    let biz = cBiz ? getVal(row.getCell(cBiz)) : '';
                    const key = (name || biz).toLowerCase().trim();
                    if (!key) return;

                    const existing = vendorDetailsMap.get(key) || {};

                    // Merge: External takes precedence for empty fields, or overwrite?
                    // "use the info in the report" -> implies external is source of truth for details
                    vendorDetailsMap.set(key, {
                        ...existing,
                        name: name || existing.name,
                        business: biz || existing.business,
                        ssn: (cSSN ? getVal(row.getCell(cSSN)) : '') || existing.ssn,
                        address: (cAddr ? getVal(row.getCell(cAddr)) : '') || existing.address,
                        email: (cEmail ? getVal(row.getCell(cEmail)) : '') || existing.email,
                        phone: (cPhone ? getVal(row.getCell(cPhone)) : '') || existing.phone,
                    });
                });
                console.log(`[Validation] Loaded vendor info from ${fName}`);
                break; // Stop after first successful file match
            }
        }
    }
    await loadExternalVendors();

    // --- 2. Process Transaction Sheets ---
    // --- Constants for Column Headers ---
    const HEADERS = {
        DATE: ['date', 'txn date', 'transaction date'],
        DESC: ['description', 'desc', 'payee', 'name'],
        AMOUNT: ['amount', 'amt', 'value'],
        CATEGORY: ['category', 'cat', 'account_category'],
        SUBCAT: ['sub-category', 'sub-cat', 'subcategory', 'subcat'],
        VENDOR: ['vendor', 'vend', 'merchant', 'merchant name'],
        CUSTOMER: ['customer', 'cust', 'client'],
        DEBIT: ['debit', 'dr', 'withdrawal'],
        CREDIT: ['credit', 'cr', 'deposit']
    };

    // --- 2. Process Transaction Sheets ---
    const bankMapDefault = { date: null, desc: null, amount: null, category: null, subCat: null, vendor: null, customer: null };
    const ccMapDefault = { date: null, desc: null, amount: null, category: null, subCat: null, vendor: null, customer: null };

    function findCol(cellVal, headerList) {
        if (!cellVal) return false;
        const v = cellVal.toString().toLowerCase().trim();
        return headerList.some(h => v === h || v.includes(h)); // Relaxed matching
    }

    for (const config of sheetConfigs) {
        let sheetTotal = 0;
        let sheet = workbook.getWorksheet(config.name);
        if (!sheet) {
            sheet = workbook.worksheets.find(s => s.name.trim().toLowerCase() === config.name.trim().toLowerCase());
        }
        if (!sheet) {
            if (showChecker) console.log(`Sheet "${config.name}" NOT FOUND`);
            continue;
        }

        // --- Skip Ledger in main loop (handled separately in Step 3) ---
        if (config.name.toLowerCase() === 'ledger' || config.type.toLowerCase() === 'ledger') {
            continue;
        }

        const tStr = config.type.toLowerCase();
        const isCC = tStr.includes('cc') || tStr.includes('card') || tStr.includes('credit') || tStr.includes('amex');
        const pType = isCC ? 'cc' : 'bank';

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
                    // console.log(`[DEBUG SCAN] Sheet "${config.name}" Row ${r}: "${rowStr}"`);
                    if (HEAD_MATCH(rowStr)) {
                        headerRowIndex = r;
                        config.offset = r;
                        break;
                    }
                }
            }
        }

        function HEAD_MATCH(rowStr) {
            return (rowStr.includes('date') && (rowStr.includes('amount') || rowStr.includes('category')));
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
            // DEBUG: Force only Credit Card Transactions
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
        }

        let processedRows = 0;

        if (showDebug && sheet.name.includes('AX CC')) {
            console.log(`[DEBUG AX CC] Sheet Found. Config Offset: ${config.offset}. Total Rows in Sheet: ${sheet.rowCount}`);
        }

        sheet.eachRow((row, r) => {
            try {
                if (r <= config.offset) {
                    // if (sheet.name.includes('AX CC')) console.log(`[AX CC] Skipping Row ${r} (<= Offset ${config.offset})`);
                    return;
                }

                const vendorVal = map.vendor ? getVal(row.getCell(map.vendor)) : '';
                const customerVal = map.customer ? getVal(row.getCell(map.customer)) : '';
                const categoryVal = map.category ? getVal(row.getCell(map.category)) : '';
                const subCatVal = map.subCat ? getVal(row.getCell(map.subCat)) : '';
                let amount = map.amount ? getVal(row.getCell(map.amount)) : 0;

                // --- Initialize stats structures if needed ---
                // Category stats (for PL/BS) with per-sheet breakdown
                // --- REDUNDANT BLOCK REMOVED ---
                // The previous logic here (Lines 713-732) was using catLower as key, which conflicts with later logic using DisplayName.
                // We strip this block and rely on the consolidated logic below.
                // -------------------------------

                // Vendor / Customer stats tracking (consolidated below)


                // Customer stats with per-sheet breakdown (existing logic retained below)

                if (typeof amount !== 'number') {
                    // Sanitize string (remove $, commas, etc)
                    const amtStr = amount.toString().trim();
                    if (amtStr === '' || amtStr === '-') {
                        amount = 0;
                    } else {
                        const cleanAmt = amtStr.replace(/[^0-9.-]/g, '');
                        const parsed = parseFloat(cleanAmt);
                        if (isNaN(parsed)) {
                            if (showChecker && amtStr.length > 0) {
                                console.warn(`[WARNING] Sheet "${sheet.name}" Row ${r}: Could not parse amount "${amount}". Skipping.`);
                            }
                            return; // Skip this row safely
                        }
                        amount = parsed;
                    }
                }

                function excelDateToJS(serial) {
                    if (typeof serial !== 'number') return serial;
                    // Excel date offset: Jan 1, 1900.
                    // Note: Excel incorrectly treats 1900 as a leap year, so we subtract 2.
                    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
                    return date;
                }

                const rawDateInput = map.date ? getVal(row.getCell(map.date)) : '';
                let dateObj = rawDateInput;
                if (typeof rawDateInput === 'number' && rawDateInput > 10000) {
                    dateObj = excelDateToJS(rawDateInput);
                }

                const rawDesc = map.desc ? getVal(row.getCell(map.desc)).toString() : '';
                const matchDesc = rawDesc.toLowerCase();

                const displayDate = dateObj instanceof Date ? dateObj.toISOString().split('T')[0] :
                    (dateObj && typeof dateObj === 'string' ? dateObj : 'N/A');



                // Offset check
                if (r === config.offset + 1) {
                    const rowValues = row.values.map(v => (v ? v.toString().toLowerCase() : ''));
                    const rowText = rowValues.join(' ');
                    if (HEAD_MATCH(rowText)) {
                        offsetWarnings.push({ sheet: sheet.name, row: r, matches: ['Header Signature Detected'] });
                    }
                }

                // --- Robust Skipping Logic ---
                const hasDesc = rawDesc.trim().length > 0;
                const hasAmount = Math.abs(amount) > 0.0001;

                // 1. If completely empty, skip silently
                if (!rawDateInput && !hasDesc && !hasAmount) return;

                // 2. Strict Date Check
                if (displayDate === 'N/A' || !rawDateInput) {
                    // Only warn if there is other data
                    if ((hasDesc || hasAmount) && showChecker) {
                        console.warn(`[WARNING] Sheet "${sheet.name}" Row ${r}: Skipped due to invalid/missing Date. (Desc: "${rawDesc}", Amt: ${amount})`);
                    }
                    return;
                }

                // 3. Skip invalid amounts (already handled above, but double check)
                if (isNaN(amount)) return;

                processedRows++;

                if (config.flip) amount *= -1;

                // Accumulate Sheet Total (Net Flow)
                sheetTotal += amount;
                if (pType === 'cc') ccTotal += amount; else bankTotal += amount; // retained for legacy or verification

                if (!categoryVal && Math.abs(amount) > 0.01) {
                    if (pType === 'cc') uncategorizedCC++; else uncategorizedBank++;
                    uncategorizedDetails.push({ sheet: config.shortName, row: r, date: displayDate, desc: rawDesc });
                }

                // Define catLower for use in vendor/customer tracking
                const catLower = categoryVal ? categoryVal.toString().trim().toLowerCase() : null;

                if (categoryVal) {
                    const catStr = categoryVal.toString().trim();

                    if (!validCategories.has(catLower)) {
                        illegalCategories.push({ value: catStr, sheet: sheet.name, row: r, date: displayDate });
                    }

                    // Use Display Name for stats if available, else usage case
                    const displayCat = uniqueCategories.get(catLower)?.displayName || catStr;

                    if (!catStats[displayCat]) catStats[displayCat] = { total: 0, subCats: {}, sheets: {} };
                    catStats[displayCat].total += amount;

                    if (amount >= 0) {
                        if (!catStats[displayCat].add) catStats[displayCat].add = 0; // init check
                        catStats[displayCat].add = (catStats[displayCat].add || 0) + amount;
                    } else {
                        if (!catStats[displayCat].sub) catStats[displayCat].sub = 0;
                        catStats[displayCat].sub = (catStats[displayCat].sub || 0) + amount;
                    }

                    // Track Sheets for Display Category
                    if (!catStats[displayCat].sheets) catStats[displayCat].sheets = {};
                    if (!catStats[displayCat].sheets[config.name]) {
                        catStats[displayCat].sheets[config.name] = { add: 0, sub: 0, total: 0 };
                    }
                    const cDS = catStats[displayCat].sheets[config.name];
                    if (amount >= 0) cDS.add += amount; else cDS.sub += amount;
                    cDS.total += amount;

                    const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';

                    // Validate subcategory if one is provided
                    if (subCatVal && sName !== '(No Sub-Cat)') {
                        const sLower = sName.toLowerCase();
                        if (!validSubCategories.has(sLower)) {
                            illegalSubCategories.push({
                                value: sName,
                                category: displayCat,
                                sheet: config.shortName,
                                row: r,
                                date: displayDate
                            });
                        }
                    }

                    catStats[displayCat].subCats[sName] = (catStats[displayCat].subCats[sName] || 0) + amount;

                    // Capture Details
                    if (showDetails && (catLower === targetDetailsCategory || displayCat.toLowerCase() === targetDetailsCategory)) {
                        detailsRows.push({
                            date: displayDate,
                            desc: rawDesc,
                            subCat: sName,
                            amount: amount,
                            sheet: sheet.name,
                            row: r
                        });
                    }
                }

                if (vendorVal) {
                    const vStr = vendorVal.toString().trim();
                    const vLower = vStr.toLowerCase();
                    if (!validVendors.has(vLower)) illegalVendors.push({ value: vStr, sheet: sheet.name, row: r, date: displayDate });

                    const displayVendor = validVendors.get(vLower) || vStr;
                    // Standardize: Expenses are Positive. Income is Negative.
                    const vendAmount = amount * -1;

                    if (!vendorStats[displayVendor]) {
                        vendorStats[displayVendor] = { total: 0, add: 0, sub: 0, sheets: {} };
                    }
                    if (!vendorStats[displayVendor].sheets) vendorStats[displayVendor].sheets = {};

                    if (vendAmount >= 0) {
                        vendorStats[displayVendor].add += vendAmount;
                    } else {
                        vendorStats[displayVendor].sub += vendAmount;
                    }
                    vendorStats[displayVendor].total += vendAmount;

                    if (!vendorStats[displayVendor].sheets[config.name]) {
                        vendorStats[displayVendor].sheets[config.name] = { add: 0, sub: 0, total: 0 };
                    }
                    const vSheet = vendorStats[displayVendor].sheets[config.name];
                    if (vendAmount >= 0) vSheet.add += vendAmount; else vSheet.sub += vendAmount;
                    vSheet.total += vendAmount;


                    // Track which report type this vendor is used in
                    if (catLower) {
                        const catConf = uniqueCategories.get(catLower);
                        if (catConf && catConf.report) {
                            if (!vendorReportUsage.has(displayVendor)) {
                                vendorReportUsage.set(displayVendor, new Map());
                            }
                            const reportMap = vendorReportUsage.get(displayVendor);
                            if (!reportMap.has(catConf.report)) {
                                reportMap.set(catConf.report, []);
                            }
                            // Store first 2 examples per report type
                            const displayCat = catConf.displayName || categoryVal.toString().trim();
                            if (reportMap.get(catConf.report).length < 2) {
                                reportMap.get(catConf.report).push({
                                    sheet: sheet.name,
                                    row: r,
                                    category: displayCat,
                                    date: displayDate
                                });
                            }
                        }
                    }

                    const is1099 = vendor1099Map.get(vLower);
                    if (is1099) {
                        const t = (is1099.type || 'NEC');
                        if (!vendor1099Stats[t]) vendor1099Stats[t] = {};
                        if (!vendor1099Stats[t][displayVendor]) {
                            vendor1099Stats[t][displayVendor] = 0;
                        }
                        vendor1099Stats[t][displayVendor] += vendAmount;
                    }
                }
                if (customerVal) {
                    const cStr = customerVal.toString().trim();
                    const cLower = cStr.toLowerCase();
                    if (!validCustomers.has(cLower)) illegalCustomers.push({ value: cStr, sheet: sheet.name, row: r, date: displayDate });

                    const displayCustomer = validCustomers.get(cLower) || cStr;

                    // Existing customer stats logic (kept unchanged)
                    if (!customerStats[displayCustomer]) {
                        customerStats[displayCustomer] = { add: 0, sub: 0, total: 0, sheets: {} };
                    }
                    if (amount >= 0) {
                        customerStats[displayCustomer].add += amount;
                    } else {
                        customerStats[displayCustomer].sub += amount;
                    }
                    customerStats[displayCustomer].total += amount;
                    if (!customerStats[displayCustomer].sheets[config.name]) {
                        customerStats[displayCustomer].sheets[config.name] = { add: 0, sub: 0, total: 0 };
                    }
                    const cSheetStat = customerStats[displayCustomer].sheets[config.name];
                    if (amount >= 0) cSheetStat.add += amount; else cSheetStat.sub += amount;
                    cSheetStat.total += amount;

                    // Track which report type this customer is used in
                    if (catLower) {
                        const catConf = uniqueCategories.get(catLower);
                        if (catConf && catConf.report) {
                            if (!customerReportUsage.has(displayCustomer)) {
                                customerReportUsage.set(displayCustomer, new Map());
                            }
                            const reportMap = customerReportUsage.get(displayCustomer);
                            if (!reportMap.has(catConf.report)) {
                                reportMap.set(catConf.report, []);
                            }
                            // Store first 2 examples per report type
                            const displayCat = catConf.displayName || categoryVal.toString().trim();
                            if (reportMap.get(catConf.report).length < 2) {
                                reportMap.get(catConf.report).push({
                                    sheet: sheet.name,
                                    row: r,
                                    category: displayCat,
                                    date: displayDate
                                });
                            }
                        }
                    }
                }
            } catch (rowError) {
                if (showChecker) console.error(`[ERROR] Sheet "${sheet.name}" Row ${r}: Crash detected. ${rowError.message}`);
            }
        });

        console.log(`[Sheet Stats] "${config.shortName}": Processed ${processedRows} rows. Total Change: ${sheetTotal.toFixed(2)}`);

        if (config.linkedAccount) {
            const linkName = config.linkedAccount;
            const lConf = uniqueCategories.get(linkName.toLowerCase());
            const aType = (lConf && lConf.accountType) ? lConf.accountType.toString().toLowerCase() : '';
            const isAsset = aType.includes('asset') || aType.includes('bank') || aType.includes('cash');
            const isLiability = aType.includes('liability') || aType.includes('credit') || aType.includes('cc') || aType.includes('payable') || aType.includes('loan');

            if (!catStats[linkName]) catStats[linkName] = { total: 0, subCats: {}, sheets: {} };
            if (!catStats[linkName].sheets) catStats[linkName].sheets = {};

            const previous = catStats[linkName].total;
            // Polarity Logic: Assets increase with positive flow (Income). Liabilities/Equity decrease.
            // Balance Sheet Items:
            // Asset: Balance += sheetTotal
            // Liability/Equity: Balance -= sheetTotal
            if (isAsset) {
                catStats[linkName].total += sheetTotal;
            } else {
                catStats[linkName].total -= sheetTotal;
            }

            // Track sheet-level contribution to BS account
            if (!catStats[linkName].sheets[config.name]) {
                catStats[linkName].sheets[config.name] = { add: 0, sub: 0, total: 0 };
            }
            const sStat = catStats[linkName].sheets[config.name];
            if (sheetTotal >= 0) sStat.add += sheetTotal; else sStat.sub += sheetTotal;

            if (isAsset) sStat.total += sheetTotal; else sStat.total -= sheetTotal;

            console.log(`[Linkage Logic] Applied ${config.shortName} Total (${sheetTotal.toFixed(2)}) to ${isAsset ? 'Asset' : 'Liability'} "${linkName}". Balance: ${previous.toFixed(2)} -> ${catStats[linkName].total.toFixed(2)}`);
        } else {
            console.log(`[Linkage Logic] Sheet "${config.name}" (Type: ${config.type}) has NO LINKED ACCOUNT. Total (${sheetTotal.toFixed(2)}) NOT applied to any Balance Sheet asset.`);
        }
    }

    // --- 3. Process Ledger ---
    // --- 3. Process Ledger ---
    // Find Ledger configuration from Setup
    const ledgerConfig = sheetConfigs.find(c => c.name.toLowerCase() === 'ledger');

    if (!ledgerConfig) {
        console.error(`\n[ERROR] Checker Failed: "Ledger" configuration not found in Setup > SheetInfo.`);
        console.error(`Your workbook contains a "Ledger" sheet, so you MUST define it in the SheetInfo table.`);
        console.error(`Please add a row to SheetInfo:`);
        console.error(`   Sheet Name: Ledger`);
        console.error(`   Type:       Ledger`);
        console.error(`   Flip:       No`);
        console.error(`   Offset:     3 (or your header row)`);
        process.exit(1);
    }

    const ledgerHeaderRow = ledgerConfig.offset || 3; // Default to 3 if offset is 0/undefined but config exists

    // Dynamic Mapping for Ledger
    const ledgerMap = { date: null, desc: null, category: null, subCat: null, vendor: null, customer: null, dr: null, cr: null };
    const ledgerHeader = ledgerSheet.getRow(ledgerHeaderRow);

    ledgerHeader.eachCell((cell, colNumber) => {
        const val = getVal(cell);

        if (findCol(val, HEADERS.DATE)) ledgerMap.date = colNumber;
        else if (findCol(val, HEADERS.DESC)) ledgerMap.desc = colNumber;
        else if (findCol(val, HEADERS.SUBCAT)) ledgerMap.subCat = colNumber;
        else if (findCol(val, HEADERS.CATEGORY)) ledgerMap.category = colNumber;
        else if (findCol(val, HEADERS.VENDOR)) ledgerMap.vendor = colNumber;
        else if (findCol(val, HEADERS.CUSTOMER)) ledgerMap.customer = colNumber;
        else if (findCol(val, HEADERS.DEBIT)) ledgerMap.dr = colNumber;
        else if (findCol(val, HEADERS.CREDIT)) ledgerMap.cr = colNumber;
    });

    // Validate that critical headers were found
    if (!ledgerMap.date || !ledgerMap.category) {
        console.error(`\n[ERROR] Ledger sheet: Required headers not found at row ${ledgerHeaderRow}`);
        console.error(`Expected headers: Date, Category (and optionally: Description, Debit, Credit, Vendor, Customer, Sub-Category)`);
        console.error(`Found mapping: ${JSON.stringify(ledgerMap)}`);
        console.error(`\nSetup sheet specifies Ledger header row at: ${ledgerHeaderRow}`);
        console.error(`Please verify your Ledger sheet has proper headers at row ${ledgerHeaderRow}.`);
        process.exit(1);
    }

    if (showChecker) {
        console.log(`\nProcessing "Ledger":`);
        console.log(`  Mapping: ${JSON.stringify(ledgerMap)}`);
    }

    let ledgerTotal = 0;
    let ledgerRows = 0;
    ledgerSheet.eachRow((row, r) => {
        try {
            if (r <= ledgerHeaderRow) return;
            const rawDate = ledgerMap.date ? getVal(row.getCell(ledgerMap.date)) : '';
            const rawDesc = ledgerMap.desc ? getVal(row.getCell(ledgerMap.desc)) : '';
            const cat = ledgerMap.category ? getVal(row.getCell(ledgerMap.category)) : '';

            // Ledger SubCat support
            const subCatVal = ledgerMap.subCat ? getVal(row.getCell(ledgerMap.subCat)) : '';

            const dr = (ledgerMap.dr && row.getCell(ledgerMap.dr).value) ? (parseFloat(getVal(row.getCell(ledgerMap.dr))) || 0) : 0;
            const cr = (ledgerMap.cr && row.getCell(ledgerMap.cr).value) ? (parseFloat(getVal(row.getCell(ledgerMap.cr))) || 0) : 0;
            const vendorVal = ledgerMap.vendor ? getVal(row.getCell(ledgerMap.vendor)) : '';
            const customerVal = ledgerMap.customer ? getVal(row.getCell(ledgerMap.customer)) : '';

            // Skip truly empty rows or rows without dates (user requirement)
            if (!rawDate && !cat && !rawDesc && !dr && !cr) return;

            if (!rawDate) {
                if (cat || dr || cr) {
                    if (showChecker) console.log(`Ledger Row ${r}: SKIPPED (Missing Date). Rows must have dates.`);
                }
                return;
            }

            if (cat) {
                const catStr = cat.toString().trim();
                const catLower = catStr.toLowerCase();
                const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');

                if (!validCategories.has(catLower)) illegalCategories.push({ value: catStr, sheet: 'Ledger', row: r, date: displayDate });

                const conf = uniqueCategories.get(catLower);

                const displayCat = conf?.displayName || catStr;
                const isAsset = conf && conf.accountType && conf.accountType.toLowerCase().includes('asset');
                const impact = isAsset ? (dr - cr) : (cr - dr);

                if (impact !== 0 || catStr) {
                    ledgerRows++;
                    ledgerTotal += impact;
                }

                if (!catStats[displayCat]) catStats[displayCat] = { total: 0, subCats: {}, sheets: {} };
                if (!catStats[displayCat].sheets) catStats[displayCat].sheets = {};
                catStats[displayCat].total += impact;
                // Track Sheet contribution (using canonical name)
                if (!catStats[displayCat].sheets[ledgerConfig.name]) catStats[displayCat].sheets[ledgerConfig.name] = { add: 0, sub: 0, total: 0 };
                const lSheet = catStats[displayCat].sheets[ledgerConfig.name];
                if (impact >= 0) lSheet.add += impact; else lSheet.sub += impact;
                lSheet.total += impact;

                // Ledger SubCat aggregation
                const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';

                // Validate subcategory if one is provided
                if (subCatVal && sName !== '(No Sub-Cat)') {
                    const sLower = sName.toLowerCase();
                    if (!validSubCategories.has(sLower)) {
                        illegalSubCategories.push({
                            value: sName,
                            category: displayCat,
                            sheet: 'Ledger',
                            row: r,
                            date: displayDate
                        });
                    }
                }

                catStats[displayCat].subCats[sName] = (catStats[displayCat].subCats[sName] || 0) + impact;

                // Capture Details
                if (showDetails && (catLower === targetDetailsCategory || displayCat.toLowerCase() === targetDetailsCategory)) {
                    detailsRows.push({
                        date: displayDate,
                        desc: rawDesc,
                        subCat: sName,
                        amount: impact,
                        sheet: 'Ledger',
                        row: r
                    });
                }

                // Vendor / Customer Stats from Ledger
                if (vendorVal) {
                    const vStr = vendorVal.toString().trim();
                    const vLower = vStr.toLowerCase();
                    if (!validVendors.has(vLower)) illegalVendors.push({ value: vStr, sheet: 'Ledger', row: r, date: displayDate });

                    const displayVendor = validVendors.get(vLower) || vStr;
                    // Vendor: Net Debit (Expense)
                    // Vendor Stats Structure: { total, add, sub, sheets }
                    // Vendor Stats Structure: { total, add, sub, sheets }
                    const impactVal = (dr - cr);
                    if (!vendorStats[displayVendor]) vendorStats[displayVendor] = { total: 0, add: 0, sub: 0, sheets: {} };
                    if (!vendorStats[displayVendor].sheets) vendorStats[displayVendor].sheets = {};

                    if (impactVal >= 0) vendorStats[displayVendor].add += impactVal; else vendorStats[displayVendor].sub += impactVal;
                    vendorStats[displayVendor].total += impactVal;

                    // Ensure sheets exists before accessing
                    if (!vendorStats[displayVendor].sheets) vendorStats[displayVendor].sheets = {};
                    if (!vendorStats[displayVendor].sheets[ledgerConfig.name]) vendorStats[displayVendor].sheets[ledgerConfig.name] = { add: 0, sub: 0, total: 0 };
                    const vSheet = vendorStats[displayVendor].sheets[ledgerConfig.name];
                    if (impactVal >= 0) vSheet.add += impactVal; else vSheet.sub += impactVal;
                    vSheet.total += impactVal;

                    // Track which report type this vendor is used in
                    const catConf = uniqueCategories.get(catLower);
                    if (catConf && catConf.report) {
                        if (!vendorReportUsage.has(displayVendor)) {
                            vendorReportUsage.set(displayVendor, new Map());
                        }
                        const reportMap = vendorReportUsage.get(displayVendor);
                        if (!reportMap.has(catConf.report)) {
                            reportMap.set(catConf.report, []);
                        }
                        // Store first 2 examples per report type
                        if (reportMap.get(catConf.report).length < 2) {
                            reportMap.get(catConf.report).push({
                                sheet: 'Ledger',
                                row: r,
                                category: displayCat,
                                date: displayDate
                            });
                        }
                    }

                    const is1099 = vendor1099Map.get(vLower);
                    if (is1099) {
                        const t = (is1099.type || 'NEC');
                        if (!vendor1099Stats[t]) vendor1099Stats[t] = {};
                        if (!vendor1099Stats[t][displayVendor]) {
                            vendor1099Stats[t][displayVendor] = 0;
                        }
                        vendor1099Stats[t][displayVendor] += impactVal;
                    }
                }
                if (customerVal) {
                    const cStr = customerVal.toString().trim();
                    const cLower = cStr.toLowerCase();
                    if (!validCustomers.has(cLower)) illegalCustomers.push({ value: cStr, sheet: 'Ledger', row: r, date: displayDate });

                    const displayCustomer = validCustomers.get(cLower) || cStr;
                    // Customer: Net Credit (Income)
                    const custImpact = (cr - dr);
                    if (!customerStats[displayCustomer]) customerStats[displayCustomer] = { total: 0, add: 0, sub: 0, sheets: {} };
                    if (!customerStats[displayCustomer].sheets) customerStats[displayCustomer].sheets = {};

                    if (custImpact >= 0) customerStats[displayCustomer].add += custImpact; else customerStats[displayCustomer].sub += custImpact;
                    customerStats[displayCustomer].total += custImpact;

                    if (!customerStats[displayCustomer].sheets[ledgerConfig.name]) customerStats[displayCustomer].sheets[ledgerConfig.name] = { add: 0, sub: 0, total: 0 };
                    const cSheet = customerStats[displayCustomer].sheets[ledgerConfig.name];
                    if (custImpact >= 0) cSheet.add += custImpact; else cSheet.sub += custImpact;
                    cSheet.total += custImpact;

                    // Track which report type this customer is used in
                    const catConf = uniqueCategories.get(catLower);
                    if (catConf && catConf.report) {
                        if (!customerReportUsage.has(displayCustomer)) {
                            customerReportUsage.set(displayCustomer, new Map());
                        }
                        const reportMap = customerReportUsage.get(displayCustomer);
                        if (!reportMap.has(catConf.report)) {
                            reportMap.set(catConf.report, []);
                        }
                        // Store first 2 examples per report type
                        if (reportMap.get(catConf.report).length < 2) {
                            reportMap.get(catConf.report).push({
                                sheet: 'Ledger',
                                row: r,
                                category: displayCat,
                                date: displayDate
                            });
                        }
                    }
                }

                // (Integration block removed - handled via standard catStats logic)
            }
        } catch (ledgerRowError) {
            if (showChecker) console.error(`[ERROR] Ledger Row ${r}: ${ledgerRowError.message}`);
        }
    });

    console.log(`[Sheet Stats] "${ledgerConfig.shortName}": Processed ${ledgerRows} rows. Total Change: ${ledgerTotal.toFixed(2)}`);
    console.log(`[Linkage Logic] Sheet "${ledgerConfig.name}" (Type: Ledger) has NO LINKED ACCOUNT. Total (${ledgerTotal.toFixed(2)}) NOT applied to any Balance Sheet asset.`);

    // --- 4. Prepare Reports ---
    const reports = { pl: [], bs: [] };
    // Filter by P&L report type using the Map values
    const pnlNames = Array.from(uniqueCategories.values())
        .filter(conf => conf.report === 'P&L')
        .map(conf => conf.displayName)
        .sort();

    reports.pl = pnlNames.map(n => ({
        label: n,
        value: catStats[n] ? catStats[n].total : 0,
        sheets: catStats[n] ? catStats[n].sheets : {},
        subCats: catStats[n] ? catStats[n].subCats : {}
    }));
    const netIncome = reports.pl.reduce((a, b) => a + b.value, 0);

    // Balance Sheet Items
    const bsNames = Array.from(uniqueCategories.values())
        .filter(conf => conf.report === 'Balance Sheet')
        .map(conf => conf.displayName)
        .sort();

    const bsItems = bsNames.map(n => {
        let val = catStats[n] ? catStats[n].total : 0;
        // Since linkage now handles polarity (Asset += sheetTotal, Liability -= sheetTotal),
        // catStats[n].total is the actual accounting balance.
        return {
            label: n,
            value: val,
            sheets: catStats[n] ? catStats[n].sheets : {},
            subCats: catStats[n] ? catStats[n].subCats : {}
        };
    });

    reports.bs = [
        ...bsItems
    ];

    // Prepare Vendor / Customer Reports
    reports.vendors = Object.keys(vendorStats).map(v => ({
        label: v,
        add: vendorStats[v].add || 0,
        sub: vendorStats[v].sub || 0,
        value: vendorStats[v].total || 0,
        sheets: vendorStats[v].sheets || {}
    })).sort((a, b) => b.value - a.value);
    reports.vendors1099NEC = Object.keys(vendor1099Stats.NEC).map(v => ({ label: v, value: vendor1099Stats.NEC[v] })).sort((a, b) => b.value - a.value);
    reports.vendors1099INT = Object.keys(vendor1099Stats.INT).map(v => ({ label: v, value: vendor1099Stats.INT[v] })).sort((a, b) => b.value - a.value);

    reports.customers = Object.keys(customerStats).map(c => ({
        label: c,
        add: customerStats[c].add,
        sub: customerStats[c].sub,
        value: customerStats[c].total,
        sheets: customerStats[c].sheets
    })).sort((a, b) => b.value - a.value);
    // Vendor report with sub breakdown
    reports.vendors = Object.keys(vendorStats).map(v => ({
        label: v,
        add: vendorStats[v].add,
        sub: vendorStats[v].sub,
        value: vendorStats[v].total,
        sheets: vendorStats[v].sheets
    })).sort((a, b) => b.value - a.value);

    // --- 5. Console Output ---
    // --- 5. Console Output ---
    function printDetailedTable(title, rows, sheetList, sheetNameMap = {}, label = "Category") {
        console.log(`\n--- ${title} ---`);
        if (!rows.length) { console.log('(No Data)'); return; }

        const LABEL_WIDTH = 30;
        const COL_WIDTH = 12;

        // Header 1
        const SECTION_WIDTH = (sheetList.length + 1) * COL_WIDTH;

        const h1 = " ".repeat(LABEL_WIDTH) +
            "Net".padStart(COL_WIDTH) + " | " +
            "Additions".padStart(SECTION_WIDTH) + " | " +
            "Subtractions".padStart(SECTION_WIDTH);
        console.log(h1);

        // Header 2
        let h2 = label.padEnd(LABEL_WIDTH);
        // Net (Grand Total)
        h2 += "Grand Total".padStart(COL_WIDTH) + " | ";

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

            // Net
            const signedNet = r.value;
            line += signedNet.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) + " | ";

            // Additions
            let rowAddTotal = 0;
            sheetList.forEach(s => {
                const val = (sheets[s] ? sheets[s].add : 0) || 0;
                rowAddTotal += val;
                // Use absolute for display in additions column
                const disp = Math.abs(val);
                line += disp === 0 ? " ".padStart(COL_WIDTH) : disp.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH);
            });
            line += Math.abs(rowAddTotal).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) + " | ";

            // Subtractions
            let rowSubTotal = 0;
            sheetList.forEach(s => {
                const val = (sheets[s] ? sheets[s].sub : 0) || 0;
                rowSubTotal += val;
                // Use absolute for display in subtractions column
                const disp = Math.abs(val);
                line += disp === 0 ? " ".padStart(COL_WIDTH) : disp.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH);
            });
            line += Math.abs(rowSubTotal).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH);

            console.log(line);

            // Print subcategory breakdowns if available
            if (r.subCats && Object.keys(r.subCats).length > 0) {
                Object.entries(r.subCats).forEach(([subName, subTotal]) => {
                    if (subName === '(No Sub-Cat)') return; // Skip the default
                    const subLine = `  > ${subName}`.substring(0, LABEL_WIDTH - 1).padEnd(LABEL_WIDTH) +
                        subTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(COL_WIDTH) +
                        " ".repeat(h2.length - LABEL_WIDTH - COL_WIDTH); // Fill rest with spaces
                    console.log(subLine);
                });
            }
        });
    }

    // Collect all sheet names for columns (Order: Configured Sheets...)
    const distinctSheets = new Set();
    const sheetNameMap = {}; // Map full name -> short name
    sheetConfigs.forEach(s => {
        distinctSheets.add(s.name);
        sheetNameMap[s.name] = s.shortName;
    });
    // Ensure Ledger is included if it was processed
    if (Object.keys(catStats).some(c => catStats[c].sheets && catStats[c].sheets[ledgerConfig.name])) {
        distinctSheets.add(ledgerConfig.name);
        sheetNameMap[ledgerConfig.name] = ledgerConfig.shortName || ledgerConfig.name;
    }
    const reportSheetList = Array.from(distinctSheets);

    // PL sub report with two-level header
    if (showAll || showPL || showPLSub) {
        if (showPLSub) {
            printDetailedTable('PROFIT & LOSS (Detailed)', reports.pl, reportSheetList, sheetNameMap);
        } else {
            console.log(`\n--- PROFIT & LOSS ---`);
            if (!reports.pl.length) console.log('(No Data)');
            else {
                const max = Math.max(...reports.pl.map(r => r.label.length), 10);
                reports.pl.forEach(r => {
                    console.log(`${r.label.padEnd(max + 5)} : ${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
                });
            }
        }
        console.log(`\n=== NET INCOME: ${netIncome.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ===\n`);
    }
    // BS sub report with two-level header
    if (showAll || showBS || showBSSub) {
        if (showBSSub) {
            printDetailedTable('BALANCE SHEET (Detailed)', reports.bs, reportSheetList, sheetNameMap);
        } else {
            console.log(`\n--- BALANCE SHEET ---`);
            if (!reports.bs.length) console.log('(No Data)');
            else {
                const max = Math.max(...reports.bs.map(r => r.label.length), 10);
                reports.bs.forEach(r => {
                    console.log(`${r.label.padEnd(max + 5)} : ${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
                });
            }
        }
        console.log('');
    }
    if (showAll || showVendor) {
        // printSection('VENDOR SPENDING', reports.vendors); <-- Replacing with Detailed Table
        console.log(`\n--- VENDOR SPENDING ---`);
        if (reports.vendors.length === 0) console.log('(No Data)');
        else {
            const h = `Vendor`.padEnd(30) +
                `Total`.padStart(15) +
                `  1099 Type`.padEnd(12) +
                `Required`.padEnd(10);
            console.log(h);
            console.log('-'.repeat(h.length));

            reports.vendors.forEach(r => {
                const info = vendor1099Map.get(r.label.toLowerCase()) || { type: '', req: '' };

                // Determine if vendor actually qualifies for 1099 reporting
                let displayReq = '';
                if (info.type) {
                    const threshold = info.type === 'INT' ? 0 : 600; // INT has $0 threshold, NEC has $600
                    const meetsThreshold = r.value > 0 && r.value >= threshold;
                    displayReq = meetsThreshold ? 'YES' : '';
                }

                console.log(
                    `${r.label.substring(0, 29).padEnd(30)}` +
                    `${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}` +
                    `  ${(info.type || '').padEnd(12)}` +
                    `${displayReq.padEnd(10)}`
                );
            });
        }

        // Explicitly warn about unknown vendors if requested
        if (illegalVendors.length > 0) {
            console.log(`\n[!] WARNING: ${illegalVendors.length} transactions have unknown Vendors.`);
            const uniqueUnknown = Array.from(new Set(illegalVendors.map(i => i.value)));
            console.log(`    Unknown Vendors: ${uniqueUnknown.join(', ')}`);
        }
    }
    // Helper: 1099 Detail Printer - NOW SILENT (CSV ONLY) or Minimal
    // User said: "1099 just prints to a file and report the print"
    const print1099 = (type, list, threshold) => {
        // No Console Output here
    };

    if (showAll || show1099) {
        print1099('NEC', reports.vendors1099NEC, 600);
        print1099('INT', reports.vendors1099INT, 0);

        // Generate CSV if any data found
        const all1099 = [
            ...reports.vendors1099NEC.map(x => ({ ...x, form: 'NEC', threshold: 600 })),
            ...reports.vendors1099INT.map(x => ({ ...x, form: 'INT', threshold: 0 }))
        ];

        // Filter by threshold & polarity
        const csvRows = [];
        all1099.forEach(r => {
            // Polarity Logic: 
            // - Value is now POSITIVE for Expenses.
            // - Threshold Check: Value > Threshold.
            // - Filter: Only Net Expenses (Positive Values).

            const isExpense = r.value > 0; // Net Payment

            if (isExpense && r.value >= r.threshold) {
                const d = vendorDetailsMap.get(r.label.toLowerCase()) || {};
                csvRows.push({
                    ...d,
                    amount: r.value, // Positive value for IRS
                    form: r.form
                });
            }
        });

        if (csvRows.length > 0) {
            const payerName = payerInfo['companyname'] || payerInfo['payername'] || payerInfo['businessname'] || payerInfo['name'] || 'Unknown_Payer';
            const safeName = payerName.replace(/[^a-z0-9]/gi, '_');
            const csvPath = path.join(path.dirname(filename), `${safeName}-1099.csv`);

            // Payer fields
            // "Name", "TIN", "Address", "City", "State", "Zip", "Country", "Email", "Phone" -> from PayerInfo
            // Since PayerInfo is loose KV, we try standard keys
            const p = payerInfo;
            const pName = payerName;
            const pTIN = p['tin'] || p['ein'] || p['taxid'] || '';
            const pAddr = p['address'] || '';
            const pCity = p['city'] || '';
            const pState = p['state'] || '';
            const pZip = p['zipcode'] || p['zip'] || '';
            const pCountry = p['country'] || '';
            const pEmail = p['email'] || '';
            const pPhone = p['phone'] || p['phonenumber'] || '';

            // Build CSV Content
            // Header: Payer... , Recipient... , Amount, Form
            const headers = [
                'Payer Name', 'Payer TIN', 'Payer Address', 'Payer City', 'Payer State', 'Payer Zip', 'Payer Country', 'Payer Email', 'Payer Phone',
                'Recipient Name', 'Recipient Business Name', 'Recipient TIN', 'Recipient Address', 'Recipient Email', 'Recipient Phone',
                'Amount', 'Form 1099 Type'
            ];

            const fileContent = [headers.join(',')];

            csvRows.forEach(r => {
                const row = [
                    `"${pName}"`, `"${pTIN}"`, `"${pAddr}"`, `"${pCity}"`, `"${pState}"`, `"${pZip}"`, `"${pCountry}"`, `"${pEmail}"`, `"${pPhone}"`,
                    `"${r.name || ''}"`, `"${r.business || ''}"`, `"${r.ssn || ''}"`, `"${r.address || ''}"`, `"${r.email || ''}"`, `"${r.phone || ''}"`,
                    r.amount.toFixed(2),
                    r.form
                ];
                fileContent.push(row.join(','));
            });

            fs.writeFileSync(csvPath, fileContent.join('\n'));
            console.log(`\n[SUCCESS] Generated 1099 CSV: ${csvPath}`);
        }
    }
    // Customer report (already handles sub)
    if (showAll || showCustomer || showCustomerSub) {
        if (showCustomerSub) {
            printDetailedTable('CUSTOMER INCOME (Detailed)', reports.customers, reportSheetList, sheetNameMap, "Customer");
        } else {
            console.log(`\n--- CUSTOMER INCOME ---`);
            if (reports.customers.length === 0) console.log('(No Data)');
            else {
                const h = `Customer`.padEnd(30) + `Net`.padStart(15) + `Additions`.padStart(15) + `Subtractions`.padStart(15);
                console.log(h);
                console.log('-'.repeat(h.length));
                reports.customers.forEach(r => {
                    console.log(`${r.label.substring(0, 29).padEnd(30)}` +
                        `${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}` +
                        `${r.add.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}` +
                        `${r.sub.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
                });
            }
        }
    }

    // Vendor sub report
    if (showAll || showVendorSub) {
        if (showVendorSub) {
            printDetailedTable('VENDOR SPENDING (Detailed)', reports.vendors, reportSheetList, sheetNameMap, "Vendor");
        } else {
            // Already shown in main "Vendor Spending" block if showVendor is on, 
            // but if showVendorSub is ON and showVendor is OFF (edge case), we show detailed.
            // If both, we might duplicate? 
            // The main vendor block shows 1099 info. This shows stats.
            // Let's assume if showVendor is on, we don't need this basic block unless sub is requested.
            // But the original code printed it.
        }
    }

    if (showDetails) {
        console.log(`\n--- DETAILS: "${targetDetailsCategory}" ---`);
        if (detailsRows.length === 0) {
            console.log('(No matching transactions found)');
        } else {
            console.log(`Date`.padEnd(12) + `Description`.padEnd(35) + `Sub-Cat`.padEnd(20) + `Amount`.padStart(12) + `  Source`);
            console.log(`-`.repeat(85));
            let total = 0;
            detailsRows.sort((a, b) => a.date.localeCompare(b.date));
            detailsRows.forEach(r => {
                total += r.amount;
                console.log(
                    `${r.date.padEnd(12)}${r.desc.substring(0, 34).padEnd(35)}${r.subCat.substring(0, 19).padEnd(20)}${r.amount.toFixed(2).padStart(12)}  ${r.sheet} (Row ${r.row})`
                );
            });
            console.log(`-`.repeat(85));
            console.log(`TOTAL`.padEnd(67) + total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(12));
        }
    }

    // --- Global Integrity Check: Total Tie ---
    // The sum of all net flows across all transaction sheets (after polarity)
    // must equal (Net Income) + (Sum of Balance Sheet changes).
    let globalFlowTotal = 0;
    const sheetFlows = [];
    sheetConfigs.forEach(conf => {
        // Find total for this sheet across all categories
        let sheetSum = 0;
        Object.keys(catStats).forEach(cat => {
            if (catStats[cat].sheets && catStats[cat].sheets[conf.name]) {
                sheetSum += catStats[cat].sheets[conf.name].total;
            }
        });
        globalFlowTotal += sheetSum;
        sheetFlows.push({ name: conf.shortName, flow: sheetSum });
    });

    // Net Income + Balance Sheet Changes
    const bsChangeTotal = reports.bs.reduce((a, b) => a + b.value, 0);
    const accountingTotal = netIncome + bsChangeTotal;

    console.log('\n--- GLOBAL INTEGRITY CHECK ---');
    console.log(`Total Sheet Flow (Bank + CC + Ledger) : ${globalFlowTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
    console.log(`Sum of Net Income + BS Changes         : ${accountingTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);

    const discrepancy = Math.abs(globalFlowTotal - accountingTotal);
    if (discrepancy < 0.01) {
        console.log('✅ [PASS] Total Tie: Financial reports are internally consistent.');
    } else {
        console.warn(`[!] WARNING: Total Tie mismatch. Discrepancy: ${discrepancy.toFixed(2)}`);
        console.warn('    Check for unlinked transaction sheets or manual category overrides.');
    }

    const hasIssues = uncategorizedDetails.length > 0 || illegalCategories.length > 0 || illegalVendors.length > 0 || illegalCustomers.length > 0 || illegalSubCategories.length > 0;
    if (hasIssues) {
        console.log('\n--- DATA INTEGRITY ISSUES ---');
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet),
            ...illegalSubCategories.map(x => x.sheet)
        ]);
        issueSheetsFound.forEach(s => {
            console.log(`\n>> Tab: ${s.toUpperCase()}`);
            const uncat = uncategorizedDetails.filter(x => x.sheet === s);
            if (uncat.length) console.log(`  [!] ${uncat.length} rows missing category`);
            const cats = new Set(illegalCategories.filter(x => x.sheet === s).map(x => x.value));
            if (cats.size) console.log(`  [!] Illegal Categories: ${Array.from(cats).join(', ')}`);
            const vends = new Set(illegalVendors.filter(x => x.sheet === s).map(x => x.value));
            if (vends.size) console.log(`  [!] Unknown Vendors: ${Array.from(vends).join(', ')}`);
            const custs = new Set(illegalCustomers.filter(x => x.sheet === s).map(x => x.value));
            if (custs.size) console.log(`  [!] Unknown Customers: ${Array.from(custs).join(', ')}`);
            const subCats = new Set(illegalSubCategories.filter(x => x.sheet === s).map(x => x.value));
            if (subCats.size) console.log(`  [!] Illegal Sub-Categories: ${Array.from(subCats).join(', ')}`);

            if (showChecker) {
                uncat.forEach(x => console.log(`      - [${x.date}] Row ${x.row}: MISSING CATEGORY ("${x.desc}")`));
                illegalCategories.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL CATEGORY "${x.value}"`));
                illegalVendors.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN VENDOR "${x.value}"`));
                illegalCustomers.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN CUSTOMER "${x.value}"`));
                illegalSubCategories.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL SUB-CATEGORY "${x.value}" in category "${x.category}"`));
            }
        });



        // Combined Missing Vendors Report
        if (illegalVendors.length > 0) {
            console.log('\n--- MISSING VENDORS FOR SETUP (Copy/Paste) ---');
            const allUniqueVendors = Array.from(new Set(illegalVendors.map(x => x.value))).sort();
            allUniqueVendors.forEach(v => console.log(v));
        }
    }

    // Output compliance errors if any
    if (global.complianceErrors && global.complianceErrors.length > 0) {
        console.error('\n--- COMPLIANCE ERRORS ---');
        global.complianceErrors.forEach(err => console.error(err));
    }

    if (duplicateCategories.length) {
        console.log('\n--- CATEGORY REPORT TYPE CONFLICTS ---');
        console.log('[!] The following categories have CONFLICTING Report types in your Setup sheet.');
        console.log('[!] Multiple rows with the same category but DIFFERENT subcategories is VALID.');
        console.log('[!] But all rows for a category must have the SAME Report type (P&L or BS).\n');
        duplicateCategories.forEach(d => {
            console.log(`  "${d.name}" (Row ${d.row}): Trying to set Report="${d.newReport}", but earlier row set it to "${d.existingReport}"`);
        });
        console.log('\nTo fix:');
        console.log('1. Open the Excel file and go to the Setup tab');
        console.log('2. Find all rows for the conflicting category');
        console.log('3. Ensure ALL rows have the SAME value in the Report column (either all "P&L" or all "BS")');
    }

    // Check for vendors/customers used in both P&L and BS
    const mixedVendors = [];
    const mixedCustomers = [];

    vendorReportUsage.forEach((reportMap, vendor) => {
        if (reportMap.size > 1) {
            mixedVendors.push({ name: vendor, reportMap });
        }
    });

    customerReportUsage.forEach((reportMap, customer) => {
        if (reportMap.size > 1) {
            mixedCustomers.push({ name: customer, reportMap });
        }
    });

    if (mixedVendors.length > 0 || mixedCustomers.length > 0) {
        console.log('\n--- VENDOR/CUSTOMER REPORT TYPE CONFLICTS ---');
        console.log('[!] The following vendors/customers are used in BOTH P&L and Balance Sheet categories.');
        console.log('[!] This usually indicates a data entry error.');
        console.log('[!] Vendors/customers should typically only appear in one report type.\n');

        if (mixedVendors.length > 0) {
            console.log('Vendors with mixed usage:');
            mixedVendors.forEach(v => {
                const reportTypes = Array.from(v.reportMap.keys());
                console.log(`\n  "${v.name}" appears in: ${reportTypes.join(' and ')}`);
                // Show example rows for each report type
                v.reportMap.forEach((examples, reportType) => {
                    console.log(`    ${reportType} examples:`);
                    examples.forEach(ex => {
                        console.log(`      - [${ex.date}] ${ex.sheet} Row ${ex.row}: Category="${ex.category}"`);
                    });
                });
            });
        }

        if (mixedCustomers.length > 0) {
            if (mixedVendors.length > 0) console.log('');
            console.log('Customers with mixed usage:');
            mixedCustomers.forEach(c => {
                const reportTypes = Array.from(c.reportMap.keys());
                console.log(`\n  "${c.name}" appears in: ${reportTypes.join(' and ')}`);
                // Show example rows for each report type
                c.reportMap.forEach((examples, reportType) => {
                    console.log(`    ${reportType} examples:`);
                    examples.forEach(ex => {
                        console.log(`      - [${ex.date}] ${ex.sheet} Row ${ex.row}: Category="${ex.category}"`);
                    });
                });
            });
        }

        console.log('\nTo fix:');
        console.log('1. Review your transactions for these vendors/customers');
        console.log('2. Ensure they are categorized consistently (all P&L or all BS)');
        console.log('3. If legitimately needed in both, consider using different vendor/customer names');
    }

    if (offsetWarnings.length) {
        console.log('\n--- OFFSET WARNINGS ---');
        offsetWarnings.forEach(w => console.log(`[!] Sheet "${w.sheet}" Row ${w.row} looks like a header (Found: ${w.matches.join(', ')}). Adjust Setup tab offset.`));
    }

    if (!saveFlag) {
        console.log('\n(Run with --save to update the Excel file)');
        return;
    }

    // --- 6. Summary Sheet Update ---
    if (!summarySheet) {
        summarySheet = workbook.addWorksheet('Summary');
    } else {
        // Clear existing content to avoid breaking workbook references/Tables
        summarySheet.eachRow((row, r) => {
            row.eachCell(cell => { cell.value = null; cell.style = {}; });
        });
    }

    // Explicitly disable worksheet-level autofilter to avoid conflicts with Table-level filters
    // summarySheet.autoFilter = null; // Removed to prevent corruption if Table exists

    summarySheet.getCell('A1').value = `Financial Summary (${new Date().toLocaleString()})`;
    summarySheet.getCell('A1').font = { size: 14, bold: true };

    let summaryRow = 3;
    summarySheet.getCell(`A${summaryRow}`).value = 'Profit & Loss';
    summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
    reports.pl.forEach(r => { summarySheet.getCell(`A${summaryRow}`).value = r.label; summarySheet.getCell(`B${summaryRow}`).value = r.value; summaryRow++; });
    summarySheet.getCell(`A${summaryRow}`).value = 'NET INCOME'; summarySheet.getCell(`B${summaryRow}`).value = netIncome;
    summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow += 3;

    summarySheet.getCell(`A${summaryRow}`).value = 'Balance Sheet';
    summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
    reports.bs.forEach(r => { summarySheet.getCell(`A${summaryRow}`).value = r.label; summarySheet.getCell(`B${summaryRow}`).value = r.value; summaryRow++; });

    if (hasIssues) {
        summaryRow += 3;
        summarySheet.getCell(`A${summaryRow}`).value = 'Data Integrity Check';
        summarySheet.getCell(`A${summaryRow}`).font = { bold: true, color: { argb: 'FFFF0000' } }; summaryRow++;
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet)
        ]);
        issueSheetsFound.forEach(s => {
            summarySheet.getCell(`A${summaryRow}`).value = `Tab: ${s.toUpperCase()}`;
            summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
            const uncat = uncategorizedDetails.filter(x => x.sheet === s).length;
            if (uncat) { summarySheet.getCell(`A${summaryRow}`).value = '  Uncategorized Rows'; summarySheet.getCell(`B${summaryRow}`).value = uncat; summaryRow++; }
            const cats = Array.from(new Set(illegalCategories.filter(x => x.sheet === s).map(x => x.value))).join(', ');
            if (cats) { summarySheet.getCell(`A${summaryRow}`).value = '  Illegal Categories'; summarySheet.getCell(`B${summaryRow}`).value = cats; summaryRow++; }
            const vends = Array.from(new Set(illegalVendors.filter(x => x.sheet === s).map(x => x.value))).join(', ');
            if (vends) { summarySheet.getCell(`A${summaryRow}`).value = '  Unknown Vendors'; summarySheet.getCell(`B${summaryRow}`).value = vends; summaryRow++; }
            summaryRow++;
        });
    }

    // Backup before save
    try {
        const d = new Date();
        const pad = (n) => n.toString().padStart(2, '0');
        const timestamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
        const ext = path.extname(filename);
        const base = path.basename(filename, ext);
        const backupPath = path.join(path.dirname(filename), `${base}_backup_${timestamp}${ext}`);
        fs.copyFileSync(filename, backupPath);
        console.log(`\nBackup created: ${backupPath}`);
    } catch (e) {
        console.error(`Warning: Failed to create backup: ${e.message}`);
    }

    try {
        await workbook.xlsx.writeFile(filename);
        if (showChecker || saveFlag) console.log(`\nSuccessfully updated financials in ${filename}`);
    } catch (e) {
        console.error('Error saving file:', e.message);
    }
}

updateFinancials().catch(console.error);

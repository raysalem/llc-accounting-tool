const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const util = require('util');

// Store original console just in case
const originalConsole = { log: console.log, warn: console.warn, error: console.error };

// --- Logger Buffer for "Notes" Tab ---
global.globalWarningCount = 0;
const logBuffer = [];
const consoleLogger = {
    log: (...args) => {
        const msg = util.format(...args);
        originalConsole.log(msg); // Print to terminal using SAFE original console
        logBuffer.push(msg); // Store for Excel
    },
    error: (...args) => {
        const msg = util.format(...args);
        originalConsole.error(msg);
        logBuffer.push('[ERROR] ' + msg);
        global.globalWarningCount = (global.globalWarningCount || 0) + 1;
    },
    warn: (...args) => {
        const msg = util.format(...args);
        originalConsole.warn(msg);
        logBuffer.push('[WARN] ' + msg);
        global.globalWarningCount = (global.globalWarningCount || 0) + 1;
    }
};

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
    const show1099All = args.includes('--1099');
    const show1099NEC = args.includes('--1099=NEC') || args.includes('--1099-nec');
    const show1099INT = args.includes('--1099=INT');
    const show1099 = show1099All || show1099NEC || show1099INT;
    const ignoreVendors = args.includes('--ignore-vendors');

    // Parse --vendor-file <path>
    const vendorFileIndex = args.indexOf('--vendor-file');
    const customVendorFile = vendorFileIndex !== -1 && args[vendorFileIndex + 1] ? args[vendorFileIndex + 1] : null;

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
  --1099          (Optional) Generate both 1099-NEC and 1099-INT reports.
  --1099=NEC      (Optional) Generate only 1099-NEC reports.
  --1099=INT      (Optional) Generate only 1099-INT reports.
  --ignore-vendors (Optional) Skip loading external "vendor.xlsx" or "vendor.csv" files.
  --vendor-file [path] (Optional) Specify a custom path to a "vendor.xlsx" or "vendor.csv" file.
  --details "Cat" (Optional) List all transactions for a specific Category (e.g., --details "Office Supplies").

Example:
  node report.js "My_Books_2025.xlsx" --pl --checker --save
        `);
        return;
    }

    const knownFlags = [
        '--save', '--pl', '--bs', '--vendor', '--vendor-sub', '--customer', '--customer-sub', '--pl-sub', '--bs-sub', '--checker', '--debug', '--details', '--help', '--1099', '--1099-nec', '--1099=NEC', '--1099=INT', '--ignore-vendors', '--vendor-file'
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
    let originalInputPath = filename; // Store original input for path resolution

    // Resolve shortcut if needed
    if (fs.existsSync(filename)) {
        console.log(`LLC Accounting Tool v${require('./package.json').version}`);
        const resolved = resolveShortcut(filename);
        if (resolved !== filename) {
            console.log(`Resolved shortcut '${filename}' -> '${resolved}'`);
            filename = resolved;
        }
    }

    // Override console for capturing output EARLY to capture setup warnings
    console.log = consoleLogger.log;
    console.warn = consoleLogger.warn;
    console.error = consoleLogger.error;

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
    // Try 'General Ledger' first, then 'Ledger', then case-insensitive scan
    let ledgerSheet = workbook.getWorksheet('General Ledger') || workbook.getWorksheet('Ledger');
    if (!ledgerSheet) {
        ledgerSheet = workbook.worksheets.find(s => {
            const n = s.name.trim().toLowerCase();
            return n === 'general ledger' || n === 'ledger';
        });
    }

    let summarySheet = workbook.getWorksheet('Summary');

    if (!setupSheet || !ledgerSheet) {
        console.error('\n[ERROR] Mandatory sheets missing from workbook:');
        if (!setupSheet) console.error(' - "Setup" sheet is missing.');
        if (!ledgerSheet) console.error(' - "General Ledger" (or "Ledger") sheet is missing.');
        return;
    }

    // --- State ---
    const validCategories = new Set(); // Stores lowercase for validation
    const validVendors = new Map();    // Maps lower -> Display Name
    const vendor1099Map = new Map();   // Maps lower -> { type: 'NEC'|'INT', req: 'YES'|'NO'|'' }
    const vendorDetailsMap = new Map(); // Maps lower -> strict Object of details

    // Helper to safely parse balances (Start/End)
    // Returns number only if valid number found. Returns null if empty string, null/undefined, or "NA"
    const parseBalance = (val) => {
        if (val === null || val === undefined) return null;
        let s = val.toString().trim();
        if (s === '' || s.toLowerCase() === 'na' || s.toLowerCase() === 'n/a' || s.toLowerCase() === 'nan') return null;
        const n = parseFloat(s);
        return isNaN(n) ? null : n;
    };
    let payerInfo = {}; // Payer/Company Info Map
    const validCustomers = new Map();  // Maps lower -> Display Name
    const uniqueCategories = new Map(); // Maps lower -> { report, accountType, displayName }
    const validSubCategories = new Set(); // Set of all valid subcategories from Setup
    const sheetConfigs = [];
    const processedSheetTotals = {}; // Track actual processed totals per sheet (post-flip)
    const walletCheckResults = []; // Store wallet validation data

    const catStats = {};
    const vendorStats = {};
    const vendor1099Stats = { NEC: {}, INT: {} };
    const customerStats = {};
    let bankTotal = 0;
    let ccTotal = 0;
    let uncategorizedBank = 0;
    let hasErrors = false; // Track global error state for final check
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

    const NORM_VEND = (s) => (s || '').toString().toLowerCase().replace(/\s+/g, ' ').trim();

    // --- Helper ---
    function getVal(cell) {
        if (!cell) return '';
        let v = cell.value;
        if (v && typeof v === 'object') {
            if (v.result !== undefined) v = v.result;
            else if (v.richText) return v.richText.map(t => t.text).join('').trim();
            else if (v.text && v.hyperlink) return v.text.trim(); // Handle Hyperlink object
        }
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
        const lookups = [
            'category', 'subcategory', 'vendors', 'vendor', 'sheetname', 'sheetnameconfig', 'report', 'linkcategory', 'linkcat',
            'name', 'fullname', 'firstname', 'lastname', 'address', 'city', 'state', 'zip', 'tin', 'ssn', 'ein', 'taxid'
        ];

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
    // The user renamed 'Category' to 'Link Asset' for the asset account linkage
    const colSheetCat = getCol('linkasset') || getCol('linkcategory') || getCol('linkcat') || getCol('category', 'last');
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
            const lowerV = NORM_VEND(vRaw);
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
            const colC = headerCols['linkasset'] || headerCols['linkcategory'] || headerCols['linkcat'] || headerCols['category'] || 0;
            const colF = headerCols['flippolarityyesno'] || headerCols['flippolarity'] || headerCols['flip'] || 3;
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
        if (showChecker) console.log('[DEBUG] SheetInfo Table missing. Scanning for "Sheet Name" header block...');

        // Scan for header row containing "Sheet Name"
        let blockHeaderRow = -1;
        let blockMap = {};

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

            if (showChecker) console.log(`[Setup] Found SheetInfo headers on Row ${blockHeaderRow}. Link Col: ${colL}, Start Col: ${colStart}, End Col: ${colEnd}`);

            setupSheet.eachRow((row, r) => {
                if (r <= blockHeaderRow) return;
                const name = colN ? getVal(row.getCell(colN)) : null;
                if (name && name.toString().trim()) {
                    configRows.push({
                        name: name,
                        cat: colL ? getVal(row.getCell(colL)) : '',
                        shortName: colS ? getVal(row.getCell(colS)) : '',
                        // Type is now the golden truth for polarity
                        type: colT ? getVal(row.getCell(colT)) : '',
                        offset: colO ? getVal(row.getCell(colO)) : '',
                        startBalance: colStart ? (parseBalance(getVal(row.getCell(colStart))) || 0) : 0,
                        endBalance: colEnd ? parseBalance(getVal(row.getCell(colEnd))) : null
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

            if (cType === 'expense') {
                doFlip = true;
            } else if (cType === 'income') {
                doFlip = false;
            } else if (cType === 'ledger') {
                doFlip = false;
                doLink = false; // Ledger never links automatically
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
                    for (const [catRaw, catData] of uniqueCategories.entries()) {
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

            if (showChecker) {
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
                hasErrors = true; // Mark global error
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

    if (showChecker) {
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
        sheetConfigs.push({ name: 'Bank Transactions', type: 'Bank', flip: false, offset: 1, linkedAccount: null });
        sheetConfigs.push({ name: 'Credit Card Transactions', type: 'CC', flip: true, offset: 1, linkedAccount: null });
    }

    // --- 1.5 Load External Vendors (Override/Enrich Setup) ---
    async function loadExternalVendors() {
        // Search locations: 
        // 1. Current Working Directory (Run Area)
        // 2. Directory of the INPUT file (Shortcut location)
        // 3. Directory of the TARGET file (Actual Excel file location)

        const cwd = process.cwd();
        const targetDir = path.dirname(filename);
        const inputDir = path.dirname(originalInputPath);

        const dirsToCheck = new Set([cwd, inputDir, targetDir]); // Use Set to deduplicate
        const checkPaths = [];

        if (customVendorFile) {
            checkPaths.push(customVendorFile);
        } else {
            dirsToCheck.forEach(dir => {
                checkPaths.push(path.join(dir, 'vendor.xlsx'));
                checkPaths.push(path.join(dir, 'vendor.csv'));
            });
        }

        let loadedPath = null;
        for (const fPath of checkPaths) {
            if (fs.existsSync(fPath)) {
                loadedPath = fPath;

                const vWb = new ExcelJS.Workbook();
                if (fPath.endsWith('.csv')) await vWb.csv.readFile(fPath);
                else await vWb.xlsx.readFile(fPath);

                const vSheet = vWb.worksheets[0];
                if (!vSheet) continue;

                // Simple Header Search
                const { map: vHeaders, headerRow: vHeaderRow } = getHeaderMap(vSheet);
                // Header mapping helper for single indices
                const getVCol = (k) => {
                    const indices = vHeaders.get(k);
                    return (indices && indices.length > 0) ? indices[0] : null;
                };

                const cName = getVCol('name') || getVCol('vendor') || getVCol('fullname') || getVCol('recipientname') || getVCol('payee');
                const cBiz = getVCol('businessname') || getVCol('business') || getVCol('company');
                const cSSN = getVCol('ssn') || getVCol('taxid') || getVCol('tin') || getVCol('ein') || getVCol('ssnein') || getVCol('taxidnumber');
                const cAddr = getVCol('address') || getVCol('street') || getVCol('streetaddress') || getVCol('addr') || getVCol('address1') || getVCol('mailingaddress');
                const cEmail = getVCol('email') || getVCol('emailaddress');
                const cPhone = getVCol('phone') || getVCol('phonenumber') || getVCol('mobile') || getVCol('phone#');
                const cCountry = getVCol('country');
                const cEntityType = getVCol('type') || getVCol('entitytype') || getVCol('recipienttype');
                const cTINType = getVCol('tintype') || getVCol('taxidtype');
                const cREQ = getVCol('1099') || getVCol('1099required') || getVCol('required');

                const cFirstName = getVCol('firstname') || getVCol('first');
                const cLastName = getVCol('lastname') || getVCol('last');
                const cCity = getVCol('city') || getVCol('town') || getVCol('citytown');
                const cState = getVCol('state') || getVCol('province');
                const cZip = getVCol('zipcode') || getVCol('zip') || getVCol('postalcode');

                // Headers found

                vSheet.eachRow((row, r) => {
                    if (r <= vHeaderRow) return;

                    let name = cName ? getVal(row.getCell(cName)) : '';
                    let biz = cBiz ? getVal(row.getCell(cBiz)) : '';
                    const firstName = cFirstName ? getVal(row.getCell(cFirstName)) : '';
                    const lastNameVal = cLastName ? getVal(row.getCell(cLastName)) : '';

                    // Construct Fallback Name if 'Name' column is empty
                    if (!name.trim()) {
                        if (firstName && lastNameVal) name = `${firstName} ${lastNameVal}`;
                        else if (lastNameVal) name = lastNameVal;
                        else if (firstName) name = firstName;
                    }

                    const ssn = (cSSN ? getVal(row.getCell(cSSN)) : '').toString().trim();
                    const addr = (cAddr ? getVal(row.getCell(cAddr)) : '').toString().trim();
                    const city = (cCity ? getVal(row.getCell(cCity)) : '').toString().trim();
                    const state = (cState ? getVal(row.getCell(cState)) : '').toString().trim();
                    const zip = (cZip ? getVal(row.getCell(cZip)) : '').toString().trim();


                    // Key candidate list (multi-mapping)
                    const keys = new Set();
                    if (name) keys.add(NORM_VEND(name));
                    if (biz) keys.add(NORM_VEND(biz));
                    if (firstName && lastNameVal) keys.add(NORM_VEND(`${firstName} ${lastNameVal}`));

                    if (keys.size === 0) return;

                    const details = {
                        name: name || (firstName && lastNameVal ? `${firstName} ${lastNameVal}` : (biz || '')),
                        firstName: firstName || '',
                        lastName: lastNameVal || (name && !biz ? name : ''),
                        business: biz || '',
                        ssn: ssn || '',
                        address: addr || '',
                        city: city || '',
                        state: state || '',
                        zip: zip || '',
                        email: (cEmail ? getVal(row.getCell(cEmail)) : '') || '',
                        phone: (cPhone ? getVal(row.getCell(cPhone)) : '') || '',
                        country: (cCountry ? getVal(row.getCell(cCountry)) : '') || '',
                        entityType: (cEntityType ? getVal(row.getCell(cEntityType)) : '') || '',
                        tinType: (cTINType ? getVal(row.getCell(cTINType)) : '') || ''
                    };

                    // Mark as 1099 required if column says so
                    if (cREQ) {
                        const val = getVal(row.getCell(cREQ)).toString().toLowerCase();
                        if (val === 'yes' || val === 'nec' || val === 'misc' || val === 'int') {
                            details.req = 'YES';
                            details.type = (val === 'yes') ? 'NEC' : val.toUpperCase();
                        }
                    }

                    keys.forEach(k => {
                        const existing = vendorDetailsMap.get(k) || {};
                        vendorDetailsMap.set(k, { ...existing, ...details });
                        if (!validVendors.has(k)) validVendors.set(k, details.name || details.business || k);
                    });
                });
                break; // Stop after first successful file match
            }
        }
        if (!loadedPath) {
            console.warn(`\n[!] WARNING: No "vendor.xlsx" or "vendor.csv" found in search paths.`);
            console.warn(`    Looked in: ${Array.from(dirsToCheck).join(', ')}`);
            console.warn(`    1099 details will be missing.\n`);
        }
    }
    if (!ignoreVendors) {
        await loadExternalVendors();
    }

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
        if (config.name.toLowerCase() === 'ledger' || config.name.toLowerCase() === 'general ledger' || config.type.toLowerCase() === 'ledger') {
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
            console.log(`  Linked Account: ${config.linkedAccount || 'NONE (Error if not Ledger)'}`);
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

                // Track global processed total for this sheet
                processedSheetTotals[config.name] = (processedSheetTotals[config.name] || 0) + amount;

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
                    const vLower = NORM_VEND(vStr);
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
                    const cLower = NORM_VEND(cStr);
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

        // Check End Balance if configured (strict check against null)
        if (config.endBalance !== null) {
            // For Expense sheets (Credit Cards), the 'sheetTotal' is positive (Expenses).
            // But for the Account Balance, spending should be negative (increasing liability/reducing cash).
            // So we invert the change for the validation math if it's an Expense sheet.
            const validationChange = (config.type.toLowerCase() === 'expense') ? -sheetTotal : sheetTotal;

            const calculatedEnd = (config.startBalance || 0) + validationChange;
            const diff = Math.abs(calculatedEnd - config.endBalance);
            walletCheckResults.push({
                sheet: config.shortName,
                start: config.startBalance || 0,
                change: validationChange,
                calcEnd: calculatedEnd,
                expected: config.endBalance,
                diff: diff,
                passed: diff < 0.01
            });
        }

        // JIT Linkage Fallback: If no linked account from Step 1, try again using raw cat column
        // This catches cases where categories might have fully loaded or matched differently
        let effectiveLink = config.linkedAccount;
        if (!effectiveLink && config.cat && config.type !== 'ledger') {
            const jitMatch = uniqueCategories.get(config.cat.toLowerCase());
            if (jitMatch) effectiveLink = jitMatch.displayName;
        }

        if (effectiveLink) {
            const linkName = effectiveLink;
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

            let linkageMsg = `[Linkage Logic] Linked "${config.name}" to ${isAsset ? 'Asset' : 'Liability'} "${linkName}".`;
            if (config.startBalance && Math.abs(config.startBalance) > 0.001) {
                catStats[linkName].total += config.startBalance;
                linkageMsg = `[Linkage Logic] Applied Starting Balance (${config.startBalance.toFixed(2)}) and ${config.shortName} Total (${sheetTotal.toFixed(2)}) to ${isAsset ? 'Asset' : 'Liability'} "${linkName}".`;
            } else {
                linkageMsg = `[Linkage Logic] Applied ${config.shortName} Total (${sheetTotal.toFixed(2)}) to ${isAsset ? 'Asset' : 'Liability'} "${linkName}".`;
            }

            // Track sheet-level contribution to BS account
            if (!catStats[linkName].sheets[config.name]) {
                catStats[linkName].sheets[config.name] = { add: 0, sub: 0, total: 0 };
            }
            const sStat = catStats[linkName].sheets[config.name];
            if (sheetTotal >= 0) sStat.add += sheetTotal; else sStat.sub += sheetTotal;

            if (isAsset) sStat.total += sheetTotal; else sStat.total -= sheetTotal;

            console.log(`${linkageMsg} Balance: ${previous.toFixed(2)} -> ${catStats[linkName].total.toFixed(2)}`);
        } else if (config.type !== 'ledger') {
            // Only warn if it's not a Ledger (Ledgers are manual)
            if (showDebug) console.log(`[Linkage Logic] Sheet "${config.name}" (Type: ${config.type}) has NO LINKED ACCOUNT. Total (${sheetTotal.toFixed(2)}) NOT applied to any Balance Sheet asset.`);
        }
    }

    // --- 3. Process Ledger ---
    // Find Ledger configuration from Setup
    // Find Ledger configuration from Setup, or fallback to default
    let ledgerConfig = sheetConfigs.find(c => c.name.toLowerCase() === 'general ledger' || c.name.toLowerCase() === 'ledger');

    if (!ledgerConfig) {
        if (showChecker) console.log(`[Info] No explicit "Ledger" config found in Setup. Using default configuration with auto-detected header.`);

        // Auto-detect header row
        let detectedOffset = 3;
        ledgerSheet.eachRow((row, r) => {
            if (r > 10) return; // Only scan first 10 rows
            const values = row.values.map(v => (v ? v.toString().toLowerCase() : ''));
            const hasDate = values.some(v => v.includes('date'));
            const hasCat = values.some(v => v.includes('category'));
            if (hasDate && hasCat) {
                detectedOffset = r;
            }
        });

        ledgerConfig = {
            name: ledgerSheet.name,
            type: 'ledger',
            flip: false,
            offset: detectedOffset,
            startBalance: 0,
            endBalance: null,
            linkedAccount: null
        };
        sheetConfigs.push(ledgerConfig);
    }

    const ledgerHeaderRow = ledgerConfig.offset || 3;

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
    let ledgerValidationTotal = 0; // Raw Dr - Cr, must be 0
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

            // Accumulate validation total (Dr should equal Cr, so Dr - Cr should be 0 across all rows)
            ledgerValidationTotal += (dr - cr);

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

            // Vendor Validation
            if (vendorVal) {
                const vStr = vendorVal.toString().trim();
                const vLower = NORM_VEND(vStr);
                // Using 'General Ledger' or dynamic name? ledgerConfig.name is better but 'Ledger' is hardcoded here in context
                if (!validVendors.has(vLower)) illegalVendors.push({ value: vStr, sheet: 'Ledger', row: r, date: (rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A')) });
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
                    processedSheetTotals[ledgerConfig.name] = (processedSheetTotals[ledgerConfig.name] || 0) + impact;
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
                    const vLower = NORM_VEND(vStr);
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
        const activeNEC = showAll || show1099All || show1099NEC;
        const activeINT = showAll || show1099All || show1099INT;

        if (activeNEC) print1099('NEC', reports.vendors1099NEC, 600);
        if (activeINT) print1099('INT', reports.vendors1099INT, 0);

        // Generate 1099 List if any data found
        const all1099 = [];
        if (activeNEC) all1099.push(...reports.vendors1099NEC.map(x => ({ ...x, form: 'NEC', threshold: 600 })));
        if (activeINT) all1099.push(...reports.vendors1099INT.map(x => ({ ...x, form: 'INT', threshold: 0 })));

        // Filter by threshold & polarity
        const csvRows = [];
        all1099.forEach(r => {
            const isExpense = r.value > 0; // Net Payment
            if (isExpense && r.value >= r.threshold) {
                const d = vendorDetailsMap.get(NORM_VEND(r.label)) || {};
                csvRows.push({
                    ...d,
                    amount: r.value,
                    form: r.form,
                    originalLabel: r.label
                });
            }
        });
        // Store for saveReport
        reports.data1099 = csvRows;

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

            /* 
            if (saveFlag) {
                fs.writeFileSync(csvPath, fileContent.join('\n'));
                console.log(`\n[SUCCESS] Generated 1099 2CSV: ${csvPath}`);
            }
            */

            // Always print to screen
            // 1. Payer Info (You) - Display ONCE
            console.log('\n--- 1099 PAYER INFO (You) ---');
            console.log(`Name:    ${pName}`);
            console.log(`TIN:     ${pTIN}`);
            console.log(`Address: ${pAddr}, ${pCity}, ${pState} ${pZip}`);
            console.log(`Phone:   ${pPhone}`);
            console.log(`Email:   ${pEmail}`);

            // 2. Recipient Info
            console.log('\n--- 1099 RECIPIENT INFO (Vendors) ---');
            console.log(`Recipient`.padEnd(30) + `Business`.padEnd(25) + `TIN`.padEnd(12) + `Amount`.padStart(12) + `  Form`);
            console.log('-'.repeat(85));
            csvRows.forEach(r => {
                const rName = (r.name || 'Unknown').substring(0, 29);
                const rBiz = (r.business || '').substring(0, 24);

                // STRICT VALIDATION for Console Output too
                const missingFields = [];
                // Check r (which is merged details)
                if (!r.ssn) missingFields.push('Tax ID (SSN/EIN)');
                if (!r.address) missingFields.push('Address');
                if (!r.city) missingFields.push('City');
                if (!r.state) missingFields.push('State');
                if (!r.zip) missingFields.push('Zip');

                const statusLine = `${rName.padEnd(30)}${rBiz.padEnd(25)}${(r.ssn || '').padEnd(12)}${r.amount.toFixed(2).padStart(12)}  ${r.form}`;
                console.log(statusLine);

                if (missingFields.length > 0) {
                    const vendKey = NORM_VEND(r.originalLabel || r.name || r.business || '');
                    const inDb = vendorDetailsMap.has(vendKey);

                    if (!inDb) {
                        console.error(`  [!] ERROR: Not found in "vendor.xlsx".`);
                        // Fuzzy check for suggestions
                        const suggestions = [];
                        for (const dbKey of vendorDetailsMap.keys()) {
                            if (dbKey.includes(vendKey) || vendKey.includes(dbKey)) suggestions.push(vendorDetailsMap.get(dbKey).name);
                        }
                        if (suggestions.length > 0) {
                            console.error(`      Did you mean: ${suggestions.join(' or ')}?`);
                        }
                    } else {
                        console.error(`  [!] ERROR: Incomplete Data! Missing: ${missingFields.join(', ')}`);
                    }
                    console.error(`      update "vendor.xlsx" to fix.`);
                    hasErrors = true;
                }
            });
            if (hasErrors) {
                throw new Error(`1099 Validation Failed: One or more vendors have incomplete details. Update "vendor.xlsx" and run again.`);
            }
        }
    }
    // Customer report (already handles sub)
    if (showAll || showCustomer || showCustomerSub) {
        if (showCustomerSub) {
            printDetailedTable('CUSTOMER INCOME (Detailed)', reports.customers, reportSheetList, sheetNameMap, "Customer");
        } else {
            // Already shown in main "Vendor Spending" block if showVendor is on, 
            // but if showVendorSub is ON and showVendor is OFF (edge case), we show detailed.
            // If both, we might duplicate? 
            // The main vendor block shows 1099 info. This shows stats.
            // Let's assume if showVendor is on, we don't need this basic block unless sub is requested.
            // But the original code printed it.
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

    if (Math.abs(ledgerValidationTotal) > 0.01) {
        console.warn(`\n[!] CRITICAL WARNING: Ledger does not sum to zero!`);
        console.warn(`    Net Mismatch (Dr - Cr): ${ledgerValidationTotal.toFixed(2)}`);
        console.warn(`    Double-entry accounting requires Debits to equal Credits. Please check your Ledger entries.`);
        hasErrors = true;
    }

    // --- Wallet Reconciliation Report ---
    if (walletCheckResults.length > 0) {
        console.log('\n--- WALLET RECONCILIATION ---');
        console.log(`Sheet Name`.padEnd(25) + `Start`.padStart(12) + `Change`.padStart(12) + `Calc End`.padStart(12) + `Exp End`.padStart(12) + `Diff`.padStart(12));
        console.log('-'.repeat(85));
        walletCheckResults.forEach(w => {
            if (!w.passed) hasErrors = true;
            const status = w.passed ? '' : ' [MISMATCH]';
            console.log(
                `${w.sheet.padEnd(25)}` +
                `${w.start.toFixed(2).padStart(12)}` +
                `${w.change.toFixed(2).padStart(12)}` +
                `${w.calcEnd.toFixed(2).padStart(12)}` +
                `${w.expected.toFixed(2).padStart(12)}` +
                `${w.diff.toFixed(2).padStart(12)}` +
                `${status}`
            );
        });
    }

    // --- Global Integrity Check: Total Tie ---
    // The sum of all net flows across all transaction sheets (after polarity)
    // must equal (Net Income) + (Sum of Balance Sheet changes).
    let globalFlowTotal = 0;
    const sheetFlows = [];
    sheetConfigs.forEach(conf => {
        // Use the tracked processed total, not the re-summed category stats
        const sheetSum = processedSheetTotals[conf.name] || 0;

        globalFlowTotal += sheetSum;
        sheetFlows.push({ name: conf.shortName, flow: sheetSum });
    });

    // Net Income + Balance Sheet Changes (EXCLUDING Linked Accounts)
    // We exclude linked accounts because they represent the "Source" (Global Flow), 
    // while "Other BS" represents the "Destination" (e.g. Transfers, Loan Paydown).
    // Formula: GlobalFlow (Source Delta) = NetIncome + OtherBS (Destination Delta)
    const linkedNames = new Set(sheetConfigs.map(c => c.linkedAccount).filter(Boolean));
    const otherBS = reports.bs.filter(r => !linkedNames.has(r.label));
    const bsChangeTotal = otherBS.reduce((a, b) => a + b.value, 0);
    const accountingTotal = netIncome + bsChangeTotal;

    if (showChecker) {
        console.log(`\n--- GLOBAL INTEGRITY CHECK ---`);
        console.log(`Sum of Net Income + BS Changes:      ${accountingTotal.toFixed(2)}`);
        console.log(`Total Sheet Flow (Bank+CC+Ledger):   ${globalFlowTotal.toFixed(2)}`);

        const diff = Math.abs(accountingTotal - globalFlowTotal);
        if (diff < 0.05) {
            console.log(`Total Tie Status:                    [OK] (Difference: ${diff.toFixed(2)})`);
        } else {
            console.warn(`Total Tie Status:                    [FAIL] Mismatch of ${diff.toFixed(2)}`);
            hasErrors = true;
        }
    }

    // --- Final Status ---
    if (showChecker) {
        if (!hasErrors) {
            console.log(`\n✅ [ALL SYSTEMS GO] Financial reports are internally consistent and wallets reconcile.`);
        } else {
            console.log(`\n❌ [CHECKS FAILED] Please review warnings above.`);
        }
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
            console.error('\n[FATAL ERROR] Unknown Vendors Detected!');
            console.error('The following vendors are not in your Setup sheet override list or vendor.xlsx:');
            const allUniqueVendors = Array.from(new Set(illegalVendors.map(x => x.value))).sort();
            allUniqueVendors.forEach(v => console.error(` - "${v}"`));

            console.error('\nPlease add these vendors to "Setup" or "vendor.xlsx" to proceed.');
            console.error('Script will now EXIT to prevent data corruption/incomplete reports.');
            process.exit(1);
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

    if (saveFlag) {
        await saveReport(filename, reports, logBuffer, {
            showPL, showBS, showVendor, showVendorSub, showCustomer, showCustomerSub, show1099, showChecker,
            showPLSub, showBSSub // Passing the sub flags
        }, vendorDetailsMap, payerInfo);
    } else {
        // Original behavior: If NO save flag, maybe we just updated the file in-place? 
        // The previous code did a backup and overwrite.
        // User request specifically asks for "--save, create a new file".
        // If --save is NOT passed, we should probably NOT touch the file or just do the in-place update if that was partial logic.
        // But previously it was "Run to update financials". 
        // Let's keep the existing in-place update IF --save is NOT passed but the user ran default mode?
        // Actually, user says "for the --save, creat a new file...". 
        // Implicitly, if --save is OFF, strict read-only report mode is safer given the new direction.
        // But to avoid breaking valid "update" workflows, I'll allow in-place update ONLY if requested?
        // Let's assume default is Read-Only unless --save is passed for the new report.
        if (showChecker) console.log('\n[Info] Run with --save to generate full Excel report.');
    }

    if (global.globalWarningCount > 0) {
        originalConsole.error(`\n[BATCH STOP] Process exited with ${global.globalWarningCount} warnings/errors.`);
        process.exit(1);
    }
}

// --- Report Generation Helper ---
async function saveReport(originalFilename, reports, logs, flags, vendorDetails, payerInfo) {
    const dir = path.dirname(originalFilename);
    const base = path.basename(originalFilename, path.extname(originalFilename));
    const newFilename = path.join(dir, `report_${base}.xlsx`);

    // We use originalConsole to avoid cluttering the buffer right at the end
    originalConsole.log(`\n[Saving] Generating report file: ${newFilename} ...`);

    const wb = new ExcelJS.Workbook();

    // Helper to add header row
    const addHeader = (sheet, columns) => {
        sheet.addRow(columns);
        sheet.getRow(1).font = { bold: true };
    };

    // 1. Profit & Loss (Standard + Sub)
    if (flags.showPL) {
        // Standard Tab
        const ws = wb.addWorksheet('Profit & Loss');
        ws.columns = [{ header: 'Category', key: 'cat', width: 35 }, { header: 'Amount', key: 'amt', width: 15 }];
        reports.pl.forEach(r => ws.addRow({ cat: r.label, amt: r.value }));
        const netPL = reports.pl.reduce((acc, r) => acc + r.value, 0);
        const lastRow = ws.addRow({ cat: 'NET INCOME', amt: netPL });
        lastRow.font = { bold: true };
        ws.getColumn(2).numFmt = '#,##0.00';

        // P&L Detailed Tab (if --pl passed, we also include details if --pl-sub OR just by default in Excel? User asked for tabs.)
        // User request: "also tab for bs-sub, pl-sub... PLUS without sub". So we always add detailed tabs if base arg present?
        // Or only if sub arg? User said "if pl and bs-sub...".
        // Let's assume: flags.showPL -> Standard Tab. flags.showPLSub -> Detailed Tab.
        // Wait, user said "also tab for bs-sub, pl-sub... PLUS without sub".
        // This suggests we should produce BOTH tabs if the detailed flag is on.
        // Actually, let's just produce the detailed tab as "Profit & Loss Detailed" if showPLSub is explicitly requested?
        // Or maybe just always include it because it's effectively "Save full report"?
        // Given "plus without sub", I'll generate both tabs logic if showPL is true (Standard) and showPLSub/showPLDetailed?
        // Flag checking logic passed in `flags` object.
        // I will add 'showPLSub' and 'showBSSub' to the flags object being passed.

        if (flags.showPLSub) {
            const wsSub = wb.addWorksheet('P&L Detailed');
            // Gather all sheet names
            const allSheets = new Set();
            reports.pl.forEach(c => Object.keys(c.sheets || {}).forEach(s => allSheets.add(s)));
            const sortedSheets = Array.from(allSheets).sort();

            const header = ['Category', 'Net Total', 'Additions', 'Subtractions'];
            sortedSheets.forEach(s => header.push(`${s} (Net)`));

            wsSub.columns = header.map(h => ({ header: h, width: h === 'Category' ? 35 : 15 }));

            reports.pl.forEach(c => {
                const row = [c.label, c.value, c.add, c.sub];
                sortedSheets.forEach(s => {
                    const sData = (c.sheets && c.sheets[s]) ? c.sheets[s].total : 0;
                    row.push(sData);
                });
                wsSub.addRow(row);
            });

            // Calculate Totals for Detailed View
            if (reports.pl.length > 0) {
                const totalRow = ['NET INCOME'];
                // 1. Net Total
                totalRow.push(reports.pl.reduce((sum, r) => sum + r.value, 0));
                // 2. Additions
                totalRow.push(reports.pl.reduce((sum, r) => sum + (r.add || 0), 0));
                // 3. Subtractions
                totalRow.push(reports.pl.reduce((sum, r) => sum + (r.sub || 0), 0));

                // 4. Per-Sheet Totals
                sortedSheets.forEach(s => {
                    const sheetSum = reports.pl.reduce((sum, r) => {
                        return sum + ((r.sheets && r.sheets[s]) ? r.sheets[s].total : 0);
                    }, 0);
                    totalRow.push(sheetSum);
                });

                const tRow = wsSub.addRow(totalRow);
                tRow.font = { bold: true };
            }

            // Format numbers
            for (let c = 2; c <= header.length; c++) wsSub.getColumn(c).numFmt = '#,##0.00';
        }
    }

    // 2. Balance Sheet (Standard + Sub)
    if (flags.showBS) {
        // Standard Tab
        const ws = wb.addWorksheet('Balance Sheet');
        ws.columns = [{ header: 'Account', key: 'acc', width: 35 }, { header: 'Balance', key: 'bal', width: 15 }];
        reports.bs.forEach(r => ws.addRow({ acc: r.label, bal: r.value }));
        ws.getColumn(2).numFmt = '#,##0.00';

        if (flags.showBSSub) {
            const wsSub = wb.addWorksheet('Balance Sheet Detailed');
            const allSheets = new Set();
            reports.bs.forEach(c => Object.keys(c.sheets || {}).forEach(s => allSheets.add(s)));
            const sortedSheets = Array.from(allSheets).sort();

            const header = ['Account', 'Balance', 'Additions', 'Subtractions'];
            sortedSheets.forEach(s => header.push(`${s} (Net)`));

            wsSub.columns = header.map(h => ({ header: h, width: h === 'Account' ? 35 : 15 }));

            reports.bs.forEach(c => {
                const row = [c.label, c.value, c.add, c.sub];
                sortedSheets.forEach(s => {
                    const sData = (c.sheets && c.sheets[s]) ? c.sheets[s].total : 0;
                    row.push(sData);
                });
                wsSub.addRow(row);
            });
            for (let c = 2; c <= header.length; c++) wsSub.getColumn(c).numFmt = '#,##0.00';
        }
    }

    // 3. Vendor Report
    if (flags.showVendor) {
        const ws = wb.addWorksheet('Vendor Report');
        ws.columns = [{ header: 'Vendor', key: 'v', width: 30 }, { header: 'Net Amount', key: 'a', width: 15 }];
        reports.vendors.forEach(v => ws.addRow({ v: v.label, a: v.value }));
        ws.getColumn(2).numFmt = '#,##0.00';
    }

    // 3b. Vendor Detailed
    if (flags.showVendorSub) {
        const ws = wb.addWorksheet('Vendor Detailed');
        const header = ['Vendor', 'Net Total', 'Additions', 'Subtractions'];
        const allSheets = new Set();
        reports.vendors.forEach(v => Object.keys(v.sheets || {}).forEach(s => allSheets.add(s)));
        const sortedSheets = Array.from(allSheets).sort();
        sortedSheets.forEach(s => header.push(`${s} (Net)`));

        ws.columns = header.map(h => ({ header: h, key: h, width: h === 'Vendor' ? 30 : 15 }));

        reports.vendors.forEach(v => {
            const row = [v.label, v.value, v.add, v.sub];
            sortedSheets.forEach(s => {
                const sData = (v.sheets && v.sheets[s]) ? v.sheets[s].total : 0;
                row.push(sData);
            });
            ws.addRow(row);
        });
        for (let c = 2; c <= header.length; c++) ws.getColumn(c).numFmt = '#,##0.00';
    }

    // 4. Customer Report
    if (flags.showCustomer) {
        const ws = wb.addWorksheet('Customer Report');
        ws.columns = [{ header: 'Customer', key: 'c', width: 30 }, { header: 'Net Amount', key: 'a', width: 15 }];
        reports.customers.forEach(c => ws.addRow({ c: c.label, a: c.value }));
        ws.getColumn(2).numFmt = '#,##0.00';
    }

    // 4b. Customer Detailed
    if (flags.showCustomerSub) {
        const ws = wb.addWorksheet('Customer Detailed');
        const header = ['Customer', 'Net Total', 'Additions', 'Subtractions'];
        const allSheets = new Set();
        reports.customers.forEach(v => Object.keys(v.sheets || {}).forEach(s => allSheets.add(s)));
        const sortedSheets = Array.from(allSheets).sort();
        sortedSheets.forEach(s => header.push(`${s} (Net)`));

        ws.columns = header.map(h => ({ header: h, width: h === 'Customer' ? 30 : 15 }));

        reports.customers.forEach(c => {
            const row = [c.label, c.value, c.add, c.sub];
            sortedSheets.forEach(s => {
                const sData = (c.sheets && c.sheets[s]) ? c.sheets[s].total : 0;
                row.push(sData);
            });
            ws.addRow(row);
        });
        for (let c = 2; c <= header.length; c++) ws.getColumn(c).numFmt = '#,##0.00';
    }

    // 5. 1099 Report
    if (flags.show1099) {
        const ws = wb.addWorksheet('1099 Data');
        ws.columns = [
            { header: 'Last Name / Business', width: 25 },
            { header: 'First Name', width: 20 },
            { header: 'Type', width: 15 }, // Business or Individual
            { header: 'TIN Type', width: 10 },
            { header: 'TIN', width: 15 },
            { header: 'Email', width: 25 },
            { header: 'Phone Number', width: 15 },
            { header: 'Address', width: 30 },
            { header: 'City', width: 15 },
            { header: 'State', width: 5 },
            { header: 'Zip Code', width: 10 },
            { header: 'Country', width: 10 },
            { header: 'Amount', width: 15 } // Configured for 1099
        ];

        // Use the synchronized 1099 data captured earlier
        (reports.data1099 || []).forEach(details => {
            // Determine Entity Type and TIN Type
            let entityType = details.entityType || '';
            if (!entityType) {
                entityType = details.firstName ? 'Individual' : 'Business';
            }

            // If Business Name is present but Last Name is missing, use Business Name
            const finalLastName = details.lastName || details.business || details.name || 'Unknown';
            const finalFirst = details.firstName || '';

            // Guess TIN Type or use explicit
            let tinType = details.tinType || '';
            if (!tinType) {
                tinType = 'EIN';
                const cleanTIN = (details.ssn || '').replace(/[^0-9]/g, '');
                if (cleanTIN.length === 9) {
                    if (entityType.toLowerCase().startsWith('ind')) tinType = 'SSN';
                }
            }

            ws.addRow([
                finalLastName,
                finalFirst,
                entityType,
                tinType,
                details.ssn || '',
                details.email || '',
                details.phone || '',
                details.address || '',
                details.city || '',
                details.state || '',
                details.zip || '',
                details.country || '',
                details.amount
            ]);
        });
        ws.getColumn(13).numFmt = '#,##0.00';
    }

    // Notes / Log Sheet (Always added if checker or debug or save)
    const logWs = wb.addWorksheet('Processing Log');
    logWs.getColumn(1).width = 120;
    logs.forEach(l => logWs.addRow([l.replace(/\x1b\[[0-9;]*m/g, '')])); // Strip ANSI colors

    await wb.xlsx.writeFile(newFilename);
    originalConsole.log(`[Saved] Report saved to: ${newFilename}`);
}

updateFinancials().catch(e => {
    console.error(`\n[CRITICAL ERROR] ${e.message}`);
    if (e.stack && process.argv.includes('--debug')) {
        originalConsole.error(e.stack);
    }
    process.exit(1);
});

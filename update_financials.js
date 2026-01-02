const ExcelJS = require('exceljs');
const fs = require('fs');

async function updateFinancials() {
    const args = process.argv.slice(2);
    const saveFlag = args.includes('--save');
    const showPL = args.includes('--pl');
    const showBS = args.includes('--bs');
    const showVendor = args.includes('--vendor');
    const showCustomer = args.includes('--customer');
    const showPLSub = args.includes('--pl-sub');
    const showChecker = args.includes('--checker');

    // Parse --details <Category>
    const detailsIndex = args.indexOf('--details');
    const targetDetailsCategory = detailsIndex !== -1 && args[detailsIndex + 1] ? args[detailsIndex + 1].toLowerCase().trim() : null;
    const showDetails = !!targetDetailsCategory;

    // Help Menu
    if (args.includes('--help')) {
        console.log(`
Usage: node update_financials.js [filename] [flags]

Description:
  Updates the financial accounting spreadsheet. It reads the Setup, Ledger, and Transaction sheets,
  categorizes transactions, balances the ledger, and generates P&L / Balance Sheet reports in standard Output format.

Arguments:
  [filename]      Path to the Excel file (default: LLC_Accounting_Template.xlsx)

Flags:
  --help          Show this help message.
  --save          Save changes to the Excel file (Summary tab and formatting).
                  (Default behavior is print-only, which does not modify the file).
  --pl            Print the Profit & Loss statement to the console.
  --bs            Print the Balance Sheet to the console.
  --checker       Run the Data Integrity Checker and verify row-by-row categorization issues.
  --pl-sub        (Optional) Print detailed P&L with sub-category breakdowns.
  --vendor        (Optional) Print spending statistics by Vendor.
  --customer      (Optional) Print income statistics by Customer.
  --details "Cat" (Optional) List all transactions for a specific Category (e.g., --details "Office Supplies").

Example:
  node update_financials.js "My_Books_2025.xlsx" --pl --checker --save
        `);
        return;
    }

    const knownFlags = [
        '--save', '--pl', '--bs', '--vendor', '--customer', '--pl-sub', '--checker', '--details', '--help'
    ];

    // Check for unknown arguments
    const unknownArgs = args.filter(a => a.startsWith('--') && !knownFlags.includes(a));
    if (unknownArgs.length > 0) {
        console.error(`Error: Unknown argument(s): ${unknownArgs.join(', ')}`);
        console.error('Run with --help to see available options.');
        process.exit(1);
    }

    const specificFilter = showPL || showBS || showVendor || showCustomer || showPLSub || showChecker || showDetails;
    const showAll = !specificFilter; // Default to showing standard report if no specific filter is set

    let filename = args.find(a => !a.startsWith('--')) || 'LLC_Accounting_Template.xlsx';
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
    const validCustomers = new Map();  // Maps lower -> Display Name
    const uniqueCategories = new Map(); // Maps lower -> { report, accountType, displayName }
    const sheetConfigs = [];

    const catStats = {};
    const vendorStats = {};
    const customerStats = {};
    let bankTotal = 0;
    let ccTotal = 0;
    let uncategorizedBank = 0;
    let uncategorizedCC = 0;

    const illegalCategories = [];
    const illegalVendors = [];
    const illegalCustomers = [];
    const uncategorizedDetails = [];
    const detailsRows = [];
    const offsetWarnings = [];

    // --- Helper ---
    function getVal(cell) {
        if (!cell) return '';
        let v = cell.value;
        if (v && typeof v === 'object') {
            if (v instanceof Date) return v;
            if (v.result !== undefined) {
                if (v.result === null || v.result === undefined) return '';
                return v.result;
            }
            if (v.richText) return v.richText.map(t => t.text).join('').trim();
            if (v.text) return v.text.toString().trim();
            if (v.formula) return '';
            return v.toString().trim();
        }
        return (v === null || v === undefined) ? '' : v.toString().trim();
    }

    function getHeaderMap(sheet, rowIdx = 1) {
        const map = new Map();
        const row = sheet.getRow(rowIdx);
        row.eachCell((cell, colNumber) => {
            const val = getVal(cell).toString().trim().toLowerCase();
            map.set(val, colNumber);
        });
        return map;
    }

    // --- 1. Read Setup (Decoupled Tables) ---
    const setupHeaders = getHeaderMap(setupSheet, 1);
    if (showChecker) console.log('Setup Headers:', Array.from(setupHeaders.keys()));

    // Table 1: Category Info
    const colCategory = setupHeaders.get('category');
    const colSubCategory = setupHeaders.get('sub-category') || setupHeaders.get('subcategory');
    const colType = setupHeaders.get('type');
    const colReport = setupHeaders.get('report');

    // Table 2: Vendors
    const colVendor = setupHeaders.get('vendors') || setupHeaders.get('vendor');

    // Table 3: Customers
    const colCustomer = setupHeaders.get('customers') || setupHeaders.get('customer');

    // Table 4: Sheet Info
    const colSheetName = setupHeaders.get('sheet name (config)') || setupHeaders.get('sheet name');
    const colSheetType = setupHeaders.get('account type') || setupHeaders.get('sheet type');
    const colFlip = setupHeaders.get('flip polarity? (yes/no)') || setupHeaders.get('flip polarity?') || setupHeaders.get('flip');
    const colOffset = setupHeaders.get('header row') || setupHeaders.get('offset');

    if (showChecker) {
        console.log('--- DEBUG MAPPING ---');
        console.log(`Col Category: ${colCategory}`);
        console.log(`Col Sub-Category: ${colSubCategory}`);
        console.log(`Col Sheet Name: ${colSheetName}`);
        console.log(`Col Sheet Type: ${colSheetType}`);
    }

    // --- Pass 1: Read Reference Tables (Categories, Vendors, Customers) ---
    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        // 1. Process Category Table
        const catName = colCategory ? getVal(row.getCell(colCategory)) : null;
        if (catName) {
            const trimmed = catName.toString().trim();
            const lower = trimmed.toLowerCase();
            const typeVal = colType ? getVal(row.getCell(colType)) : '';
            const subCatVal = colSubCategory ? getVal(row.getCell(colSubCategory)) : '';
            const report = colReport ? getVal(row.getCell(colReport)) : '';
            validCategories.add(lower);
            uniqueCategories.set(lower, {
                report,
                accountType: typeVal,
                subCategory: subCatVal,
                displayName: trimmed
            });
        }

        // 2. Process Vendor Table
        const vendor = colVendor ? getVal(row.getCell(colVendor)) : null;
        if (vendor) {
            const vRaw = vendor.toString().trim();
            validVendors.set(vRaw.toLowerCase(), vRaw);
        }

        // 3. Process Customer Table
        const customer = colCustomer ? getVal(row.getCell(colCustomer)) : null;
        if (customer) {
            const cRaw = customer.toString().trim();
            validCustomers.set(cRaw.toLowerCase(), cRaw);
        }
    });

    // --- Pass 2: Read Sheet Configurations & Link ---
    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        // 4. Process Sheet Info Table
        const confSheetName = colSheetName ? getVal(row.getCell(colSheetName)) : null;
        if (confSheetName) {
            const confType = colSheetType ? getVal(row.getCell(colSheetType)) : '';
            const confFlip = colFlip ? getVal(row.getCell(colFlip)) : '';
            const confOffset = colOffset ? getVal(row.getCell(colOffset)) : '';

            if (confSheetName && confType) {
                const cType = confType.toString().trim();
                let link = null;

                // Try to find a linked GL Account based on "Account Type" match
                // We match Sheet.Type (Col J) against Category.SubCategory (Col B) or Category.Type (Col C)
                // or even Category.Name (Col A) for maximum flexibility.
                for (const [catRaw, catData] of uniqueCategories.entries()) {
                    const catSub = catData.subCategory ? catData.subCategory.toLowerCase() : 'N/A';
                    const targetType = cType.toLowerCase();

                    if (showChecker && cType === 'Bank') {
                        // targeted debug
                        console.log(`Checking Cat: "${catData.displayName}" | Sub: "${catSub}" vs Target: "${targetType}"`);
                    }

                    // Check 'Type' (Asset/Liability)
                    if (catData.accountType && catData.accountType.toLowerCase() === targetType) {
                        link = catData.displayName;
                        break;
                    }
                    // Check 'Sub-Category' (Bank/General) - often used for 'Bank'
                    if (catData.subCategory && catData.subCategory.toLowerCase() === targetType) {
                        link = catData.displayName;
                        break;
                    }
                    // Check exact Category Name match
                    if (catData.displayName && catData.displayName.toLowerCase() === targetType) {
                        link = catData.displayName;
                        break;
                    }
                }

                if (showChecker) {
                    console.log(`[Linkage Result] Sheet "${confSheetName}" (Type: "${cType}") -> Linked to: "${link || 'NONE'}"`);
                    if (!link) {
                        console.log('Setup Headers:', JSON.stringify(Array.from(setupHeaders.keys())));
                    }
                }

                sheetConfigs.push({
                    name: confSheetName.toString().trim(),
                    type: cType,
                    flip: !!(confFlip && confFlip.toString().toLowerCase().includes('y')),
                    offset: parseInt(confOffset) || 0,
                    linkedAccount: link
                });
            }
        }
    });

    if (sheetConfigs.length === 0) {
        // Fallback defaults if no config found in Setup
        console.warn('[!] No sheet configurations found in Setup. Using defaults.');
        sheetConfigs.push({ name: 'Bank Transactions', type: 'Bank', flip: false, offset: 1, linkedAccount: null });
        sheetConfigs.push({ name: 'Credit Card Transactions', type: 'CC', flip: true, offset: 1, linkedAccount: null });
    }

    // --- 2. Process Transaction Sheets ---
    // --- Constants for Column Headers ---
    const HEADERS = {
        DATE: ['date', 'txn date', 'transaction date'],
        DESC: ['description', 'desc', 'payee', 'merchant', 'name'],
        AMOUNT: ['amount', 'amt', 'value'],
        CATEGORY: ['category', 'cat', 'account_category'],
        SUBCAT: ['sub-category', 'sub-cat', 'subcategory', 'subcat'],
        VENDOR: ['vendor', 'vend', 'merchant name'],
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

        const tStr = config.type.toLowerCase();
        const isCC = tStr.includes('cc') || tStr.includes('card') || tStr.includes('credit') || tStr.includes('amex');
        const pType = isCC ? 'cc' : 'bank';

        // Dynamic Map detection
        let headerRowIndex = config.offset || 1;

        // Smarter scan: Check rows 1-5 for actual header signature
        for (let r = 1; r <= 5; r++) {
            const rowVals = sheet.getRow(r).values;
            if (Array.isArray(rowVals)) {
                // Check if row contains major headers
                const rowStr = rowVals.map(v => v ? v.toString().toLowerCase() : '').join(' ');
                if (HEAD_MATCH(rowStr)) {
                    headerRowIndex = r;
                    config.offset = r;
                    break;
                }
            }
        }

        function HEAD_MATCH(rowStr) {
            return (rowStr.includes('date') && (rowStr.includes('amount') || rowStr.includes('category')));
        }

        const headerRow = sheet.getRow(headerRowIndex);
        const map = isCC ? { ...ccMapDefault } : { ...bankMapDefault };

        headerRow.eachCell((cell, colNumber) => {
            const val = getVal(cell);

            if (findCol(val, HEADERS.DATE)) map.date = colNumber;
            else if (findCol(val, HEADERS.DESC)) map.desc = colNumber;
            else if (findCol(val, HEADERS.AMOUNT)) map.amount = colNumber;
            else if (findCol(val, HEADERS.SUBCAT)) map.subCat = colNumber; // Check subcat before cat to avoid partial match if 'category' in 'sub-category'
            else if (findCol(val, HEADERS.CATEGORY)) map.category = colNumber;
            else if (findCol(val, HEADERS.VENDOR)) map.vendor = colNumber;
            else if (findCol(val, HEADERS.CUSTOMER)) map.customer = colNumber;
        });

        if (showChecker) {
            console.log(`\nProcessing "${sheet.name}":`);
            console.log(`  Header Row: ${headerRowIndex}`);
            console.log(`  Mapping: ${JSON.stringify(map)}`);
        }

        sheet.eachRow((row, r) => {
            if (r <= config.offset) return;

            const vendorVal = map.vendor ? getVal(row.getCell(map.vendor)) : '';
            const customerVal = map.customer ? getVal(row.getCell(map.customer)) : '';
            const categoryVal = map.category ? getVal(row.getCell(map.category)) : '';
            const subCatVal = map.subCat ? getVal(row.getCell(map.subCat)) : '';
            let amount = map.amount ? getVal(row.getCell(map.amount)) : 0;
            if (typeof amount !== 'number') amount = parseFloat(amount) || 0;

            const rawDate = map.date ? getVal(row.getCell(map.date)) : '';
            const rawDesc = map.desc ? getVal(row.getCell(map.desc)).toString().toLowerCase() : '';

            // Offset check
            if (r === config.offset + 1) {
                const rowValues = row.values.map(v => (v ? v.toString().toLowerCase() : ''));
                if (HEAD_MATCH(rowValues.join(' '))) {
                    offsetWarnings.push({ sheet: sheet.name, row: r, matches: ['Header Signature Detected'] });
                }
            }

            if (!rawDate && !rawDesc && !amount) return;
            if (rawDesc.includes('total') || rawDesc.includes('balance') || rawDesc.includes('sum')) return;
            if (!rawDate) return;

            if (config.flip) amount *= -1;

            // Accumulate Sheet Total (Net Flow)
            sheetTotal += amount;
            if (pType === 'cc') ccTotal += amount; else bankTotal += amount; // retained for legacy or verification

            const displayDate = rawDate instanceof Date ? rawDate.toISOString().split('T')[0] : (rawDate || 'N/A');

            if (!categoryVal && Math.abs(amount) > 0.01) {
                if (pType === 'cc') uncategorizedCC++; else uncategorizedBank++;
                uncategorizedDetails.push({ sheet: sheet.name, row: r, date: displayDate, desc: rawDesc });
            } else if (categoryVal) {
                const catStr = categoryVal.toString().trim();
                const catLower = catStr.toLowerCase();

                if (!validCategories.has(catLower)) {
                    illegalCategories.push({ value: catStr, sheet: sheet.name, row: r, date: displayDate });
                }

                // Use Display Name for stats if available, else usage case
                const displayCat = uniqueCategories.get(catLower)?.displayName || catStr;

                if (!catStats[displayCat]) catStats[displayCat] = { total: 0, subCats: {} };
                catStats[displayCat].total += amount;

                const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';
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
                vendorStats[displayVendor] = (vendorStats[displayVendor] || 0) + amount;
            }
            if (customerVal) {
                const cStr = customerVal.toString().trim();
                const cLower = cStr.toLowerCase();
                if (!validCustomers.has(cLower)) illegalCustomers.push({ value: cStr, sheet: sheet.name, row: r, date: displayDate });

                const displayCustomer = validCustomers.get(cLower) || cStr;
                customerStats[displayCustomer] = (customerStats[displayCustomer] || 0) + amount;
            }
        });

        // Apply Accumulated Sheet Total to Linked Account (if defined)
        if (config.linkedAccount) {
            const linkName = config.linkedAccount;
            const conf = uniqueCategories.get(linkName.toLowerCase());
            // Account Logic: 
            // Assets (Bank): Dr (Normal). Net Flow (Inc-Exp). Increase = Dr.
            // catStats storage: Dr is Negative.
            // Flow: Income (+), Expense (-).
            // +100 (Inc) -> Debit Bank -> Stats should decrease (more negative).
            if (!catStats[linkName]) catStats[linkName] = { total: 0, subCats: {} };
            const previous = catStats[linkName].total;
            catStats[linkName].total -= sheetTotal;
            if (showChecker || showBS) {
                console.log(`[Linkage Logic] Applied Sheet Total (${sheetTotal.toFixed(2)}) to Account "${linkName}". Balance: ${previous.toFixed(2)} -> ${catStats[linkName].total.toFixed(2)}`);
            }

            if (showChecker) {
                console.log(`Applied Sheet Total (${sheetTotal}) to Linked Account "${linkName}". New Balance: ${catStats[linkName].total}`);
            }
        } else {
            if (showChecker || showBS) {
                console.log(`[Linkage Logic] Sheet "${config.name}" (Type: ${config.type}) has NO LINKED ACCOUNT. Total (${sheetTotal.toFixed(2)}) NOT applied to any Balance Sheet asset.`);
            }
        }
    }

    // --- 3. Process Ledger ---
    // Dynamic Mapping for Ledger
    const ledgerMap = { date: null, desc: null, category: null, subCat: null, vendor: null, customer: null, dr: null, cr: null };
    const ledgerHeader = ledgerSheet.getRow(1);

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

    // No default fallbacks - strict header matching required.

    if (showChecker) {
        console.log(`\nProcessing "Ledger":`);
        console.log(`  Mapping: ${JSON.stringify(ledgerMap)}`);
    }

    ledgerSheet.eachRow((row, r) => {
        if (r === 1) return;
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

            // Default to P&L logic if report type isn't explicit, but check valid config first
            // Standardize: Credit positive (+), Debit negative (-)
            // P&L: Income (+), Expense (-)
            // BS: Liability (+), Equity (+), Asset (-)
            const impact = (cr - dr);

            const displayCat = conf?.displayName || catStr;

            if (!catStats[displayCat]) catStats[displayCat] = { total: 0, subCats: {} };
            catStats[displayCat].total += impact;

            // Ledger SubCat aggregation
            const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';
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
                vendorStats[displayVendor] = (vendorStats[displayVendor] || 0) + (dr - cr);
            }
            if (customerVal) {
                const cStr = customerVal.toString().trim();
                const cLower = cStr.toLowerCase();
                if (!validCustomers.has(cLower)) illegalCustomers.push({ value: cStr, sheet: 'Ledger', row: r, date: displayDate });

                const displayCustomer = validCustomers.get(cLower) || cStr;
                // Customer: Net Credit (Income)
                customerStats[displayCustomer] = (customerStats[displayCustomer] || 0) + (cr - dr);
            }

            // (Integration block removed - handled via standard catStats logic)
        }
    });

    // --- 4. Prepare Reports ---
    const reports = { pl: [], bs: [] };
    // Filter by P&L report type using the Map values
    const pnlNames = Array.from(uniqueCategories.values())
        .filter(conf => conf.report === 'P&L')
        .map(conf => conf.displayName)
        .sort();

    reports.pl = pnlNames.map(n => ({ label: n, value: catStats[n] ? catStats[n].total : 0 }));
    const netIncome = reports.pl.reduce((a, b) => a + b.value, 0);

    // Balance Sheet Items
    const bsNames = Array.from(uniqueCategories.values())
        .filter(conf => conf.report === 'Balance Sheet')
        .map(conf => conf.displayName)
        .sort();

    const bsItems = bsNames.map(n => {
        let val = catStats[n] ? catStats[n].total : 0;
        // Flip sign for Assets so they display positively (if Dr > Cr)
        const conf = uniqueCategories.get(n.toLowerCase());
        if (conf && conf.accountType && conf.accountType.toLowerCase().includes('asset')) {
            val *= -1;
        }
        return { label: n, value: val };
    });

    reports.bs = [
        ...bsItems
    ];

    // Prepare Vendor / Customer Reports
    reports.vendors = Object.keys(vendorStats).map(v => ({ label: v, value: vendorStats[v] })).sort((a, b) => b.value - a.value);
    reports.customers = Object.keys(customerStats).map(c => ({ label: c, value: customerStats[c] })).sort((a, b) => b.value - a.value);

    // --- 5. Console Output ---
    function printSection(title, rows) {
        console.log(`\n--- ${title} ---`);
        if (!rows.length) { console.log('(No Data)'); return; }
        const max = Math.max(...rows.map(r => r.label.length), 10);
        rows.forEach(r => console.log(`${r.label.padEnd(max + 5)} : ${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`));
    }

    if (showAll || showPL || showPLSub) {
        console.log(`\n--- PROFIT & LOSS ---`);
        if (!reports.pl.length) console.log('(No Data)');
        else {
            const max = Math.max(...reports.pl.map(r => r.label.length), 10);
            reports.pl.forEach(r => {
                console.log(`${r.label.padEnd(max + 5)} : ${r.value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
                // Sub-Category Detail
                if (showPLSub && catStats[r.label] && catStats[r.label].subCats) {
                    const subs = catStats[r.label].subCats;
                    const subKeys = Object.keys(subs).filter(k => Math.abs(subs[k]) > 0.01);
                    if (!(subKeys.length === 1 && subKeys[0] === '(No Sub-Cat)')) {
                        subKeys.sort().forEach(sub => {
                            console.log(`   > ${sub.padEnd(max + 1)} : ${subs[sub].toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }).padStart(15)}`);
                        });
                    }
                }
            });
        }
        console.log(`\n=== NET INCOME: ${netIncome.toFixed(2)} ===\n`);
    }
    if (showAll || showBS) printSection('BALANCE SHEET', reports.bs);
    if (showAll || showVendor) printSection('VENDOR SPENDING', reports.vendors);
    if (showAll || showCustomer) printSection('CUSTOMER INCOME', reports.customers);

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

    const hasIssues = uncategorizedDetails.length > 0 || illegalCategories.length > 0 || illegalVendors.length > 0 || illegalCustomers.length > 0;
    if (hasIssues) {
        console.log('\n--- DATA INTEGRITY ISSUES ---');
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet)
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

            if (showChecker) {
                uncat.forEach(x => console.log(`      - [${x.date}] Row ${x.row}: MISSING CATEGORY ("${x.desc}")`));
                illegalCategories.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: ILLEGAL CATEGORY "${x.value}"`));
                illegalVendors.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN VENDOR "${x.value}"`));
                illegalCustomers.filter(x => x.sheet === s).forEach(x => console.log(`      - [${x.date}] Row ${x.row}: UNKNOWN CUSTOMER "${x.value}"`));
            }
        });
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

    try {
        await workbook.xlsx.writeFile(filename);
        if (showChecker || saveFlag) console.log(`\nSuccessfully updated financials in ${filename}`);
    } catch (e) {
        console.error('Error saving file:', e.message);
    }
}

updateFinancials().catch(console.error);

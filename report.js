// Entry point: reads an accounting workbook, builds the P&L, balance sheet,
// vendor/customer and 1099 reports, prints them and (with --save) writes them
// to report_<name>.xlsx / .pdf. Run with --help for options.
//
// The work is split into phases under lib/report/, which share state through
// the context object created in lib/report/context.js.
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { originalConsole, captureConsole } = require('./lib/logger');
const { parseArgs, parseOptionValue } = require('./lib/cli');
const book = require('./lib/book');
const { resolveShortcut } = require('./lib/util');
const { loadOpeningBalancesFile, resolveOpeningBalances } = require('./lib/opening-balances');
const { createContext } = require('./lib/report/context');
const { readSetup } = require('./lib/report/setup');
const { readSheetConfigs } = require('./lib/report/sheet-config');
const { getHeaderMap } = require('./lib/report/header-map');
const { loadExternalVendors, findVendorFile } = require('./lib/report/external-vendors');
const { processTransactionSheets } = require('./lib/report/transactions');
const { processLedger } = require('./lib/report/ledger');
const { validateWallets } = require('./lib/report/wallets');
const { buildReports } = require('./lib/report/build-reports');
const { printStatements } = require('./lib/report/print-statements');
const { printDiagnostics } = require('./lib/report/diagnostics');
const { writeSummaryAndSave } = require('./lib/report/summary');
const { buildStamp, stampLines } = require('./lib/report/stamp');

async function updateFinancials() {
    const args = process.argv.slice(2); // as given: this is what the report stamp records
    let runArgs = args;
    let bookConfig = null;
    if (args.some(a => a === '--book' || a.startsWith('--book=')) && !args.includes('--help')) {
        // --book: read and check the book first, so a bad book fails before the workbook is opened.
        const ref = parseOptionValue(args, '--book');
        if (!ref) throw new Error('--book needs a book id or the path of a book file (.toml).');
        bookConfig = book.loadBook(ref);
        runArgs = book.mergeBookArgs(args, bookConfig);
    }
    const opts = parseArgs(runArgs);
    if (opts.help) return;
    const { showDebug, taxYear, THRESHOLD_1099_NEC } = opts;

    // The private file (tax IDs) is only read when a 1099 report is asked for.
    const privateData = bookConfig && opts.show1099 ? book.loadPrivate(bookConfig) : null;

    // Read and check the opening balances file first, so a bad file fails before the workbook is opened.
    // On the command line it wins over the book's opening_balances.
    const openingBalancesPath = opts.openingBalancesPath || (bookConfig && bookConfig.openingBalancesPath) || null;
    const openingBalances = openingBalancesPath ? loadOpeningBalancesFile(openingBalancesPath) : null;

    // With --book the vendors are in the book, so vendor.xlsx / vendor.csv are not searched for
    // (and not warned about) unless --vendor-file names one explicitly.
    const ignoreVendors = opts.ignoreVendors || Boolean(bookConfig && !opts.customVendorFile);

    let filename = opts.filename;
    const originalInputPath = filename; // Store original input for path resolution

    // Resolve shortcut if needed, then print the report stamp (what this report is made from)
    let stamp = null;
    if (fs.existsSync(filename)) {
        const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
        console.log(`LLC Accounting Tool v${pkg.version}`);
        console.log(`Tax year: ${taxYear} (1099-NEC threshold: $${THRESHOLD_1099_NEC.toLocaleString()})`);

        const resolved = resolveShortcut(filename);
        if (resolved !== filename) {
            if (showDebug) console.log(`Resolved shortcut '${filename}' -> '${resolved}'`);
            filename = resolved;
        }

        if (fs.existsSync(filename)) {
            const inputs = [{ role: 'Workbook', path: filename }];
            if (bookConfig) inputs.push({ role: 'Book', path: bookConfig.file });
            if (privateData) inputs.push({ role: 'Private file', path: bookConfig.privatePath });
            if (openingBalancesPath) inputs.push({ role: 'Opening balances', path: openingBalancesPath });
            const vendorFile = ignoreVendors ? null : findVendorFile({ customVendorFile: opts.customVendorFile, filename, originalInputPath });
            if (vendorFile) inputs.push({ role: 'Vendor file', path: vendorFile });
            stamp = buildStamp({ args, taxYear, inputs });
            console.log('Report stamp:');
            stampLines(stamp).forEach(line => console.log(line));
        }
    }

    // Override console for capturing output EARLY to capture setup warnings
    captureConsole();

    if (!fs.existsSync(filename)) {
        console.error(`Error: File '${filename}' not found.`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    try {
        if (showDebug) console.log(`Loading workbook: ${filename}...`);
        await workbook.xlsx.readFile(filename);
    } catch (e) {
        console.error(`Error reading file: ${e.message}`);
        return;
    }

    // With --book the Setup sheet is not read: the book holds all the configuration.
    const setupSheet = bookConfig ? null : workbook.getWorksheet('Setup');
    // Try 'General Ledger' first, then 'Ledger', then case-insensitive scan
    let ledgerSheet = workbook.getWorksheet('General Ledger') || workbook.getWorksheet('Ledger');
    if (!ledgerSheet) {
        ledgerSheet = workbook.worksheets.find(s => {
            const n = s.name.trim().toLowerCase();
            return n === 'general ledger' || n === 'ledger';
        });
    }

    const summarySheet = workbook.getWorksheet('Summary');

    if ((!bookConfig && !setupSheet) || !ledgerSheet) {
        console.error('\n[ERROR] Mandatory sheets missing from workbook:');
        if (!bookConfig && !setupSheet) console.error(' - "Setup" sheet is missing.');
        if (!ledgerSheet) console.error(' - "General Ledger" (or "Ledger") sheet is missing.');
        return;
    }

    const ctx = createContext({
        ...opts, ignoreVendors, filename, originalInputPath, workbook, setupSheet, ledgerSheet, summarySheet, stamp, book: bookConfig,
    });

    if (bookConfig) {
        book.checkBookAgainstWorkbook(bookConfig, workbook.worksheets.map(s => s.name));
        if (workbook.getWorksheet('Setup')) console.log(`Setup sheet ignored: config comes from ${path.basename(bookConfig.file)}`);
        ctx.getHeaderMap = sheet => getHeaderMap(sheet, showDebug); // vendor.xlsx uses it
        book.applyBook(ctx, bookConfig);
    } else {
        await readSetup(ctx);
        await readSheetConfigs(ctx);
    }
    if (openingBalances) {
        const { count, plug } = resolveOpeningBalances(openingBalances, ctx);
        const fmt = n => n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        console.log(`Opening balances: ${count} account${count === 1 ? '' : 's'} from ${path.basename(openingBalancesPath)} (closing year ${openingBalances.year}). Plug to Retained Earnings (Opening Balance): ${fmt(plug)}`);
    }
    await loadExternalVendors(ctx);
    book.applyPrivate(ctx, privateData); // after vendor.xlsx, so the private file's tax IDs win
    await processTransactionSheets(ctx);
    await processLedger(ctx);
    await validateWallets(ctx);
    await buildReports(ctx);
    await printStatements(ctx);
    await printDiagnostics(ctx);

    if (!ctx.saveFlag) {
        console.log('\n(Run with --save to update the Excel file)');
        return;
    }

    await writeSummaryAndSave(ctx);
}

updateFinancials().catch(e => {
    console.error(`\n[CRITICAL ERROR] ${e.message}`);
    if (e.stack && process.argv.includes('--debug')) {
        originalConsole.error(e.stack);
    }
    process.exit(1);
});

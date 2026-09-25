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
const { parseArgs } = require('./lib/cli');
const { resolveShortcut } = require('./lib/util');
const { createContext } = require('./lib/report/context');
const { readSetup } = require('./lib/report/setup');
const { readSheetConfigs } = require('./lib/report/sheet-config');
const { loadExternalVendors } = require('./lib/report/external-vendors');
const { processTransactionSheets } = require('./lib/report/transactions');
const { processLedger } = require('./lib/report/ledger');
const { validateWallets } = require('./lib/report/wallets');
const { buildReports } = require('./lib/report/build-reports');
const { printStatements } = require('./lib/report/print-statements');
const { printDiagnostics } = require('./lib/report/diagnostics');
const { writeSummaryAndSave } = require('./lib/report/summary');

async function updateFinancials() {
    const opts = parseArgs(process.argv.slice(2));
    if (opts.help) return;
    const { showDebug, taxYear, THRESHOLD_1099_NEC } = opts;

    let filename = opts.filename;
    const originalInputPath = filename; // Store original input for path resolution

    // Resolve shortcut if needed
    if (fs.existsSync(filename)) {
        const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
        console.log(`LLC Accounting Tool v${pkg.version}`);
        console.log(`Tax year: ${taxYear} (1099-NEC threshold: $${THRESHOLD_1099_NEC.toLocaleString()})`);

        const resolved = resolveShortcut(filename);
        if (resolved !== filename) {
            if (showDebug) console.log(`Resolved shortcut '${filename}' -> '${resolved}'`);
            filename = resolved;
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

    const setupSheet = workbook.getWorksheet('Setup');
    // Try 'General Ledger' first, then 'Ledger', then case-insensitive scan
    let ledgerSheet = workbook.getWorksheet('General Ledger') || workbook.getWorksheet('Ledger');
    if (!ledgerSheet) {
        ledgerSheet = workbook.worksheets.find(s => {
            const n = s.name.trim().toLowerCase();
            return n === 'general ledger' || n === 'ledger';
        });
    }

    const summarySheet = workbook.getWorksheet('Summary');

    if (!setupSheet || !ledgerSheet) {
        console.error('\n[ERROR] Mandatory sheets missing from workbook:');
        if (!setupSheet) console.error(' - "Setup" sheet is missing.');
        if (!ledgerSheet) console.error(' - "General Ledger" (or "Ledger") sheet is missing.');
        return;
    }

    const ctx = createContext({
        ...opts, filename, originalInputPath, workbook, setupSheet, ledgerSheet, summarySheet,
    });

    await readSetup(ctx);
    await readSheetConfigs(ctx);
    await loadExternalVendors(ctx);
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

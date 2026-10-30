// Command-line parsing for report.js.
const { parseTaxYear, get1099Threshold } = require('./accounting');

// Returns the value of a '--flag value' or '--flag=value' option, or null if absent.
function parseOptionValue(args, flag) {
    const eq = args.find(a => a.startsWith(`${flag}=`));
    if (eq) return eq.slice(flag.length + 1) || null;
    const i = args.indexOf(flag);
    return i !== -1 && args[i + 1] ? args[i + 1] : null;
}

// Flags that take a value ('--flag value' or '--flag=value').
const OPTION_VALUE_FLAGS = ['--year', '--details', '--vendor-file', '--opening-balances', '--book'];

// args without a '--flag value' / '--flag=value' option.
function stripOption(args, flag) {
    const out = [];
    for (let i = 0; i < args.length; i++) {
        if (args[i] === flag) { i++; continue; }
        if (args[i].startsWith(`${flag}=`)) continue;
        out.push(args[i]);
    }
    return out;
}

// The first argument that is neither a flag nor a flag's value, or null.
function findPositional(args) {
    const valueIdx = new Set(OPTION_VALUE_FLAGS.map(f => args.indexOf(f)).filter(i => i !== -1).map(i => i + 1));
    return args.find((a, i) => !a.startsWith('--') && !valueIdx.has(i)) || null;
}

// Parses report.js arguments. Prints help and returns { help: true } for --help.
// Exits the process on an invalid --year or an unknown flag.
function parseArgs(args) {
    const saveFlag = args.includes('--save');
    const showAll = args.includes('--all');
    const showPL = args.includes('--pl') || showAll;
    const showBS = args.includes('--bs') || showAll;
    const showPLSub = args.includes('--pl-sub') || showAll;
    const showBSSub = args.includes('--bs-sub') || showAll;
    const showVendor = args.includes('--vendor') || showAll;
    const showVendorSub = args.includes('--vendor-sub');
    const showCustomer = args.includes('--customer') || showAll;
    const showCustomerSub = args.includes('--customer-sub');
    const showChecker = args.includes('--checker') || showAll;

    const showDebug = args.includes('--debug');
    const show1099All = args.includes('--1099') || showAll;
    const show1099NEC = args.includes('--1099=NEC') || args.includes('--1099-nec');
    const show1099INT = args.includes('--1099=INT');
    const show1099 = show1099All || show1099NEC || show1099INT;
    const ignoreVendors = args.includes('--ignore-vendors');

    // Parse --year=YYYY / --year YYYY (defaults to DEFAULT_TAX_YEAR in lib/accounting.js)
    let taxYear;
    try {
        taxYear = parseTaxYear(args);
    } catch (e) {
        console.error(`Error: ${e.message}`);
        process.exit(1);
    }
    const THRESHOLD_1099_NEC = get1099Threshold('NEC', taxYear);

    // Parse --vendor-file <path>
    const vendorFileIndex = args.indexOf('--vendor-file');
    const customVendorFile = vendorFileIndex !== -1 && args[vendorFileIndex + 1] ? args[vendorFileIndex + 1] : null;

    // Parse --opening-balances <path> / --opening-balances=<path>
    const openingBalancesPath = parseOptionValue(args, '--opening-balances');
    if (!openingBalancesPath && args.some(a => a === '--opening-balances' || a.startsWith('--opening-balances='))) {
        console.error('Error: --opening-balances needs a file path.');
        process.exit(1);
    }

    // Parse --details <Filter>
    const detailsIndex = args.indexOf('--details');
    const targetDetailsFilter = detailsIndex !== -1 && args[detailsIndex + 1] ? args[detailsIndex + 1].toLowerCase().trim() : null;
    const showDetails = !!targetDetailsFilter;

    // Help Menu
    if (args.includes('--help')) {
        console.log(`
Usage: node report.js [filename] [flags]

Description:
  Updates the financial accounting spreadsheet. It reads the Setup, Ledger, and Transaction sheets,
  categorizes transactions, balances the ledger, and generates P&L / Balance Sheet reports in standard Output format.
  Supports .lnk and .url shortcut files as input.

Arguments:
  [filename]      Path to the Excel file or shortcut (default: LLC_Accounting_Template.xlsx).
                  Not used with --book: the book names its workbook.

Flags:
  --help          Show this help message.
  --all           Run ALL standard reports (P&L Detailed, BS Detailed, Vendor, Customer, 1099). Print-only.
  --save          Save changes to the Excel file (Summary tab and formatting).
                  (Default behavior is print-only, which does not modify the file).
  --checker       Run the Data Integrity Checker and verify row-by-row categorization issues.
  --debug         Enable verbose debug output for troubleshooting.
  --pl            (Optional) Print standard Profit & Loss report.
  --bs            (Optional) Print standard Balance Sheet report.
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
  --details "Name" (Optional) List all transactions matching a specific Category, Vendor, or Customer.
  --book <id|path>  (Optional) Read ALL configuration (company, sheets, categories, vendors, customers,
                  default flags, default year) from a book file instead of the workbook's Setup sheet.
                  <id> is looked up in books.toml (or the file named by env LLC_BOOKS); <path> ends in .toml.
                  --year on the command line overrides the book's year; default_flags are added to your flags.
                  See docs/setup.md and examples/example.book.toml.
  --opening-balances <path>  (Optional) TOML file of last year's closing balance-sheet balances
                  (see examples/opening-balances.example.toml). --year must be the year after the file's year.
  --year=YYYY     (Optional) Tax year to report on. Rows dated in other years are skipped.
                  Also selects the 1099-NEC threshold ($600 before 2026, $2,000 from 2026).
                  Defaults to DEFAULT_TAX_YEAR in lib/accounting.js.

Example:
  node report.js "My_Books_2025.xlsx" --checker --save
        `);
        return { help: true };
    }

    const knownFlags = [
        '--save', '--pl', '--bs', '--vendor', '--vendor-sub', '--customer', '--customer-sub', '--pl-sub', '--bs-sub', '--checker', '--debug', '--details', '--help', '--1099', '--1099-nec', '--1099=NEC', '--1099=INT', '--ignore-vendors', '--vendor-file', '--all', '--year', '--opening-balances', '--book'
    ];

    // Check for unknown arguments
    const unknownArgs = args.filter(a => a.startsWith('--') && !knownFlags.includes(a) && !a.startsWith('--year=') && !a.startsWith('--opening-balances=') && !a.startsWith('--book='));
    if (unknownArgs.length > 0) {
        console.error(`Error: Unknown argument(s): ${unknownArgs.join(', ')}`);
        console.error('Run with --help to see available options.');
        process.exit(1);
    }

    const filename = findPositional(args) || 'LLC_Accounting_Template.xlsx';

    return {
        saveFlag,
        showAll,
        showPL,
        showBS,
        showPLSub,
        showBSSub,
        showVendor,
        showVendorSub,
        showCustomer,
        showCustomerSub,
        showChecker,
        // --checker given explicitly (not through --all): an unbalanced ledger then stops the run.
        strictChecker: args.includes('--checker'),
        showDebug,
        show1099All,
        show1099NEC,
        show1099INT,
        show1099,
        ignoreVendors,
        taxYear,
        THRESHOLD_1099_NEC,
        customVendorFile,
        openingBalancesPath,
        targetDetailsFilter,
        showDetails,
        filename,
    };
}

module.exports = { parseArgs, parseOptionValue, stripOption, findPositional };

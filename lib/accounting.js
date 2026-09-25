// Shared helpers used by report.js and load_transactions.js.

// Tax year used when --year is not given. Update this once per year.
const DEFAULT_TAX_YEAR = 2026;

// 1099-NEC / 1099-MISC reporting thresholds by payment year.
// The threshold rose from $600 to $2,000 for payments made after Dec 31, 2025.
const THRESHOLD_1099_NEC_BEFORE_2026 = 600;
const THRESHOLD_1099_NEC_FROM_2026 = 2000;
const THRESHOLD_1099_INT = 0; // Conservative: report all interest (IRS threshold is generally $10)

function get1099Threshold(type, year) {
    if (type === 'INT') return THRESHOLD_1099_INT;
    return year >= 2026 ? THRESHOLD_1099_NEC_FROM_2026 : THRESHOLD_1099_NEC_BEFORE_2026;
}

// Parses a money value from a number or a bank/Excel-formatted string.
// Handles "$1,234.56", "-1,234.56", "(1,234.56)", "-$5", "1,234.56-" and "1,234.56 CR".
// Returns 0 for blank values and NaN for text that is not a number.
function parseAmount(value) {
    if (value === null || value === undefined) return 0;
    if (typeof value === 'number') return value;
    if (typeof value === 'object' && value.result !== undefined) return parseAmount(value.result);

    let s = value.toString().trim();
    if (s === '' || s === '-') return 0;

    let negative = false;
    if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1).trim(); }
    if (/\s*CR$/i.test(s)) { negative = true; s = s.replace(/\s*CR$/i, ''); }
    if (/-$/.test(s)) { negative = true; s = s.slice(0, -1); }
    if (s.startsWith('-')) { negative = !negative; s = s.slice(1); }

    s = s.replace(/[$,\s]/g, '');
    if (!/^\d*\.?\d+$|^\d+\.$/.test(s)) return NaN;

    const n = parseFloat(s);
    return negative ? -n : n;
}

// Parses --year=YYYY or --year YYYY. Returns DEFAULT_TAX_YEAR when absent.
function parseTaxYear(args) {
    let raw = null;
    const eq = args.find(a => a.startsWith('--year='));
    if (eq) raw = eq.split('=')[1];
    const idx = args.indexOf('--year');
    if (idx !== -1) raw = args[idx + 1] === undefined ? '' : args[idx + 1];
    if (raw === null) return DEFAULT_TAX_YEAR;

    const year = parseInt(raw, 10);
    if (!/^\d{4}$/.test(String(raw)) || isNaN(year)) {
        throw new Error(`Invalid --year value "${raw}". Use a 4-digit year, e.g. --year=2025.`);
    }
    return year;
}

// Returns the calendar year of a date value, or null if it cannot be determined.
// Excel dates are read by exceljs as UTC midnight, so the UTC year is used.
function getYear(value) {
    if (value instanceof Date) return isNaN(value) ? null : value.getUTCFullYear();
    if (typeof value === 'number' && value > 10000) {
        return new Date(Math.round((value - 25569) * 86400 * 1000)).getUTCFullYear();
    }
    if (typeof value === 'string' && value.trim()) {
        const iso = value.match(/^(\d{4})-\d{1,2}-\d{1,2}/);
        if (iso) return parseInt(iso[1], 10);
        const us = value.match(/^\d{1,2}\/\d{1,2}\/(\d{2,4})/);
        if (us) return us[1].length === 2 ? 2000 + parseInt(us[1], 10) : parseInt(us[1], 10);
        const d = new Date(value);
        return isNaN(d) ? null : d.getFullYear();
    }
    return null;
}

// Parses a date string from a bank export into a UTC-midnight Date so the calendar
// day does not shift with the machine's time zone. Returns the input unchanged if
// it is not a string, and an Invalid Date if the string cannot be parsed.
function parseDateUTC(value) {
    if (typeof value !== 'string') return value;
    const s = value.trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
        const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
        return new Date(Date.UTC(y, +m[1] - 1, +m[2]));
    }
    const d = new Date(s);
    return isNaN(d) ? d : new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
}

module.exports = {
    parseDateUTC,
    DEFAULT_TAX_YEAR,
    get1099Threshold,
    parseAmount,
    parseTaxYear,
    getYear,
};

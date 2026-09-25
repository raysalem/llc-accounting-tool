// Helpers for reading values out of Excel cells and normalizing names.
const { parseAmount } = require('../accounting');

// Returns a cell's value as a number, Date or trimmed string. Formula cells
// return their result; rich text and hyperlinks return their text.
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

// Helper to safely parse balances (Start/End)
// Returns number only if valid number found. Returns null if empty string, null/undefined, or "NA"
const parseBalance = (val) => {
    if (val === null || val === undefined) return null;
    let s = val.toString().trim();
    if (s === '' || s.toLowerCase() === 'na' || s.toLowerCase() === 'n/a' || s.toLowerCase() === 'nan') return null;
    const n = parseAmount(s);
    return isNaN(n) ? null : n;
};

// Normalizes vendor/customer names for matching: lowercase, single spaces.
const NORM_VEND = (s) => (s || '').toString().toLowerCase().replace(/\s+/g, ' ').trim();

module.exports = { getVal, parseBalance, NORM_VEND };

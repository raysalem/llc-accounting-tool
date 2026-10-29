// Classifies a Balance Sheet category's Account Type (from Setup) as an
// asset, liability or equity account. Shared by build-reports.js and
// lib/opening-balances.js so both agree on which side of A = L + E it sits.

const ASSET_WORDS = ['asset', 'bank', 'cash', 'receivable'];
const LIABILITY_WORDS = ['liability', 'credit', 'cc', 'payable', 'loan', 'debt'];
const EQUITY_WORDS = ['equity', 'capital', 'contribution', 'distribution', 'earning'];

// Returns 'asset', 'liability' or 'equity'. A type that matches nothing is
// treated as an asset, as the balance sheet always has.
function classifyBalanceSheetType(accountType) {
    const t = (accountType || '').toString().toLowerCase();
    if (ASSET_WORDS.some(w => t.includes(w))) return 'asset';
    if (LIABILITY_WORDS.some(w => t.includes(w))) return 'liability';
    if (EQUITY_WORDS.some(w => t.includes(w))) return 'equity';
    return 'asset';
}

module.exports = { classifyBalanceSheetType };

// Column-header names recognized on transaction and ledger sheets.
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

const bankMapDefault = { date: null, desc: null, amount: null, category: null, subCat: null, vendor: null, customer: null };
const ccMapDefault = { date: null, desc: null, amount: null, category: null, subCat: null, vendor: null, customer: null };

function findCol(cellVal, headerList) {
    if (!cellVal) return false;
    const v = cellVal.toString().toLowerCase().trim();
    return headerList.some(h => v === h || v.includes(h)); // Relaxed matching
}

// True if a row's text looks like a header row (has Date plus Amount or Category).
function HEAD_MATCH(rowStr) {
    return (rowStr.includes('date') && (rowStr.includes('amount') || rowStr.includes('category')));
}

module.exports = { HEADERS, bankMapDefault, ccMapDefault, findCol, HEAD_MATCH };

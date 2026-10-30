// The book file's vocabulary: valid keys, types and flags. Kept apart from the
// validation code so the docs, the validator and the tests share one list.

const SHEET_TYPES = ['bank', 'cc', 'income', 'expense', 'ledger'];
const CATEGORY_TYPES = ['asset', 'liability', 'equity', 'income', 'expense', 'transfer'];
const BALANCE_SHEET_TYPES = ['asset', 'liability', 'equity'];
const FORM_1099_VALUES = ['NEC', 'INT', 'MISC', 'none'];
const LEDGER_SHEET_NAMES = ['ledger', 'general ledger']; // the names ledger.js looks for

// Report flags a book's default_flags may set. Flags that take a value, and
// --help, are left out: --year is the book's `year`, --book is the book itself.
const DEFAULT_FLAG_CHOICES = [
    '--all', '--pl', '--bs', '--pl-sub', '--bs-sub', '--vendor', '--vendor-sub', '--customer', '--customer-sub',
    '--1099', '--1099=NEC', '--1099=INT', '--checker', '--save', '--debug', '--ignore-vendors',
];

const BOOK_KEYS = {
    id: { kind: 'string', required: true },
    name: { kind: 'string' },
    workbook: { kind: 'string', required: true },
    year: { kind: 'year' },
    opening_balances: { kind: 'string' },
    private: { kind: 'string' },
    default_flags: { kind: 'strings' },
    blank_sheets: { kind: 'strings' }, // extra empty tabs for the user (e.g. a scratch "Worksheet"); the report ignores them
};

const COMPANY_KEYS = {
    name: { kind: 'string' }, address: { kind: 'string' }, city: { kind: 'string' }, state: { kind: 'string' },
    zip: { kind: 'string' }, country: { kind: 'string' }, email: { kind: 'string' }, phone: { kind: 'string' },
};

// description: a free-text note for the user on sheets, categories, vendors and customers.
// The tool reads it only to check it is text; it is never printed or used.
const DESCRIPTION = { kind: 'string' };

const SHEET_KEYS = {
    name: { kind: 'string', required: true },
    type: { kind: 'string', required: true },
    flip: { kind: 'boolean' },
    header_row: { kind: 'positive' },
    link: { kind: 'string' },
    short_name: { kind: 'string' },
    start_balance: { kind: 'money' },
    end_balance: { kind: 'money' },
    description: DESCRIPTION,
};

const CATEGORY_KEYS = {
    name: { kind: 'string', required: true },
    type: { kind: 'string', required: true },
    subcategories: { kind: 'strings' },
    transfer_account: { kind: 'string' },
    description: DESCRIPTION,
};

const VENDOR_KEYS = {
    name: { kind: 'string', required: true },
    business_name: { kind: 'string' },
    recipient_name: { kind: 'string' },
    form_1099: { kind: 'string' },
    address: { kind: 'string' }, city: { kind: 'string' }, state: { kind: 'string' }, zip: { kind: 'string' },
    country: { kind: 'string' }, email: { kind: 'string' }, phone: { kind: 'string' },
    description: DESCRIPTION,
};

const CUSTOMER_KEYS = { name: { kind: 'string', required: true }, description: DESCRIPTION };

const TOP_LEVEL_KEYS = ['book', 'company', 'sheets', 'categories', 'vendors', 'customers'];
const PRIVATE_TOP_LEVEL_KEYS = ['company', 'vendors'];
const TAX_ID_KEY = 'tin';

module.exports = {
    SHEET_TYPES, CATEGORY_TYPES, BALANCE_SHEET_TYPES, FORM_1099_VALUES, LEDGER_SHEET_NAMES, DEFAULT_FLAG_CHOICES,
    BOOK_KEYS, COMPANY_KEYS, SHEET_KEYS, CATEGORY_KEYS, VENDOR_KEYS, CUSTOMER_KEYS, TOP_LEVEL_KEYS,
    PRIVATE_TOP_LEVEL_KEYS, TAX_ID_KEY,
};

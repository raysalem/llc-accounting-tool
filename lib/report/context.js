// The shared state passed between the report phases (lib/report/*.js).
// Each phase reads what it needs from ctx and adds what later phases use.

function createContext(base) {
    return {
        ...base,

        // --- Setup reference data ---
        validCategories: new Set(), // Stores lowercase for validation
        validVendors: new Map(),    // Maps lower -> Display Name
        vendor1099Map: new Map(),   // Maps lower -> { type: 'NEC'|'INT', req: 'YES'|'NO'|'' }
        vendorDetailsMap: new Map(), // Maps lower -> strict Object of details
        payerInfo: {}, // Payer/Company Info Map
        validCustomers: new Map(),  // Maps lower -> Display Name
        uniqueCategories: new Map(), // Maps lower -> { report, accountType, displayName }
        validSubCategories: new Set(), // Set of all valid subcategories from Setup
        sheetConfigs: [],
        processedSheetTotals: {}, // Track actual processed totals per sheet (post-flip)
        walletCheckResults: [], // Store wallet validation data

        // --- Accumulated totals ---
        catStats: {},
        vendorStats: {},
        vendor1099Stats: { NEC: {}, INT: {}, MISC: {} },
        customerStats: {},
        bankTotal: 0,
        ccTotal: 0,
        uncategorizedBank: 0,
        hasErrors: false, // Track global error state for final check
        uncategorizedCC: 0,
        skippedOutOfYear: 0, // Rows skipped because they fall outside the tax year

        // --- Data integrity findings ---
        illegalCategories: [],
        illegalVendors: [],
        illegalCustomers: [],
        illegalSubCategories: [],
        uncategorizedDetails: [],
        detailsRows: [],
        offsetWarnings: [],
        duplicateCategories: [], // Track duplicate category definitions

        // Track which report types (P&L vs BS) vendors/customers are used in
        // Structure: Map<vendor, Map<reportType, Array<{sheet, row, category, date}>>>
        vendorReportUsage: new Map(),
        customerReportUsage: new Map(),
    };
}

module.exports = { createContext };

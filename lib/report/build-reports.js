// Builds the P&L, balance sheet, vendor and customer report rows from the
// accumulated totals, including opening balance equity and net income.
const { classifyBalanceSheetType } = require('./account-class');

async function buildReports(ctx) {
    const {
        catStats, customerStats, openingBalances, sheetConfigs, uniqueCategories, vendor1099Stats, vendorStats,
    } = ctx;

    const reports = { pl: [], bs: [] };
    // Filter by P&L report type using the Map values
    const pnlNames = Array.from(uniqueCategories.values())
        .filter(conf => conf.report === 'P&L')
        .map(conf => conf.displayName)
        .sort();

    reports.pl = pnlNames.map(n => {
        const stats = catStats[n] || { total: 0, sheets: {} };
        let add = 0, sub = 0;
        Object.values(stats.sheets || {}).forEach(s => {
            add += Math.abs(s.add || 0);
            sub += Math.abs(s.sub || 0);
        });
        return {
            label: n,
            value: stats.total,
            add, sub,
            sheets: stats.sheets,
            subCats: stats.subCats || {}
        };
    });
    const netIncome = reports.pl.reduce((a, b) => a + b.value, 0);

    // Balance Sheet Items
    const bsNames = Array.from(uniqueCategories.values())
        .filter(conf => conf.report === 'Balance Sheet')
        .map(conf => conf.displayName)
        .sort();

    const bsItems = bsNames.map(n => {
        const stats = catStats[n] || { total: 0, sheets: {} };
        let add = 0, sub = 0;
        Object.values(stats.sheets || {}).forEach(s => {
            add += Math.abs(s.add || 0);
            sub += Math.abs(s.sub || 0);
        });

        // Calculate Opening Balance for this Category from SheetInfo
        let opening = 0;
        sheetConfigs.forEach(s => {
            if (s.linkedAccount && s.linkedAccount.toLowerCase() === n.toLowerCase()) {
                opening += (s.startBalance || 0);
            }
        });

        return {
            label: n,
            value: stats.total,
            opening: opening,
            add, sub,
            sheets: stats.sheets,
            subCats: stats.subCats || {}
        };
    });

    const assets = [];
    const liabilities = [];
    const equity = [];

    // --- Automatic Opening Balance Equity Offset ---
    // If the user provided "Start Balance" values in Setup (SheetInfo), 
    // we must offset them into Equity so the Balance Sheet ties.
    let openingBalanceEquity = 0;
    sheetConfigs.forEach(s => {
        if (s.startBalance && s.linkedAccount) {
            const lConf = uniqueCategories.get(s.linkedAccount.toLowerCase());
            const aType = (lConf && lConf.accountType) ? lConf.accountType.toString().toLowerCase() : '';
            const isAsset = aType.includes('asset') || aType.includes('bank') || aType.includes('cash');

            // Assets increase equity, Liabilities decrease it
            if (isAsset) openingBalanceEquity += s.startBalance;
            else openingBalanceEquity -= s.startBalance;
        }
    });

    // Opening balances from --opening-balances: assets add, liabilities and equity subtract,
    // so the balances in the file tie out the same way. Linked accounts are already included
    // above, because the file set their sheet's startBalance.
    const fileBalanceByLabel = new Map(openingBalances.map(o => [o.label.toLowerCase(), o]));
    openingBalances.forEach(o => {
        openingBalanceEquity += (o.kind === 'asset') ? o.amount : -o.amount;
    });

    if (Math.abs(openingBalanceEquity) > 0.001) {
        equity.push({
            label: "Retained Earnings (Opening Balance)",
            value: openingBalanceEquity,
            opening: openingBalanceEquity, // All of it is "Opening"
            add: 0,
            sub: 0,
            sheets: { "System": { total: openingBalanceEquity, add: 0, sub: 0 } },
            subCats: {}
        });
    }

    // Prepare P&L items with 0 opening (PL is always period-based)
    reports.pl.forEach(r => r.opening = 0);

    bsItems.forEach(item => {
        const conf = uniqueCategories.get(item.label.toLowerCase());
        const kind = classifyBalanceSheetType(conf && conf.accountType);
        // Opening balance from the --opening-balances file (accounts not linked to a sheet).
        const fileAmount = (fileBalanceByLabel.get(item.label.toLowerCase()) || { amount: 0 }).amount;

        if (kind === 'asset') {
            // Determine if this Category is linked to a Sheet (Source Account)
            const isLinked = sheetConfigs.some(s => s.linkedAccount && s.linkedAccount.toLowerCase() === item.label.toLowerCase());

            // FIX: Polarity for Destination Assets (Purchased via Spending)
            // If an Asset is NOT a Source Sheet (Linked), it is likely an Accumulator (e.g. Equipment, Investment).
            // Spending (Negative Flow) should INCREASE its value.
            // Polarity Fix: Only flip if the value is NEGATIVE (Scanner derived spending).
            // If the value is POSITIVE (Ledger Debit), keep it positive.
            // The opening balance counts in this test: spending already arrives positive
            // (transactions.js), so a negative flow on an asset that starts with a balance
            // is a real decrease (a sale or a depreciation credit), not spending to flip.
            if (!isLinked && item.value + fileAmount < 0) {
                // Invert Polarity
                item.value = item.value * -1;
                item.opening = item.opening * -1;
                // Add/Sub should swap (Spending was Sub, now is Add) - But effectively we just swap values
                const oldAdd = item.add;
                const oldSub = item.sub;
                item.add = oldSub;
                item.sub = oldAdd;
            }
        }

        // Added after the flip so the flip never changes the opening balance's sign.
        item.value += fileAmount;
        item.opening += fileAmount;

        if (kind === 'asset') assets.push(item);
        else if (kind === 'liability') liabilities.push(item);
        else equity.push(item);
    });

    // Net Income flows to Retained Earnings
    equity.push({
        label: "(Result) Net Income",
        value: netIncome,
        add: (netIncome >= 0 ? netIncome : 0),
        sub: (netIncome < 0 ? Math.abs(netIncome) : 0),
        sheets: { "P&L Summary": { total: netIncome } },
        subCats: {}
    });

    reports.assets = assets;
    reports.liabilities = liabilities;
    reports.equity = equity;
    reports.bs = [...assets, ...liabilities, ...equity];

    // CHECK: Assets should never be negative
    assets.forEach(a => {
        if (a.value < -0.01) {
            console.warn(`\n[!] CRITICAL WARNING: Asset "${a.label}" is NEGATIVE (${a.value.toFixed(2)}). Assets should generally be positive.`);
            global.globalWarningCount = (global.globalWarningCount || 0) + 1;
        }
    });

    // Prepare Vendor / Customer Reports
    reports.vendors1099NEC = Object.keys(vendor1099Stats.NEC || {}).map(v => ({ label: v, value: vendor1099Stats.NEC[v] })).sort((a, b) => b.value - a.value);
    reports.vendors1099INT = Object.keys(vendor1099Stats.INT || {}).map(v => ({ label: v, value: vendor1099Stats.INT[v] })).sort((a, b) => b.value - a.value);
    reports.vendors1099MISC = Object.keys(vendor1099Stats.MISC || {}).map(v => ({ label: v, value: vendor1099Stats.MISC[v] })).sort((a, b) => b.value - a.value);

    reports.customers = Object.keys(customerStats).map(c => ({
        label: c,
        add: customerStats[c].add || 0,
        sub: customerStats[c].sub || 0,
        value: customerStats[c].total || 0,
        sheets: customerStats[c].sheets || {},
        subCats: customerStats[c].subCats || {}
    })).sort((a, b) => b.value - a.value);

    reports.vendors = Object.keys(vendorStats).map(v => ({
        label: v,
        add: vendorStats[v].add || 0,
        sub: vendorStats[v].sub || 0,
        value: vendorStats[v].total || 0,
        sheets: vendorStats[v].sheets || {},
        subCats: vendorStats[v].subCats || {}
    })).sort((a, b) => b.value - a.value);

    Object.assign(ctx, { netIncome, reports });
}

module.exports = { buildReports };

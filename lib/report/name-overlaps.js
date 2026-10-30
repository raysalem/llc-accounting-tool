// Warns about names shared between categories, sub-categories, vendors and customers.
// Used for both the Setup sheet and a --book file.

function warnNameOverlaps(ctx) {
    const { validCategories, validCustomers, validSubCategories, validVendors } = ctx;

    // --- Check for Customer/Vendor Overlap ---
    const overlapSet = new Set();
    for (const [vLower, vName] of validVendors) {
        if (validCustomers.has(vLower)) {
            overlapSet.add(vName); // or validCustomers.get(vLower)
        }
    }
    
    if (overlapSet.size > 0) {
        console.warn(`\n[!] CRITICAL ACCOUNTING WARNING: Overlap detected between Vendors and Customers.`);
        console.warn(`    The following names appear in BOTH Vendor and Customer tables:`);
        console.warn(`    ${Array.from(overlapSet).join(', ')}`);
        console.warn(`    Principle: An entity should be EITHER a Vendor (Expense side) OR a Customer (Income side), never both.`);
        console.warn(`    Fix: Rename one (e.g., "Client X (Vendor)" vs "Client X").\n`);
    }
    
    // --- Check for Sub-Category / Customer / Vendor Overlap ---
    // User Request: "ensure category, sub-catgory and vendor customer names, never are onlky used for the correct column"
    const subCatOverlaps = { cust: [], vend: [] };
    
    // Check SubCategories vs Customers
    validSubCategories.forEach(sub => {
        if (validCustomers.has(sub)) subCatOverlaps.cust.push(validCustomers.get(sub));
        if (validVendors.has(sub)) subCatOverlaps.vend.push(validVendors.get(sub));
    });
    
    if (subCatOverlaps.cust.length > 0) {
        console.warn(`\n[!] AMBIGUITY WARNING: Sub-Category vs Customer Overlap.`);
        console.warn(`    The following names are defined as BOTH a Sub-Category AND a Customer:`);
        console.warn(`    ${subCatOverlaps.cust.join(', ')}`);
        console.warn(`    Risk: If "HOLDCO" is a Customer, it should generally NOT be a Sub-Category.`);
        console.warn(`    This causes confusion in reports. Please rename one (e.g., "HOLDCO Sub" vs "HOLDCO Cust").\n`);
    }
    
    if (subCatOverlaps.vend.length > 0) {
        console.warn(`\n[!] AMBIGUITY WARNING: Sub-Category vs Vendor Overlap.`);
        console.warn(`    The following names are defined as BOTH a Sub-Category AND a Vendor:`);
        console.warn(`    ${subCatOverlaps.vend.join(', ')}`);
        console.warn(`    Fix: Rename the Sub-Category or the Vendor to be distinct.\n`);
    }
    
    // Check Categories vs Customers/Vendors (Less common, but possible)
    const catOverlaps = { cust: [], vend: [] };
    validCategories.forEach(cat => {
        if (validCustomers.has(cat)) catOverlaps.cust.push(validCustomers.get(cat));
        if (validVendors.has(cat)) catOverlaps.vend.push(validVendors.get(cat));
    });
    
    if (catOverlaps.cust.length > 0 || catOverlaps.vend.length > 0) {
        console.warn(`\n[!] AMBIGUITY WARNING: Category Name Overlaps.`);
        if (catOverlaps.cust.length) console.warn(`    Category == Customer: ${catOverlaps.cust.join(', ')}`);
        if (catOverlaps.vend.length) console.warn(`    Category == Vendor:   ${catOverlaps.vend.join(', ')}`);
        console.warn(`    Ideally, structural Categories (e.g. Sales, Rent) should not share names with entities.\n`);
    }
    
}

module.exports = { warnNameOverlaps };

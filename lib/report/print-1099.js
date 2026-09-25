// Prints the 1099 preparation section (payer info, vendors over the threshold,
// missing details) and stores the rows in reports.data1099 for the saved report's
// "1099 Data" sheet.
const { NORM_VEND } = require('./cells');

function print1099Report(ctx) {
    const {
        THRESHOLD_1099_NEC, payerInfo, reports, show1099, show1099All, show1099INT,
        show1099NEC, showAll, vendorDetailsMap,
    } = ctx;

    if (showAll || show1099) {
        const activeNEC = showAll || show1099All || show1099NEC;
        const activeINT = showAll || show1099All || show1099INT;

        // Display Payer Info (Company Info)
        const p = payerInfo;
        const pName = p['companyname'] || p['payername'] || p['businessname'] || p['name'] || 'Unknown_Payer';
        const pTIN = p['tin'] || p['ein'] || p['taxid'] || 'MISSING';
        const pAddr = p['address'] || 'MISSING';
        const pEmail = p['email'] || 'MISSING';
        const pPhone = p['phone'] || 'MISSING';

        console.log(`\n--- 1099 PREPARATION ---`);
        console.log(`Payer Information (from Setup > CompanyInfo):`);
        console.log(`  Name:    ${pName}`);
        console.log(`  Tax ID:  ${pTIN}`);
        console.log(`  Address: ${pAddr}`);
        console.log(`  Contact: ${pEmail} / ${pPhone}`);
        console.log('-'.repeat(40));

        // Generate 1099 List if any data found
        const all1099 = [];
        if (activeNEC) all1099.push(...(reports.vendors1099NEC || []).map(x => ({ ...x, form: 'NEC', threshold: THRESHOLD_1099_NEC })));
        if (activeINT) all1099.push(...(reports.vendors1099INT || []).map(x => ({ ...x, form: 'INT', threshold: 0 })));
        if (activeNEC || show1099All) all1099.push(...(reports.vendors1099MISC || []).map(x => ({ ...x, form: 'MISC', threshold: THRESHOLD_1099_NEC })));

        // Filter by threshold & polarity
        const csvRows = [];
        all1099.forEach(r => {
            const isExpense = r.value > 0; // Net Payment
            if (isExpense && r.value >= r.threshold) {
                const d = vendorDetailsMap.get(NORM_VEND(r.label)) || {};
                csvRows.push({
                    ...d,
                    amount: r.value,
                    form: r.form,
                    originalLabel: r.label
                });
            }
        });
        // Store for saveReport
        reports.data1099 = csvRows;

        if (csvRows.length > 0) {
            const payerName = payerInfo['companyname'] || payerInfo['payername'] || payerInfo['businessname'] || payerInfo['name'] || 'Unknown_Payer';

            // Payer fields
            // Since PayerInfo is loose KV, we try standard keys
            const p = payerInfo;
            const pName = payerName;
            const pTIN = p['tin'] || p['ein'] || p['taxid'] || '';
            const pAddr = p['address'] || '';
            const pCity = p['city'] || '';
            const pState = p['state'] || '';
            const pZip = p['zipcode'] || p['zip'] || '';
            const pEmail = p['email'] || '';
            const pPhone = p['phone'] || p['phonenumber'] || '';

            // Always print to screen
            // 1. Payer Info (You) - Display ONCE
            console.log('\n--- 1099 PAYER INFO (You) ---');
            console.log(`Name:    ${pName}`);
            console.log(`TIN:     ${pTIN}`);
            console.log(`Address: ${pAddr}, ${pCity}, ${pState} ${pZip}`);
            console.log(`Phone:   ${pPhone}`);
            console.log(`Email:   ${pEmail}`);

            // 2. Recipient Info
            console.log('\n--- 1099 RECIPIENT INFO (Vendors) ---');
            console.log(`Recipient`.padEnd(30) + `Business`.padEnd(25) + `TIN`.padEnd(12) + `Amount`.padStart(12) + `  Form`);
            console.log('-'.repeat(85));
            csvRows.forEach(r => {
                const rName = (r.name || 'Unknown').substring(0, 29);
                const rBiz = (r.business || '').substring(0, 24);

                // STRICT VALIDATION for Console Output too
                const missingFields = [];
                // Check r (which is merged details)
                if (!r.ssn) missingFields.push('Tax ID (SSN/EIN)');
                if (!r.address) missingFields.push('Address');
                if (!r.city) missingFields.push('City');
                if (!r.state) missingFields.push('State');
                if (!r.zip) missingFields.push('Zip');

                const statusLine = `${rName.padEnd(30)}${rBiz.padEnd(25)}${(r.ssn || '').padEnd(12)}${r.amount.toFixed(2).padStart(12)}  ${r.form}`;
                console.log(statusLine);

                if (missingFields.length > 0) {
                    const vendKey = NORM_VEND(r.originalLabel || r.name || r.business || '');
                    const inDb = vendorDetailsMap.has(vendKey);

                    if (!inDb) {
                        console.error(`  [!] ERROR: Not found in "vendor.xlsx".`);
                        // Fuzzy check for suggestions
                        const suggestions = [];
                        for (const dbKey of vendorDetailsMap.keys()) {
                            if (dbKey.includes(vendKey) || vendKey.includes(dbKey)) suggestions.push(vendorDetailsMap.get(dbKey).name);
                        }
                        if (suggestions.length > 0) {
                            console.error(`      Did you mean: ${suggestions.join(' or ')}?`);
                        }
                    } else {
                        console.error(`  [!] ERROR: Incomplete Data! Missing: ${missingFields.join(', ')}`);
                    }
                    console.error(`      update "vendor.xlsx" to fix.`);
                    ctx.hasErrors = true;
                }
            });
            if (ctx.hasErrors) {
                console.warn(`\n[!] 1099 WARNING: One or more vendors have incomplete details. Please update "vendor.xlsx" for full compliance.`);
            }
        }
    }
}

module.exports = { print1099Report };

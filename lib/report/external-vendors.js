// Loads vendor.xlsx / vendor.csv (or --vendor-file) to add or override vendor
// details used for 1099 reporting.
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { NORM_VEND, getVal } = require('./cells');

async function loadExternalVendors(ctx) {
    const {
        customVendorFile, filename, getHeaderMap, ignoreVendors, originalInputPath, validVendors,
        vendor1099Map, vendorDetailsMap,
    } = ctx;

    // --- 1.5 Load External Vendors (Override/Enrich Setup) ---
    async function loadExternalVendors() {
        // Search locations: 
        // 1. Current Working Directory (Run Area)
        // 2. Directory of the INPUT file (Shortcut location)
        // 3. Directory of the TARGET file (Actual Excel file location)

        const cwd = process.cwd();
        const targetDir = path.dirname(filename);
        const inputDir = path.dirname(originalInputPath);

        const dirsToCheck = new Set([cwd, inputDir, targetDir]); // Use Set to deduplicate
        const checkPaths = [];

        if (customVendorFile) {
            checkPaths.push(customVendorFile);
        } else {
            dirsToCheck.forEach(dir => {
                checkPaths.push(path.join(dir, 'vendor.xlsx'));
                checkPaths.push(path.join(dir, 'vendor.csv'));
            });
        }

        let loadedPath = null;
        for (const fPath of checkPaths) {
            if (fs.existsSync(fPath)) {
                loadedPath = fPath;

                const vWb = new ExcelJS.Workbook();
                if (fPath.endsWith('.csv')) await vWb.csv.readFile(fPath);
                else await vWb.xlsx.readFile(fPath);

                const vSheet = vWb.worksheets[0];
                if (!vSheet) continue;

                // Simple Header Search
                const { map: vHeaders, headerRow: vHeaderRow } = getHeaderMap(vSheet);
                // Header mapping helper for single indices
                const getVCol = (k) => {
                    const indices = vHeaders.get(k);
                    return (indices && indices.length > 0) ? indices[0] : null;
                };

                const cName = getVCol('name') || getVCol('vendor') || getVCol('fullname') || getVCol('recipientname') || getVCol('payee');
                const cBiz = getVCol('businessname') || getVCol('business') || getVCol('company');
                const cSSN = getVCol('ssn') || getVCol('taxid') || getVCol('tin') || getVCol('ein') || getVCol('ssnein') || getVCol('taxidnumber');
                const cAddr = getVCol('address') || getVCol('street') || getVCol('streetaddress') || getVCol('addr') || getVCol('address1') || getVCol('mailingaddress');
                const cEmail = getVCol('email') || getVCol('emailaddress');
                const cPhone = getVCol('phone') || getVCol('phonenumber') || getVCol('mobile') || getVCol('phone#');
                const cCountry = getVCol('country');
                const cEntityType = getVCol('type') || getVCol('entitytype') || getVCol('recipienttype');
                const cTINType = getVCol('tintype') || getVCol('taxidtype') || getVCol('tin_type');
                const cREQ = getVCol('1099') || getVCol('1099required') || getVCol('required') || getVCol('req');

                const cFirstName = getVCol('firstname') || getVCol('first');
                const cLastName = getVCol('lastname') || getVCol('last');
                const cCity = getVCol('city') || getVCol('town') || getVCol('citytown');
                const cState = getVCol('state') || getVCol('province');
                const cZip = getVCol('zipcode') || getVCol('zip') || getVCol('postalcode');

                // Headers found

                vSheet.eachRow((row, r) => {
                    if (r <= vHeaderRow) return;

                    let name = cName ? getVal(row.getCell(cName)) : '';
                    const biz = cBiz ? getVal(row.getCell(cBiz)) : '';
                    const firstName = cFirstName ? getVal(row.getCell(cFirstName)) : '';
                    const lastNameVal = cLastName ? getVal(row.getCell(cLastName)) : '';

                    // Construct Fallback Name if 'Name' column is empty
                    if (!name.trim()) {
                        if (firstName && lastNameVal) name = `${firstName} ${lastNameVal}`;
                        else if (lastNameVal) name = lastNameVal;
                        else if (firstName) name = firstName;
                    }

                    const ssn = (cSSN ? getVal(row.getCell(cSSN)) : '').toString().trim();
                    const addr = (cAddr ? getVal(row.getCell(cAddr)) : '').toString().trim();
                    const city = (cCity ? getVal(row.getCell(cCity)) : '').toString().trim();
                    const state = (cState ? getVal(row.getCell(cState)) : '').toString().trim();
                    const zip = (cZip ? getVal(row.getCell(cZip)) : '').toString().trim();


                    // Key candidate list (multi-mapping)
                    const keys = new Set();
                    if (name) keys.add(NORM_VEND(name));
                    if (biz) keys.add(NORM_VEND(biz));
                    if (firstName && lastNameVal) keys.add(NORM_VEND(`${firstName} ${lastNameVal}`));

                    if (keys.size === 0) return;

                    const details = {
                        name: name || (firstName && lastNameVal ? `${firstName} ${lastNameVal}` : (biz || '')),
                        firstName: firstName || '',
                        lastName: lastNameVal || (name && !biz ? name : ''),
                        business: biz || '',
                        ssn: ssn || '',
                        address: addr || '',
                        city: city || '',
                        state: state || '',
                        zip: zip || '',
                        email: (cEmail ? getVal(row.getCell(cEmail)) : '') || '',
                        phone: (cPhone ? getVal(row.getCell(cPhone)) : '') || '',
                        country: (cCountry ? getVal(row.getCell(cCountry)) : '') || '',
                        entityType: (cEntityType ? getVal(row.getCell(cEntityType)) : '') || '',
                        tinType: (cTINType ? getVal(row.getCell(cTINType)) : '') || ''
                    };

                    // Mark as 1099 required if column says so
                    if (cREQ) {
                        const val = getVal(row.getCell(cREQ)).toString().toLowerCase();
                        if (val === 'yes' || val === 'nec' || val === 'misc' || val === 'int') {
                            details.req = 'YES';
                            details.type = (val === 'yes') ? 'NEC' : val.toUpperCase();
                        } else if (val === 'no' || val === 'n' || val === 'false') {
                            details.req = 'NO';
                        }
                    }

                    keys.forEach(k => {
                        const existing = vendorDetailsMap.get(k) || {};
                        vendorDetailsMap.set(k, { ...existing, ...details });
                        if (!validVendors.has(k)) validVendors.set(k, details.name || details.business || k);

                        // CRITICAL: Also update the 1099 tracking map so these vendors are recognized during transaction processing
                        if (details.req === 'YES') {
                            vendor1099Map.set(k, { type: details.type || 'NEC', req: 'YES' });
                        } else if (details.req === 'NO') {
                            // Explicitly disable 1099 if external file says NO (overrides Setup)
                            vendor1099Map.delete(k);
                        }
                    });
                });
                break; // Stop after first successful file match
            }
        }
        if (!loadedPath) {
            console.warn(`\n[!] WARNING: No "vendor.xlsx" or "vendor.csv" found in search paths.`);
            console.warn(`    Looked in: ${Array.from(dirsToCheck).join(', ')}`);
            console.warn(`    1099 details will be missing.\n`);
        }
    }
    if (!ignoreVendors) {
        await loadExternalVendors();
    }
}

module.exports = { loadExternalVendors };

/**
 * Diagnostic: Find where a specific vendor/customer is used
 * 
 * Usage: node diagnose_vendor_usage.js "file.xlsx" "vendor-name"
 */

const ExcelJS = require('exceljs');
const fs = require('fs');

async function diagnoseVendorUsage() {
    const args = process.argv.slice(2);
    let filename = args[0];
    const searchName = args[1] ? args[1].toLowerCase() : null;

    if (!searchName) {
        console.error('Usage: node diagnose_vendor_usage.js "file.xlsx" "vendor-name"');
        return;
    }

    // Resolve shortcut if needed
    if (filename.endsWith('.lnk') && fs.existsSync(filename)) {
        const { execSync } = require('child_process');
        try {
            const resolved = execSync(`powershell -Command "(New-Object -ComObject WScript.Shell).CreateShortcut('${filename}').TargetPath"`, { encoding: 'utf8' }).trim();
            if (resolved && fs.existsSync(resolved)) {
                console.log(`Resolved: ${filename} -> ${resolved}`);
                filename = resolved;
            }
        } catch (e) {
            // Ignore
        }
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const setupSheet = workbook.getWorksheet('Setup');

    // Read categories and their report types
    const categoryReports = new Map();
    const headerRow = setupSheet.getRow(1);
    let colCategory, colReport;

    headerRow.eachCell((cell, colNumber) => {
        const val = cell.value ? cell.value.toString().toLowerCase().replace(/[^a-z0-9]/g, '') : '';
        if (val === 'category') colCategory = colNumber;
        if (val === 'report') colReport = colNumber;
    });

    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const cat = colCategory ? row.getCell(colCategory).value : null;
        const report = colReport ? row.getCell(colReport).value : null;
        if (cat && report) {
            categoryReports.set(cat.toString().trim().toLowerCase(), report.toString().trim());
        }
    });

    console.log(`\n=== Searching for vendor/customer: "${searchName}" ===\n`);

    const matches = [];

    // Search all sheets
    workbook.eachSheet(sheet => {
        if (sheet.name === 'Setup' || sheet.name === 'Summary') return;

        const headerRow = sheet.getRow(1);
        let colVendor, colCustomer, colCategory;

        headerRow.eachCell((cell, colNumber) => {
            const val = cell.value ? cell.value.toString().toLowerCase().replace(/[^a-z0-9]/g, '') : '';
            if (val === 'vendor') colVendor = colNumber;
            if (val === 'customer') colCustomer = colNumber;
            if (val === 'category') colCategory = colNumber;
        });

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            const vendor = colVendor ? row.getCell(colVendor).value : null;
            const customer = colCustomer ? row.getCell(colCustomer).value : null;
            const category = colCategory ? row.getCell(colCategory).value : null;

            const vendorMatch = vendor && vendor.toString().trim().toLowerCase() === searchName;
            const customerMatch = customer && customer.toString().trim().toLowerCase() === searchName;

            if (vendorMatch || customerMatch) {
                const catStr = category ? category.toString().trim() : '(no category)';
                const report = categoryReports.get(catStr.toLowerCase()) || '(unknown)';

                matches.push({
                    sheet: sheet.name,
                    row: rowNumber,
                    type: vendorMatch ? 'Vendor' : 'Customer',
                    category: catStr,
                    report: report
                });
            }
        });
    });

    if (matches.length === 0) {
        console.log(`No matches found for "${searchName}"`);
        return;
    }

    console.log(`Found ${matches.length} transaction(s):\n`);
    console.log('Sheet'.padEnd(30) + 'Row'.padEnd(6) + 'Type'.padEnd(12) + 'Category'.padEnd(25) + 'Report');
    console.log('-'.repeat(90));

    matches.forEach(m => {
        console.log(
            m.sheet.padEnd(30) +
            m.row.toString().padEnd(6) +
            m.type.padEnd(12) +
            m.category.substring(0, 24).padEnd(25) +
            m.report
        );
    });

    // Summary
    const plCount = matches.filter(m => m.report === 'P&L').length;
    const bsCount = matches.filter(m => m.report === 'BS' || m.report === 'Balance Sheet').length;

    console.log('\n=== Summary ===');
    console.log(`P&L transactions: ${plCount}`);
    console.log(`Balance Sheet transactions: ${bsCount}`);

    if (plCount > 0 && bsCount > 0) {
        console.log('\n⚠️  This vendor/customer appears in BOTH P&L and BS categories!');
        console.log('This usually indicates a data entry error.');
    }
}

diagnoseVendorUsage().catch(console.error);

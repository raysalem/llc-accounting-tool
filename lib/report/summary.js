// Handles --save: fills the Summary sheet in memory and writes the
// report_<name>.xlsx / .pdf output.
const { logBuffer, originalConsole } = require('../logger');
const { saveReport } = require('../output/excel-report');

async function writeSummaryAndSave(ctx) {
    const {
        filename, hasIssues, illegalCategories, illegalCustomers, illegalVendors, netIncome,
        payerInfo, reportSheetList, reports, saveFlag, sheetNameMap, show1099, showAll, showBS,
        showBSSub, showChecker, showCustomer, showCustomerSub, showDebug, showPL, showPLSub,
        showVendor, showVendorSub, uncategorizedDetails, vendor1099Stats, vendorDetailsMap,
        workbook,
    } = ctx;

    // --- 6. Summary Sheet Update ---
    if (!ctx.summarySheet) {
        ctx.summarySheet = workbook.addWorksheet('Summary');
    } else {
        // Clear existing content to avoid breaking workbook references/Tables
        ctx.summarySheet.eachRow((row, r) => {
            row.eachCell(cell => { cell.value = null; cell.style = {}; });
        });
    }

    // Explicitly disable worksheet-level autofilter to avoid conflicts with Table-level filters
    // summarySheet.autoFilter = null; // Removed to prevent corruption if Table exists

    ctx.summarySheet.getCell('A1').value = `Financial Summary (${new Date().toLocaleString()})`;
    ctx.summarySheet.getCell('A1').font = { size: 14, bold: true };

    let summaryRow = 3;
    ctx.summarySheet.getCell(`A${summaryRow}`).value = 'Profit & Loss';
    ctx.summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
    reports.pl.forEach(r => { ctx.summarySheet.getCell(`A${summaryRow}`).value = r.label; ctx.summarySheet.getCell(`B${summaryRow}`).value = r.value; summaryRow++; });
    ctx.summarySheet.getCell(`A${summaryRow}`).value = 'NET INCOME'; ctx.summarySheet.getCell(`B${summaryRow}`).value = netIncome;
    ctx.summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow += 3;

    ctx.summarySheet.getCell(`A${summaryRow}`).value = 'Balance Sheet';
    ctx.summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
    reports.bs.forEach(r => { ctx.summarySheet.getCell(`A${summaryRow}`).value = r.label; ctx.summarySheet.getCell(`B${summaryRow}`).value = r.value; summaryRow++; });

    if (hasIssues) {
        summaryRow += 3;
        ctx.summarySheet.getCell(`A${summaryRow}`).value = 'Data Integrity Check';
        ctx.summarySheet.getCell(`A${summaryRow}`).font = { bold: true, color: { argb: 'FFFF0000' } }; summaryRow++;
        const issueSheetsFound = new Set([
            ...uncategorizedDetails.map(x => x.sheet),
            ...illegalCategories.map(x => x.sheet),
            ...illegalVendors.map(x => x.sheet),
            ...illegalCustomers.map(x => x.sheet)
        ]);
        issueSheetsFound.forEach(s => {
            ctx.summarySheet.getCell(`A${summaryRow}`).value = `Tab: ${s.toUpperCase()}`;
            ctx.summarySheet.getCell(`A${summaryRow}`).font = { bold: true }; summaryRow++;
            const uncat = uncategorizedDetails.filter(x => x.sheet === s).length;
            if (uncat) { ctx.summarySheet.getCell(`A${summaryRow}`).value = '  Uncategorized Rows'; ctx.summarySheet.getCell(`B${summaryRow}`).value = uncat; summaryRow++; }
            const cats = Array.from(new Set(illegalCategories.filter(x => x.sheet === s).map(x => x.value))).join(', ');
            if (cats) { ctx.summarySheet.getCell(`A${summaryRow}`).value = '  Illegal Categories'; ctx.summarySheet.getCell(`B${summaryRow}`).value = cats; summaryRow++; }
            const vends = Array.from(new Set(illegalVendors.filter(x => x.sheet === s).map(x => x.value))).join(', ');
            if (vends) { ctx.summarySheet.getCell(`A${summaryRow}`).value = '  Unknown Vendors'; ctx.summarySheet.getCell(`B${summaryRow}`).value = vends; summaryRow++; }
            summaryRow++;
        });
    }

    reports.sheetList = reportSheetList;
    reports.sheetNameMap = sheetNameMap;

    if (show1099 && showDebug) {
        console.log('\n[DEBUG] Vendor 1099 Stats Dump:');
        console.log(JSON.stringify(vendor1099Stats, null, 2));
    }

    if (saveFlag) {
        await saveReport(filename, reports, logBuffer, {
            showPL: showPL || showPLSub || showAll,
            showBS: showBS || showBSSub || showAll,
            showVendor: showVendor || showVendorSub || showAll,
            showCustomer: showCustomer || showCustomerSub || showAll,
            show1099: show1099 || showAll,
            showPLSub, showBSSub, showVendorSub, showCustomerSub,
            showChecker: showChecker || showAll,
            showAll
        }, vendorDetailsMap, payerInfo);
    }

    if (global.globalWarningCount > 0) {
        originalConsole.error(`\n[BATCH STOP] Process exited with ${global.globalWarningCount} warnings/errors.`);
        process.exit(1);
    }
}

module.exports = { writeSummaryAndSave };

// Writes the --save output: report_<name>.xlsx with one sheet per requested
// report, the processing log and run metadata, plus a PDF copy when pdfkit is installed.
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { originalConsole } = require('../logger');
const { generatePDF, PDFDocument } = require('./pdf');

async function saveReport(originalFilename, reports, logs, flags, vendorDetails, payerInfo) {
    const dir = path.dirname(originalFilename);
    const base = path.basename(originalFilename, path.extname(originalFilename));
    const newFilename = path.join(dir, `report_${base}.xlsx`);

    originalConsole.log(`\n[Saving] Generating report file: ${newFilename} ...`);

    const wb = new ExcelJS.Workbook();

    // 1. Profit & Loss (Standard + Sub)
    if (flags.showPL) {
        const ws = wb.addWorksheet('Profit & Loss');
        ws.columns = [{ header: 'Category', key: 'cat', width: 35 }, { header: 'Amount', key: 'amt', width: 15 }];
        reports.pl.forEach(r => ws.addRow({ cat: r.label, amt: r.value }));
        const netPL = reports.pl.reduce((acc, r) => acc + r.value, 0);
        const lastRow = ws.addRow({ cat: 'NET INCOME', amt: netPL });
        lastRow.font = { bold: true };
        ws.getColumn(2).numFmt = '#,##0.00';

        if (flags.showPLSub) {
            const wsSub = wb.addWorksheet('Profit & Loss Detailed');
            const sortedSheets = reports.sheetList || [];
            const sheetNameMap = reports.sheetNameMap || {};
            const header = ['Category', 'Grand Total'];
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Add)`));
            header.push('Total Additions');
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Sub)`));
            header.push('Total Subtractions');
            wsSub.columns = header.map(h => ({ header: h, width: h === 'Category' ? 35 : 15 }));

            reports.pl.forEach(c => {
                const row = [c.label, c.value];
                sortedSheets.forEach(s => row.push(c.sheets && c.sheets[s] ? Math.abs(c.sheets[s].add || 0) : 0));
                row.push(Math.abs(c.add || 0));
                sortedSheets.forEach(s => row.push(c.sheets && c.sheets[s] ? Math.abs(c.sheets[s].sub || 0) : 0));
                row.push(Math.abs(c.sub || 0));
                wsSub.addRow(row);

                if (c.subCats && Object.keys(c.subCats).length > 0) {
                    Object.entries(c.subCats).forEach(([subName, subTotal]) => {
                        // Show "(No Sub-Cat)" only when other sub-categories exist, so the lines add up to the total
                        if (subName === '(No Sub-Cat)' && Object.keys(c.subCats).length === 1) return;
                        const subRow = ['  > ' + subName, subTotal];
                        for (let i = 2; i < header.length; i++) subRow.push(null);
                        const r = wsSub.addRow(subRow);
                        r.font = { italic: true };
                    });
                }
            });

            if (reports.pl.length > 0) {
                const totalRow = ['NET INCOME', netPL];
                sortedSheets.forEach(s => {
                    totalRow.push(reports.pl.reduce((sum, r) => sum + (r.sheets && r.sheets[s] ? Math.abs(r.sheets[s].add || 0) : 0), 0));
                });
                totalRow.push(reports.pl.reduce((sum, r) => sum + Math.abs(r.add || 0), 0));
                sortedSheets.forEach(s => {
                    totalRow.push(reports.pl.reduce((sum, r) => sum + (r.sheets && r.sheets[s] ? Math.abs(r.sheets[s].sub || 0) : 0), 0));
                });
                totalRow.push(reports.pl.reduce((sum, r) => sum + Math.abs(r.sub || 0), 0));
                const tRow = wsSub.addRow(totalRow);
                tRow.font = { bold: true };
            }
            for (let c = 2; c <= header.length; c++) wsSub.getColumn(c).numFmt = '#,##0.00';
        }
    }

    // 2. Balance Sheet (Standard + Sub)
    if (flags.showBS) {
        const ws = wb.addWorksheet('Balance Sheet');
        ws.columns = [{ header: 'Account', key: 'acc', width: 35 }, { header: 'Opening Balance', key: 'open', width: 22 }, { header: 'Ending Balance', key: 'bal', width: 22 }];

        const addBSGroup = (title, items) => {
            const groupHeader = ws.addRow({ acc: title });
            groupHeader.font = { bold: true, underline: true };
            items.forEach(r => ws.addRow({ acc: r.label, open: r.opening || 0, bal: r.value }));
            const total = items.reduce((sum, r) => sum + r.value, 0);
            const totalRow = ws.addRow({ acc: `TOTAL ${title}`, bal: total });
            totalRow.font = { bold: true };
            ws.addRow([]);
            return total;
        };

        const totalAssets = addBSGroup('ASSETS', reports.assets);
        const totalLiabilities = addBSGroup('LIABILITIES', reports.liabilities);
        const totalEquity = addBSGroup('EQUITY', reports.equity);

        const summaryRow = ws.addRow({ acc: 'TOTAL LIABILITIES + EQUITY', bal: totalLiabilities + totalEquity });
        summaryRow.font = { bold: true };
        const eqDiff = Math.abs(totalAssets - (totalLiabilities + totalEquity));
        ws.addRow({ acc: 'Equation Status', bal: eqDiff < 0.01 ? 'OK' : `FAIL (Diff: ${eqDiff.toFixed(2)})` });
        ws.getColumn(2).numFmt = '#,##0.00';
        ws.getColumn(3).numFmt = '#,##0.00';

        if (flags.showBSSub) {
            const wsSub = wb.addWorksheet('Balance Sheet Detailed');
            const sortedSheets = reports.sheetList || [];
            const sheetNameMap = reports.sheetNameMap || {};
            const header = ['Account', 'Opening Balance', 'Ending Balance'];
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Add)`));
            header.push('Total Additions');
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Sub)`));
            header.push('Total Subtractions');
            wsSub.columns = header.map(h => ({ header: h, width: h === 'Account' ? 35 : 15 }));

            const addBSSubGroup = (title, items) => {
                const groupHeader = wsSub.addRow([title]);
                groupHeader.font = { bold: true, underline: true };
                items.forEach(c => {
                    // Row Structure: Account, Opening, Ending, Sheets(Add)..., Total Add, Sheets(Sub)..., Total Sub
                    const row = [c.label, c.opening || 0, c.value];
                    sortedSheets.forEach(s => row.push(c.sheets && c.sheets[s] ? Math.abs(c.sheets[s].add || 0) : 0));
                    row.push(Math.abs(c.add || 0));
                    sortedSheets.forEach(s => row.push(c.sheets && c.sheets[s] ? Math.abs(c.sheets[s].sub || 0) : 0));
                    row.push(Math.abs(c.sub || 0));
                    wsSub.addRow(row);
                    if (c.subCats && Object.keys(c.subCats).length > 0) {
                        Object.entries(c.subCats).forEach(([subName, subTotal]) => {
                            // Show "(No Sub-Cat)" only when other sub-categories exist, so the lines add up to the total
                            if (subName === '(No Sub-Cat)' && Object.keys(c.subCats).length === 1) return;
                            const subRow = ['  > ' + subName, null, subTotal];
                            // Fill blank columns for sub-cats
                            for (let i = 3; i < header.length; i++) subRow.push(null);
                            const r = wsSub.addRow(subRow);
                            r.font = { italic: true };
                        });
                    }

                });

                // Calculate Total Row
                const totalRow = [`TOTAL ${title}`, items.reduce((sum, r) => sum + (r.opening || 0), 0), items.reduce((sum, r) => sum + r.value, 0)];
                sortedSheets.forEach(s => {
                    const sheetAdd = items.reduce((sum, r) => sum + (r.sheets && r.sheets[s] ? Math.abs(r.sheets[s].add || 0) : 0), 0);
                    totalRow.push(sheetAdd);
                });
                totalRow.push(items.reduce((sum, r) => sum + Math.abs(r.add || 0), 0));
                sortedSheets.forEach(s => {
                    const sheetSub = items.reduce((sum, r) => sum + (r.sheets && r.sheets[s] ? Math.abs(r.sheets[s].sub || 0) : 0), 0);
                    totalRow.push(sheetSub);
                });
                totalRow.push(items.reduce((sum, r) => sum + Math.abs(r.sub || 0), 0));
                const tRowLabel = wsSub.addRow(totalRow);
                tRowLabel.font = { bold: true };
                wsSub.addRow([]);

            };
            addBSSubGroup('ASSETS', reports.assets);
            addBSSubGroup('LIABILITIES', reports.liabilities);
            addBSSubGroup('EQUITY', reports.equity);
            const statusRow = ['Equation Status', null, eqDiff < 0.01 ? 'OK' : `FAIL (Diff: ${eqDiff.toFixed(2)})`];
            wsSub.addRow(statusRow);
            // Apply numeric format to ALL amount columns
            for (let c = 2; c <= header.length; c++) wsSub.getColumn(c).numFmt = '#,##0.00';
        }
    }

    // 3. Vendor Report (Standard + Sub)
    if (flags.showVendor) {
        const ws = wb.addWorksheet('Vendor Spending');
        ws.columns = [{ header: 'Vendor', key: 'v', width: 30 }, { header: 'Net Amount', key: 'a', width: 15 }];
        reports.vendors.forEach(v => ws.addRow({ v: v.label, a: v.value }));
        ws.getColumn(2).numFmt = '#,##0.00';
        if (flags.showVendorSub) {
            const wsSub = wb.addWorksheet('Vendor Spending Detailed');
            const sortedSheets = reports.sheetList || [];
            const sheetNameMap = reports.sheetNameMap || {};
            const header = ['Vendor', 'Grand Total'];
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Add)`));
            header.push('Total Additions');
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Sub)`));
            header.push('Total Subtractions');
            wsSub.columns = header.map(h => ({ header: h, width: h === 'Vendor' ? 30 : 15 }));
            reports.vendors.forEach(v => {
                const row = [v.label, v.value];
                sortedSheets.forEach(s => row.push(v.sheets && v.sheets[s] ? Math.abs(v.sheets[s].add || 0) : 0));
                row.push(Math.abs(v.add || 0));
                sortedSheets.forEach(s => row.push(v.sheets && v.sheets[s] ? Math.abs(v.sheets[s].sub || 0) : 0));
                row.push(Math.abs(v.sub || 0));
                wsSub.addRow(row);
                if (v.subCats && Object.keys(v.subCats).length > 0) {
                    Object.entries(v.subCats).forEach(([subName, subTotal]) => {
                        // Show "(No Sub-Cat)" only when other sub-categories exist, so the lines add up to the total
                        if (subName === '(No Sub-Cat)' && Object.keys(v.subCats).length === 1) return;
                        const subRow = ['  > ' + subName, subTotal];
                        for (let i = 2; i < header.length; i++) subRow.push(null);
                        const r = wsSub.addRow(subRow);
                        r.font = { italic: true };
                    });
                }
            });
            for (let c = 2; c <= header.length; c++) wsSub.getColumn(c).numFmt = '#,##0.00';
        }
    }

    // 4. Customer Report (Standard + Sub)
    if (flags.showCustomer) {
        const ws = wb.addWorksheet('Customer Income');
        ws.columns = [{ header: 'Customer', key: 'c', width: 30 }, { header: 'Net Amount', key: 'a', width: 15 }];
        reports.customers.forEach(c => ws.addRow({ c: c.label, a: c.value }));
        ws.getColumn(2).numFmt = '#,##0.00';
        if (flags.showCustomerSub) {
            const wsSub = wb.addWorksheet('Customer Income Detailed');
            const sortedSheets = reports.sheetList || [];
            const sheetNameMap = reports.sheetNameMap || {};
            const header = ['Customer', 'Grand Total'];
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Add)`));
            header.push('Total Additions');
            sortedSheets.forEach(s => header.push(`${sheetNameMap[s] || s} (Sub)`));
            header.push('Total Subtractions');
            wsSub.columns = header.map(h => ({ header: h, width: h === 'Customer' ? 30 : 15 }));
            reports.customers.forEach(c => {
                const row = [c.label, c.value];
                sortedSheets.forEach(s => row.push(c.sheets && c.sheets[s] ? Math.abs(c.sheets[s].add || 0) : 0));
                row.push(Math.abs(c.add || 0));
                sortedSheets.forEach(s => row.push(c.sheets && c.sheets[s] ? Math.abs(c.sheets[s].sub || 0) : 0));
                row.push(Math.abs(c.sub || 0));
                wsSub.addRow(row);
                if (c.subCats && Object.keys(c.subCats).length > 0) {
                    Object.entries(c.subCats).forEach(([subName, subTotal]) => {
                        // Show "(No Sub-Cat)" only when other sub-categories exist, so the lines add up to the total
                        if (subName === '(No Sub-Cat)' && Object.keys(c.subCats).length === 1) return;
                        const subRow = ['  > ' + subName, subTotal];
                        for (let i = 2; i < header.length; i++) subRow.push(null);
                        const r = wsSub.addRow(subRow);
                        r.font = { italic: true };
                    });
                }
            });
            for (let c = 2; c <= header.length; c++) wsSub.getColumn(c).numFmt = '#,##0.00';
        }
    }

    // 5. 1099 Report
    if (flags.show1099) {
        const ws = wb.addWorksheet('1099 Data');
        ws.columns = [
            // R = Recipient
            { header: 'R Name', width: 25 }, { header: 'R Business Name', width: 25 },
            { header: 'R TIN', width: 15 }, { header: 'R Address', width: 30 },
            { header: 'R Email', width: 25 }, { header: 'R Phone', width: 15 },
            // P = Payer
            { header: 'P Name', width: 25 }, { header: 'P TIN', width: 15 },
            { header: 'P Address', width: 30 }, { header: 'P City', width: 15 },
            { header: 'P State', width: 5 }, { header: 'P Zip', width: 10 },
            { header: 'P Country', width: 10 }, { header: 'P Email', width: 20 }, { header: 'P Phone', width: 15 },
            // Details
            { header: 'Amount', width: 15 }, { header: 'Form 1099 Type', width: 10 }
        ];

        // Prepare Payer Info once
        const p = payerInfo || {};
        const pName = p['companyname'] || p['payername'] || p['businessname'] || p['name'] || '';
        const pTIN = p['tin'] || p['ein'] || p['taxid'] || '';
        const pAddr = p['address'] || '';
        const pCity = p['city'] || '';
        const pState = p['state'] || '';
        const pZip = p['zipcode'] || p['zip'] || '';
        const pCountry = p['country'] || '';
        const pEmail = p['email'] || '';
        const pPhone = p['phone'] || p['phonenumber'] || '';

        (reports.data1099 || []).forEach(details => {
            ws.addRow([
                details.name || '', details.business || '', details.ssn || '', details.address || '', details.email || '', details.phone || '',
                pName, pTIN, pAddr, pCity, pState, pZip, pCountry, pEmail, pPhone,
                details.amount, details.form
            ]);
        });
        ws.getColumn(16).numFmt = '#,##0.00'; // Amount Column
    }

    // 6. Processing Log
    const logWs = wb.addWorksheet('Processing Log');
    logWs.getColumn(1).width = 120;
    logs.forEach(l => logWs.addRow([l.replace(/\x1b\[[0-9;]*m/g, '')]));

    const newVersion = JSON.parse(fs.readFileSync(path.join(__dirname, '..', '..', 'package.json'), 'utf8')).version;

    // Recover original input path directly from process.argv
    const rawArgs = process.argv.slice(2);
    const originalInputPath = rawArgs.find(a => !a.startsWith('--')) || 'LLC_Accounting_Template.xlsx';

    // Metadata Sheet
    addMetadataSheet(wb, newVersion, originalInputPath, rawArgs.join(' '));

    await wb.xlsx.writeFile(newFilename);
    originalConsole.log(`[Saved] Report saved to: ${newFilename}`);

    // PDF Generation
    if (PDFDocument) {
        const pdfName = newFilename.replace(path.extname(newFilename), '.pdf');
        await generatePDF(pdfName, reports, newVersion, originalInputPath);
    }

}

function addMetadataSheet(wb, version, filename, command) {
    let ws = wb.getWorksheet('Report Info');
    if (ws) {
        wb.removeWorksheet(ws.id); // ExcelJS remove by ID
    }
    ws = wb.addWorksheet('Report Info');

    ws.columns = [{ header: 'Property', width: 25 }, { header: 'Value', width: 80 }];
    ws.addRow(['Report Version', version]);
    ws.addRow(['Generated Date', new Date().toLocaleString()]);
    try {
        ws.addRow(['Accounting File Last Save', fs.statSync(filename).mtime.toLocaleString()]);
    } catch (e) { ws.addRow(['Accounting File Last Save', 'Unknown']); }
    ws.addRow(['Command Used', command]);
}

module.exports = { saveReport };

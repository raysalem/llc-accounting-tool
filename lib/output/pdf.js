// Writes a PDF version of the saved reports. pdfkit is optional: when it is not
// installed, PDFDocument is undefined and no PDF is written.
const fs = require('fs');
const path = require('path');
const { originalConsole } = require('../logger');

let PDFDocument;
try { PDFDocument = require('pdfkit'); } catch (e) { /* ignore if missing */ }

async function generatePDF(pdfPath, reports, version, xlsxFile) {
    return new Promise((resolve, reject) => {
        try {
            // PDF in Landscape for better table fitting
            const doc = new PDFDocument({ margin: 40, layout: 'landscape' });
            const stream = fs.createWriteStream(pdfPath);
            doc.pipe(stream);

            doc.fontSize(24).text('Financial Report', { align: 'center' });
            doc.moveDown();
            doc.fontSize(14).text(`Version: ${version}`, { align: 'center' });
            doc.text(`Date: ${new Date().toLocaleString()}`, { align: 'center' });
            doc.text(`Source: ${path.basename(xlsxFile)}`, { align: 'center' });
            doc.moveDown(2);

            // Helper to print simple 2-col table
            const printTable = (title, rows, extractFunc) => {
                doc.addPage();
                doc.fontSize(18).text(title, { underline: true });
                doc.moveDown();
                doc.fontSize(10).font('Courier');

                if (!rows || rows.length === 0) {
                    doc.text('(No Data)');
                    return;
                }
                rows.forEach(row => {
                    const line = extractFunc(row);
                    doc.text(line);
                });
                doc.font('Helvetica');
            };

            // Helper to print DETAILED 5-col table (Category | Open | Add | Sub | End)
            const printDetailed = (title, rows) => {
                doc.addPage();
                doc.fontSize(18).text(title, { underline: true });
                doc.moveDown();
                doc.fontSize(9).font('Courier'); // Smaller font for width

                if (!rows || rows.length === 0) {
                    doc.text('(No Data)');
                    return;
                }

                // Header
                const h1 = "Category".padEnd(40);
                const h2 = "Opening".padStart(15);
                const h3 = "Additions".padStart(15);
                const h4 = "Subtracts".padStart(15);
                const h5 = "Ending".padStart(15);
                doc.text(`${h1} ${h2} ${h3} ${h4} ${h5}`);
                doc.text('-'.repeat(100));

                rows.forEach(r => {
                    const name = (r.label || '').substring(0, 38).padEnd(40);
                    const open = (r.opening || 0).toFixed(2).padStart(15);
                    const add = (r.add || 0).toFixed(2).padStart(15);
                    const sub = (r.sub || 0).toFixed(2).padStart(15);
                    const end = (r.value || 0).toFixed(2).padStart(15);
                    doc.text(`${name} ${open} ${add} ${sub} ${end}`);
                });
                doc.font('Helvetica');
            };

            // P&L
            if (reports.pl) {
                printTable('Profit & Loss Summary', reports.pl, r => `${r.label.padEnd(60)} ${r.value.toFixed(2).padStart(20)}`);
                const plTotal = reports.pl.reduce((s, x) => s + (x.value || 0), 0);
                doc.font('Courier').fontSize(10);
                doc.text('-'.repeat(85));
                doc.font('Courier-Bold');
                doc.text(`NET INCOME:`.padEnd(60) + plTotal.toFixed(2).padStart(20));
                doc.font('Helvetica');
            }

            // BS Summary
            if (reports.bs) {
                printTable('Balance Sheet Summary', reports.bs, r => `${r.label.padEnd(60)} ${r.value.toFixed(2).padStart(20)}`);
            }

            // BS Detailed (Assets)
            if (reports.assets) {
                printDetailed('Assets (Detailed)', reports.assets);
            }

            // BS Detailed (Liabilities)
            if (reports.liabilities) {
                printDetailed('Liabilities (Detailed)', reports.liabilities);
            }

            // BS Detailed (Equity)
            if (reports.equity) {
                printDetailed('Equity (Detailed)', reports.equity);
            }

            // Vendor Summary
            if (reports.vendors) {
                printTable('Vendor Spending', reports.vendors, r => `${r.label.substring(0, 45).padEnd(50)} ${r.value.toFixed(2).padStart(15)}`);
            }

            // Customer Summary
            if (reports.customers) {
                printTable('Customer Income', reports.customers, r => `${r.label.substring(0, 45).padEnd(50)} ${r.value.toFixed(2).padStart(15)}`);
            }

            doc.end();
            stream.on('finish', () => {
                originalConsole.log(`[PDF] Saved to: ${pdfPath}`);
                resolve();
            });
            stream.on('error', reject);

        } catch (e) {
            originalConsole.error('PDF Generation Failed:', e.message);
            resolve(); // Don't crash main process
        }
    });
}

module.exports = { generatePDF, PDFDocument };

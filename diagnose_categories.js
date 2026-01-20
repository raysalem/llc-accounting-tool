/**
 * Diagnostic Tool: Show Category Classifications
 * 
 * This script reads the Setup sheet and shows how each category is classified (P&L vs BS).
 * Helps users diagnose why categories appear in the wrong report.
 */

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function diagnoseCategoryClassification() {
    const args = process.argv.slice(2);
    let filename = args.find(a => !a.startsWith('--')) || 'LLC_Accounting_Template.xlsx';

    // Resolve shortcut if needed
    if (filename.endsWith('.lnk') && fs.existsSync(filename)) {
        const { execSync } = require('child_process');
        try {
            const resolved = execSync(`powershell -Command "(New-Object -ComObject WScript.Shell).CreateShortcut('${filename}').TargetPath"`, { encoding: 'utf8' }).trim();
            if (resolved && fs.existsSync(resolved)) {
                console.log(`Resolved shortcut: ${filename} -> ${resolved}`);
                filename = resolved;
            }
        } catch (e) {
            // Ignore
        }
    }

    if (!fs.existsSync(filename)) {
        console.error(`Error: File '${filename}' not found.`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);

    const setupSheet = workbook.getWorksheet('Setup');
    if (!setupSheet) {
        console.error('Error: Setup sheet not found.');
        return;
    }

    // Find columns
    const headerRow = setupSheet.getRow(1);
    let colCategory, colReport, colType;

    headerRow.eachCell((cell, colNumber) => {
        const val = cell.value ? cell.value.toString().toLowerCase().replace(/[^a-z0-9]/g, '') : '';
        if (val === 'category') colCategory = colNumber;
        if (val === 'report') colReport = colNumber;
        if (val === 'type') colType = colNumber;
    });

    if (!colCategory || !colReport) {
        console.error('Error: Could not find Category or Report columns in Setup sheet.');
        return;
    }

    console.log('\n=== CATEGORY CLASSIFICATION REPORT ===\n');
    console.log('Category'.padEnd(30) + 'Type'.padEnd(15) + 'Report');
    console.log('-'.repeat(60));

    const categories = [];
    setupSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const catName = colCategory ? row.getCell(colCategory).value : null;
        if (catName) {
            const cat = catName.toString().trim();
            const type = colType ? (row.getCell(colType).value || '').toString().trim() : '';
            const report = colReport ? (row.getCell(colReport).value || '').toString().trim() : '';

            categories.push({ cat, type, report });
        }
    });

    // Sort and display
    categories.sort((a, b) => {
        if (a.report !== b.report) return a.report.localeCompare(b.report);
        return a.cat.localeCompare(b.cat);
    });

    let currentReport = '';
    categories.forEach(({ cat, type, report }) => {
        if (report !== currentReport) {
            console.log('');
            currentReport = report;
        }
        console.log(cat.padEnd(30) + type.padEnd(15) + report);
    });

    console.log('\n=== SUMMARY ===');
    const plCount = categories.filter(c => c.report === 'P&L').length;
    const bsCount = categories.filter(c => c.report === 'BS' || c.report === 'Balance Sheet').length;
    const otherCount = categories.filter(c => c.report !== 'P&L' && c.report !== 'BS' && c.report !== 'Balance Sheet').length;

    console.log(`P&L Categories: ${plCount}`);
    console.log(`Balance Sheet Categories: ${bsCount}`);
    if (otherCount > 0) console.log(`Other/Unclassified: ${otherCount}`);

    console.log('\nTo fix misclassified categories:');
    console.log('1. Open the Excel file');
    console.log('2. Go to the Setup tab');
    console.log('3. Find the category in the Category column');
    console.log('4. Change the Report column to either "P&L" or "BS"');
    console.log('5. Save the file\n');
}

diagnoseCategoryClassification().catch(console.error);

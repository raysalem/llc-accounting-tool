// Checks that a workbook is not open in Excel before anything else is read or
// changed. Excel holds an open workbook locked, so opening it for update fails.
const fs = require('fs');

const LOCKED_CODES = ['EBUSY', 'EPERM', 'EACCES'];

// Throws a clear Error when file is locked (or not writable). fsImpl is for tests.
function assertWorkbookUnlocked(file, fsImpl = fs) {
    let fd;
    try {
        fd = fsImpl.openSync(file, 'r+');
    } catch (e) {
        if (e.code === 'ENOENT') throw new Error(`Workbook "${file}" not found. Fix: check "workbook" in the book file, or create it with: node generate_excel.js --book <book>`);
        if (LOCKED_CODES.includes(e.code)) {
            throw new Error(`Workbook "${file}" is locked (${e.code}). Fix: close it in Excel (and check it is not read-only), then run this again. Nothing was changed.`);
        }
        throw e;
    }
    fsImpl.closeSync(fd);
}

module.exports = { assertWorkbookUnlocked, LOCKED_CODES };

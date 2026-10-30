// --book: one configuration file per book. See docs/setup.md ("Book file").
//   parse.js     validates the book file and the private file (every problem, with a fix)
//   resolve.js   finds a book by id (books.toml / LLC_BOOKS) or by path
//   apply.js     fills the report context exactly as the Setup sheet would
//   lists.js     the generated Lists sheet for the dropdowns
//   lock.js      the "workbook open in Excel" check
//   workbook.js  a starting workbook for a book
const { stripOption, findPositional } = require('../cli');
const { loadBook, loadPrivate } = require('./resolve');
const { parseBookText, parsePrivateText, checkBookAgainstWorkbook } = require('./parse');
const { applyBook, applyPrivate } = require('./apply');
const { writeListsSheet, listsFromBook, listsRefs, applyDropdowns, LISTS_SHEET } = require('./lists');
const { assertWorkbookUnlocked } = require('./lock');
const { buildBookWorkbook } = require('./workbook');

// Turns report.js's arguments into the ones the book implies: --book removed, the book's
// default_flags added, its workbook as the file, and its year as --year unless the command
// line gave one (the command line wins). Throws if a workbook path was also given.
function mergeBookArgs(args, book) {
    const rest = stripOption(args, '--book');
    if (findPositional(rest)) {
        throw new Error(`--book already names the workbook ("${book.workbookPath}"). Fix: remove "${findPositional(rest)}" from the command, or change "workbook" in the book file.`);
    }
    const hasYear = rest.some(a => a === '--year' || a.startsWith('--year='));
    const merged = [book.workbookPath, ...book.defaultFlags, ...rest];
    if (!hasYear && book.year) merged.push(`--year=${book.year}`);
    return merged;
}

module.exports = {
    loadBook, loadPrivate, parseBookText, parsePrivateText, checkBookAgainstWorkbook, mergeBookArgs,
    applyBook, applyPrivate, writeListsSheet, listsFromBook, listsRefs, applyDropdowns, LISTS_SHEET, assertWorkbookUnlocked,
    buildBookWorkbook,
};

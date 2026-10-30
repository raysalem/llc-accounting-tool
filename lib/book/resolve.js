// Finds and reads a book: by id through the registry (books.toml), or by path.
const fs = require('fs');
const path = require('path');
const { parse: parseToml } = require('smol-toml');
const { didYouMean } = require('./suggest');
const { parseBookText, parsePrivateText } = require('./parse');

const REGISTRY_ENV = 'LLC_BOOKS'; // environment variable holding the path of books.toml
const REGISTRY_FILE = 'books.toml'; // looked for in the current folder when LLC_BOOKS is not set
const BOOK_EXTENSION = /\.toml$/i;

// Reads the registry: { id: absolute path of the book file }. Paths are relative to the registry's folder.
function readRegistry(registryFile) {
    let raw;
    try {
        raw = parseToml(fs.readFileSync(registryFile, 'utf8'));
    } catch (e) {
        throw new Error(`Registry "${registryFile}" is not valid TOML: ${(e.message || '').split('\n')[0]}`);
    }
    const books = new Map();
    const problems = [];
    for (const [id, value] of Object.entries(raw)) {
        if (typeof value !== 'string' || !value.trim()) {
            problems.push(`"${id}" must be the path of a book file in quotes, e.g. ${id} = "${id}/${id}.book.toml".`);
        } else {
            books.set(id, path.resolve(path.dirname(registryFile), value.trim()));
        }
    }
    if (problems.length > 0) {
        throw new Error(`Registry "${registryFile}" has ${problems.length} problem(s):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }
    return books;
}

// Turns --book's value into { file, registryId }. A value ending in .toml, or an
// existing file, is a path; anything else is an id looked up in the registry.
function locateBook(ref, { env = process.env, cwd = process.cwd() } = {}) {
    const asPath = path.resolve(cwd, ref);
    if (BOOK_EXTENSION.test(ref) || (fs.existsSync(asPath) && fs.statSync(asPath).isFile())) {
        if (!fs.existsSync(asPath)) throw new Error(`Book file "${ref}" not found. Fix: check the path, or use a book id from ${REGISTRY_FILE}.`);
        return { file: asPath, registryId: null };
    }
    const registryFile = env[REGISTRY_ENV] ? path.resolve(cwd, env[REGISTRY_ENV]) : path.join(cwd, REGISTRY_FILE);
    if (!fs.existsSync(registryFile)) {
        throw new Error(`--book "${ref}" is not a book file, and there is no registry at "${registryFile}". Fix: pass the path of the book file (ending in .toml), or create ${REGISTRY_FILE} with a line such as ${ref} = "<path>/${ref}.book.toml" (set ${REGISTRY_ENV} to use a registry elsewhere).`);
    }
    const books = readRegistry(registryFile);
    if (!books.has(ref)) {
        const ids = Array.from(books.keys());
        throw new Error(`Book id "${ref}" is not in the registry "${registryFile}".${didYouMean(ref, ids)} Known ids: ${ids.length ? ids.join(', ') : '(none)'}.`);
    }
    return { file: books.get(ref), registryId: ref };
}

// Reads, validates and returns the book for --book's value. Throws one Error with every problem.
function loadBook(ref, options) {
    const { file, registryId } = locateBook(ref, options);
    if (!fs.existsSync(file)) throw new Error(`Book file "${file}" not found${registryId ? ` (listed in the registry as "${registryId}")` : ''}.`);
    const book = parseBookText(fs.readFileSync(file, 'utf8'), file);
    if (registryId && registryId !== book.id) {
        throw new Error(`Registry id "${registryId}" does not match id "${book.id}" in "${path.basename(file)}". Fix: make them the same.`);
    }
    return book;
}

// Reads and validates the book's private file. Returns null when the book names none.
// A named file that is missing is an error with the fix.
function loadPrivate(book) {
    if (!book.privatePath) return null;
    if (!fs.existsSync(book.privatePath)) {
        throw new Error(`Private file "${book.privatePath}" not found. Fix: create it (see examples/example.private.toml) or remove "private" from [book].`);
    }
    return parsePrivateText(fs.readFileSync(book.privatePath, 'utf8'), book.privatePath, book);
}

module.exports = { loadBook, loadPrivate, locateBook, readRegistry, REGISTRY_ENV, REGISTRY_FILE };

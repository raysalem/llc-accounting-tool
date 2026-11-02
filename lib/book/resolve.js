// Finds and reads a book: by id through the registry (books.toml), or by path.
const fs = require('fs');
const path = require('path');
const { parse: parseToml } = require('smol-toml');
const { didYouMean } = require('./suggest');
const { parseBookText, parsePrivateText } = require('./parse');

const REGISTRY_ENV = 'LLC_BOOKS'; // environment variable holding the path of books.toml
const REGISTRY_FILE = 'books.toml'; // looked for in the current folder when LLC_BOOKS is not set
const BOOK_EXTENSION = /\.toml$/i;

// The site file (llc-site.toml) lists every LLC and the local site's settings:
//   [site]
//   port = 5170
//   [books]
//   val = "../taxes/2026/val/valencia.book.toml"
// The older flat books.toml (one `id = "path"` line per book, no tables) is still read.
const SITE_FILE = 'llc-site.toml';
const DEFAULT_PORT = 5170;
const MIN_PORT = 1024;
const MAX_PORT = 65535;
const SITE_TOP_LEVEL_KEYS = ['site', 'books'];
const SITE_KEYS = ['port'];

// Reads a registry or site file: { file, books: Map(id -> absolute path), site: { port } }.
// Paths are relative to the file's folder. Throws one Error listing every problem.
function readSiteConfig(registryFile) {
    let raw;
    try {
        raw = parseToml(fs.readFileSync(registryFile, 'utf8'));
    } catch (e) {
        throw new Error(`Registry "${registryFile}" is not valid TOML: ${(e.message || '').split('\n')[0]}`);
    }
    const problems = [];
    const isTable = v => v && typeof v === 'object' && !Array.isArray(v);
    const siteFormat = isTable(raw.books) || isTable(raw.site);
    let entries = raw;
    const site = { port: DEFAULT_PORT };
    if (siteFormat) {
        for (const key of Object.keys(raw)) {
            if (!SITE_TOP_LEVEL_KEYS.includes(key)) {
                problems.push(`Unknown section "${key}".${didYouMean(key, SITE_TOP_LEVEL_KEYS)} Allowed: [site], [books].`);
            }
        }
        entries = isTable(raw.books) ? raw.books : {};
        if (!isTable(raw.books)) problems.push('No [books] section. Fix: add [books] with one line per LLC, e.g. val = "val/val.book.toml".');
        if (isTable(raw.site)) {
            for (const key of Object.keys(raw.site)) {
                if (!SITE_KEYS.includes(key)) problems.push(`[site] "${key}" is not a setting.${didYouMean(key, SITE_KEYS)} Allowed: ${SITE_KEYS.join(', ')}.`);
            }
            if (raw.site.port !== undefined) {
                const port = raw.site.port;
                if (typeof port !== 'number' && typeof port !== 'bigint') problems.push(`[site] port must be a number, e.g. port = ${DEFAULT_PORT} (no quotes).`);
                else if (Number(port) < MIN_PORT || Number(port) > MAX_PORT || !Number.isInteger(Number(port))) problems.push(`[site] port ${port} is out of range; use ${MIN_PORT}-${MAX_PORT}, e.g. ${DEFAULT_PORT}.`);
                else site.port = Number(port);
            }
        }
    }
    const books = new Map();
    for (const [id, value] of Object.entries(entries)) {
        if (typeof value !== 'string' || !value.trim()) {
            problems.push(`"${id}" must be the path of a book file in quotes, e.g. ${id} = "${id}/${id}.book.toml".`);
        } else {
            books.set(id, path.resolve(path.dirname(registryFile), value.trim()));
        }
    }
    if (problems.length > 0) {
        throw new Error(`Registry "${registryFile}" has ${problems.length} problem(s):\n${problems.map(p => `  - ${p}`).join('\n')}`);
    }
    return { file: path.resolve(registryFile), books, site };
}

// Reads the registry: { id: absolute path of the book file }.
function readRegistry(registryFile) {
    return readSiteConfig(registryFile).books;
}

// The registry file to use: LLC_BOOKS if set, else llc-site.toml, else books.toml in cwd.
// Returns the path even when the file does not exist (callers report it).
function findRegistryFile({ env = process.env, cwd = process.cwd() } = {}) {
    if (env[REGISTRY_ENV]) return path.resolve(cwd, env[REGISTRY_ENV]);
    const site = path.join(cwd, SITE_FILE);
    return fs.existsSync(site) ? site : path.join(cwd, REGISTRY_FILE);
}

// Turns --book's value into { file, registryId }. A value ending in .toml, or an
// existing file, is a path; anything else is an id looked up in the registry.
function locateBook(ref, { env = process.env, cwd = process.cwd() } = {}) {
    const asPath = path.resolve(cwd, ref);
    if (BOOK_EXTENSION.test(ref) || (fs.existsSync(asPath) && fs.statSync(asPath).isFile())) {
        if (!fs.existsSync(asPath)) throw new Error(`Book file "${ref}" not found. Fix: check the path, or use a book id from ${REGISTRY_FILE}.`);
        return { file: asPath, registryId: null };
    }
    const registryFile = findRegistryFile({ env, cwd });
    if (!fs.existsSync(registryFile)) {
        throw new Error(`--book "${ref}" is not a book file, and there is no registry at "${registryFile}". Fix: pass the path of the book file (ending in .toml), or create ${SITE_FILE} with [books] and a line such as ${ref} = "<path>/${ref}.book.toml" (an older flat ${REGISTRY_FILE} also works) (set ${REGISTRY_ENV} to use a registry elsewhere).`);
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

module.exports = {
    loadBook, loadPrivate, locateBook, readRegistry, readSiteConfig, findRegistryFile,
    REGISTRY_ENV, REGISTRY_FILE, SITE_FILE, DEFAULT_PORT,
};

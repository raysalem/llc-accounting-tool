// The local site for all LLCs: node site.js [path to llc-site.toml]
// Without a path: LLC_BOOKS, else llc-site.toml in the current folder.
// Opens http://localhost:<port>/ (127.0.0.1 only). Ctrl+C stops it.
const { exec } = require('child_process');
const fs = require('fs');
const book = require('./lib/book');
const { Site } = require('./lib/site/server');

async function main() {
    const argv = process.argv.slice(2);
    if (argv.includes('--help')) {
        console.log('Usage: node site.js [llc-site.toml] [--no-open]\n'
            + '  Serves every LLC in llc-site.toml at http://localhost:<port>/ (this PC only).\n'
            + '  Recomputes an LLC when its workbook, book file or opening balances are saved.');
        return;
    }
    const file = argv.find(a => !a.startsWith('--')) || book.findRegistryFile();
    if (!fs.existsSync(file)) {
        console.error(`No site file at "${file}". Create llc-site.toml:\n\n[site]\nport = 5170\n\n[books]\nval = "path/to/val.book.toml"\n`);
        process.exit(1);
    }
    const site = new Site(file);
    const url = await site.start();
    console.log(`LLC books: ${url}  (${site.books.size} LLC${site.books.size === 1 ? '' : 's'} from ${file}). Ctrl+C to stop.`);
    console.log(`Version ID ${site.versionId}`);
    if (!argv.includes('--no-open') && process.platform === 'win32') exec(`start "" "${url}"`);
    process.on('SIGINT', () => { site.stop(); process.exit(0); });
}

main().catch(e => {
    console.error(`[site] ${e.message}`);
    process.exit(1);
});

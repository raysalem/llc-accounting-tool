// The site's model of each LLC: its book, the files to watch, the last good result,
// the result before it (for before/after), and the current error.
// Each recompute runs report.js in its own process (--json), so a broken book or a
// half-saved workbook cannot take the site down, and report.js's globals stay per run.
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const book = require('../book');

const ROOT = path.join(__dirname, '..', '..');
const RETRY_MS = 1500;
const MAX_TRIES = 3;
// A save still in progress (Excel writes a temp file, then renames it) reads as a broken file.
const PARTIAL_SAVE = /Error reading file|EBUSY|EPERM|end of central directory|Corrupted zip|invalid signature|ENOENT/i;

const TMP_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-site-'));

// What counts as "a change": the reports and checks, not the stamp, run time or log.
function signature(d) {
    return JSON.stringify([d.pl, d.assets, d.liabilities, d.equity, d.vendors, d.customers, d.review1099, d.checks, d.taxYear]);
}

class Book {
    constructor(id, bookFile) {
        this.id = id;
        this.bookFile = bookFile;
        this.name = id;
        this.files = [bookFile];  // watched; filled in by load()
        this.config = null;
        this.current = null;      // last good result (report.js --json)
        this.previous = null;     // the good result before it: the "before" of before/after
        this.error = null;        // { kind: 'config'|'report', message }
        this.reportExit = null;
        this.running = false;
        this.queued = false;
        this.updated = null;
    }

    // Reads the book file. A bad book keeps the last good result and sets the error.
    load() {
        try {
            this.config = book.loadBook(this.bookFile);
            this.name = this.config.name || this.id;
            this.files = [this.bookFile, this.config.workbookPath, this.config.openingBalancesPath, this.config.privatePath]
                .filter(Boolean).map(f => path.resolve(f));
            return true;
        } catch (e) {
            this.error = { kind: 'config', message: e.message };
            return false;
        }
    }

    runOnce() {
        const out = path.join(TMP_DIR, `${this.id}.json`);
        try { fs.rmSync(out, { force: true }); } catch { /* not there */ }
        // --1099: the 1099 review needs the private file to say "tax ID on file" (the value never leaves report.js).
        const args = [path.join(ROOT, 'report.js'), '--book', this.bookFile, '--all', '--1099', '--json', out];
        return new Promise(resolve => {
            const child = spawn(process.execPath, args, { cwd: path.dirname(this.bookFile), env: { ...process.env, NO_COLOR: '1' } });
            let text = '';
            child.stdout.on('data', d => { text += d; });
            child.stderr.on('data', d => { text += d; });
            child.on('close', code => {
                let data = null;
                try { data = JSON.parse(fs.readFileSync(out, 'utf8')); } catch { /* no result */ }
                try { fs.rmSync(out, { force: true }); } catch { /* ignore */ }
                resolve({ code, text, data });
            });
        });
    }

    // Recomputes; returns true when something changed for the page.
    async recompute() {
        if (this.running) { this.queued = true; return false; }
        this.running = true;
        this.error = null;
        if (this.load()) {
            let r = null;
            for (let attempt = 1; attempt <= MAX_TRIES; attempt++) {
                r = await this.runOnce();
                if (r.data || !PARTIAL_SAVE.test(r.text)) break;
                await new Promise(res => setTimeout(res, RETRY_MS));
            }
            this.reportExit = r.code;
            if (r.data) {
                // "Before" moves only when the numbers or checks changed, so a recompute with the
                // same result (a book comment edited, a save with no change) keeps the last real change.
                if (this.current && signature(this.current) !== signature(r.data)) this.previous = this.current;
                this.current = r.data;
                this.updated = new Date().toISOString();
            } else {
                // The report stopped before it had results (e.g. a sheet the book names is missing).
                const lines = r.text.split(/\r?\n/).filter(l => /error|\[!\]|not found|missing|fix/i.test(l));
                this.error = { kind: 'report', message: (lines.length ? lines : r.text.split(/\r?\n/).slice(-15)).join('\n') };
            }
        }
        this.running = false;
        if (this.queued) { this.queued = false; return this.recompute(); }
        return true;
    }

    summary() {
        const c = this.current;
        return {
            id: this.id,
            name: this.name,
            running: this.running,
            updated: this.updated,
            error: this.error,
            hasResult: Boolean(c),
            ok: Boolean(c && c.ok) && !this.error,
            taxYear: c ? c.taxYear : (this.config && this.config.year) || null,
            netIncome: c ? c.netIncome : null,
            failedChecks: c ? c.checks.filter(k => !k.ok).map(k => ({ title: k.title, count: k.items.length })) : [],
            files: this.files.map(f => path.basename(f)),
        };
    }
}

module.exports = { Book, TMP_DIR };

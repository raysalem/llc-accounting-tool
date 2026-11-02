// The local site: one page for every LLC listed in llc-site.toml.
// Listens on 127.0.0.1 only and answers only requests addressed to localhost, so no
// other machine (and no other web site, through DNS rebinding) can read the books.
const fs = require('fs');
const http = require('http');
const path = require('path');
const book = require('../book');
const { Book } = require('./books');

const DEBOUNCE_MS = 1000;
const PUBLIC = path.join(__dirname, 'public');
const STATIC = { '/': 'index.html', '/app.js': 'app.js', '/style.css': 'style.css' };
// The FAQ page shows the repo's docs/faq.md (one copy; read on each request so edits show at once).
const { versionId } = require('../version');
const { readGitCommit } = require('../report/stamp');
const PKG_VERSION = JSON.parse(fs.readFileSync(path.join(__dirname, '..', '..', 'package.json'), 'utf8')).version;
// Version ID of the running site, and of the code on disk now (re-read at most every 5 s).
const siteVersion = () => versionId(PKG_VERSION, readGitCommit().commit);
let diskCheck = { at: 0, id: null };
function diskVersion() {
    if (Date.now() - diskCheck.at > 5000) diskCheck = { at: Date.now(), id: siteVersion() };
    return diskCheck.id;
}
const FAQ_FILE = path.join(__dirname, '..', '..', 'docs', 'faq.md');
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8' };

class Site {
    constructor(siteFile, { log = console.log } = {}) {
        this.siteFile = path.resolve(siteFile);
        this.log = log;
        this.versionId = siteVersion();  // the code this site.js is running
        this.books = new Map();      // id -> Book
        this.siteError = null;
        this.port = 5170;            // replaced by [site] port when the file is read
        this.version = 0;            // bumped on every change; pages listen through /api/events
        this.clients = new Set();
        this.watchers = new Map();   // folder -> fs.FSWatcher
        this.timers = new Map();
    }

    // Reads llc-site.toml; adds and removes books. A bad file keeps the current books.
    readSite() {
        try {
            const cfg = book.readSiteConfig(this.siteFile);
            this.siteError = null;
            this.port = cfg.site.port;
            for (const id of [...this.books.keys()]) {
                if (!cfg.books.has(id) || cfg.books.get(id) !== this.books.get(id).bookFile) this.books.delete(id);
            }
            const added = [];
            for (const [id, file] of cfg.books) {
                if (!this.books.has(id)) { this.books.set(id, new Book(id, file)); added.push(id); }
            }
            return added;
        } catch (e) {
            this.siteError = e.message;
            return [];
        }
    }

    changed() {
        this.version++;
        const msg = `data: ${this.version}\n\n`;
        for (const res of this.clients) res.write(msg);
    }

    async recompute(id) {
        const b = this.books.get(id);
        if (!b) return;
        this.changed(); // shows "updating"
        await b.recompute();
        this.rewatch();
        this.changed();
    }

    // Watches the folder of every file in use (Excel replaces the file on save, which a file watcher loses).
    rewatch() {
        const wanted = new Map(); // folder -> Map(lower file name -> [book ids] | 'site')
        const want = (file, who) => {
            const dir = path.dirname(file);
            if (!wanted.has(dir)) wanted.set(dir, new Map());
            const names = wanted.get(dir);
            const key = path.basename(file).toLowerCase();
            if (!names.has(key)) names.set(key, new Set());
            names.get(key).add(who);
        };
        want(this.siteFile, '\u0000site');
        for (const b of this.books.values()) b.files.forEach(f => want(f, b.id));
        this.wanted = wanted;
        for (const [dir, w] of this.watchers) {
            if (!wanted.has(dir)) { w.close(); this.watchers.delete(dir); }
        }
        for (const dir of wanted.keys()) {
            if (this.watchers.has(dir) || !fs.existsSync(dir)) continue;
            try {
                const w = fs.watch(dir, (event, name) => this.onFileEvent(dir, name));
                w.on('error', () => { w.close(); this.watchers.delete(dir); });
                this.watchers.set(dir, w);
            } catch (e) {
                this.log(`[site] cannot watch ${dir}: ${e.message}`);
            }
        }
    }

    onFileEvent(dir, name) {
        if (!name) return;
        const names = this.wanted && this.wanted.get(dir);
        const who = names && names.get(String(name).toLowerCase());
        if (!who) return;
        for (const id of who) {
            clearTimeout(this.timers.get(id));
            this.timers.set(id, setTimeout(() => {
                if (id === '\u0000site') {
                    const added = this.readSite();
                    this.rewatch();
                    this.changed();
                    added.forEach(a => this.recompute(a));
                } else {
                    this.recompute(id);
                }
            }, DEBOUNCE_MS));
        }
    }

    state() {
        return {
            version: this.version,
            versionId: this.versionId,
            // The code changed after site.js started: the page asks for a restart.
            codeChanged: diskVersion() !== this.versionId ? diskVersion() : null,
            siteFile: this.siteFile,
            siteError: this.siteError,
            books: [...this.books.values()].map(b => b.summary()),
        };
    }

    bookDetail(id) {
        const b = this.books.get(id);
        if (!b) return null;
        return { ...b.summary(), current: b.current, previous: b.previous };
    }

    // Only requests made to this machine by name or loopback address are answered.
    allowedHost(req) {
        const host = String(req.headers.host || '').toLowerCase();
        return host === `localhost:${this.port}` || host === `127.0.0.1:${this.port}`;
    }

    handle(req, res) {
        if (!this.allowedHost(req)) { res.writeHead(403); res.end('Forbidden'); return; }
        const url = new URL(req.url, `http://127.0.0.1:${this.port}`);
        const send = (code, body, type = 'application/json; charset=utf-8') => {
            res.writeHead(code, { 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
            res.end(typeof body === 'string' ? body : JSON.stringify(body));
        };
        if (req.method === 'GET' && STATIC[url.pathname]) {
            const file = path.join(PUBLIC, STATIC[url.pathname]);
            return send(200, fs.readFileSync(file, 'utf8'), TYPES[path.extname(file)]);
        }
        if (req.method === 'GET' && url.pathname === '/api/faq') return send(200, fs.readFileSync(FAQ_FILE, 'utf8'), 'text/markdown; charset=utf-8');
        if (req.method === 'GET' && url.pathname === '/api/state') return send(200, this.state());
        if (req.method === 'GET' && url.pathname === '/api/events') {
            res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-store', Connection: 'keep-alive' });
            res.write(`data: ${this.version}\n\n`);
            this.clients.add(res);
            req.on('close', () => this.clients.delete(res));
            return undefined;
        }
        const m = url.pathname.match(/^\/api\/book\/([^/]+)(\/refresh)?$/);
        if (m) {
            const id = decodeURIComponent(m[1]);
            if (!this.books.has(id)) return send(404, { error: `No LLC "${id}" in ${path.basename(this.siteFile)}` });
            if (m[2]) {
                if (req.method !== 'POST') return send(405, { error: 'POST only' });
                this.recompute(id);
                return send(202, { ok: true });
            }
            return send(200, this.bookDetail(id));
        }
        return send(404, { error: 'Not found' });
    }

    async start() {
        this.readSite();
        if (this.siteError && this.books.size === 0) throw new Error(this.siteError);
        this.rewatch();
        this.server = http.createServer((req, res) => {
            try { this.handle(req, res); } catch (e) {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
                res.end(e.message);
            }
        });
        await new Promise((resolve, reject) => {
            this.server.once('error', reject);
            this.server.listen(this.port, '127.0.0.1', resolve);
        });
        // Keep-alive comments so idle event streams are not closed by the browser.
        this.ping = setInterval(() => { for (const r of this.clients) r.write(': ping\n\n'); }, 20000);
        await Promise.all([...this.books.keys()].map(id => this.recompute(id)));
        return `http://localhost:${this.port}/`;
    }

    stop() {
        clearInterval(this.ping);
        for (const w of this.watchers.values()) w.close();
        for (const r of this.clients) r.end();
        if (this.server) this.server.close();
    }
}

module.exports = { Site };

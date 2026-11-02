// The local site (site.js): reads llc-site.toml, computes every LLC, serves state and
// book data on 127.0.0.1, refuses other Host names, keeps the last good result when a
// book breaks, and moves "before" only when the numbers change. Fake books only.
const { spawnSync } = require('child_process');
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const { Site } = require('../lib/site/server');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'llc-site-test-'));
const PORT = 5190 + Math.floor(Math.random() * 50);
let failures = 0;

function check(label, ok, detail = '') {
    if (ok) console.log(`✅ [PASS] ${label}`);
    else { console.error(`❌ [FAIL] ${label}${detail ? ` — ${detail}` : ''}`); failures++; }
}

function get(urlPath, host = `localhost:${PORT}`) {
    return new Promise((resolve, reject) => {
        const req = http.request({ host: '127.0.0.1', port: PORT, path: urlPath, headers: { Host: host } }, res => {
            let body = '';
            res.on('data', d => { body += d; });
            res.on('end', () => resolve({ status: res.statusCode, body }));
        });
        req.on('error', reject);
        req.end();
    });
}

function makeBook(id, name) {
    const dir = path.join(TMP, id);
    fs.mkdirSync(dir);
    const text = fs.readFileSync(path.join(ROOT, 'examples/example.book.toml'), 'utf8')
        .replace('id = "example"', `id = "${id}"`)
        .replace('name = "Example Holdings LLC"', `name = "${name}"`)
        .replace('Example_2026.xlsx', `${id}.xlsx`);
    const file = path.join(dir, `${id}.book.toml`);
    fs.writeFileSync(file, text);
    fs.copyFileSync(path.join(ROOT, 'examples/example.private.toml'), path.join(dir, 'example.private.toml'));
    const r = spawnSync(process.execPath, [path.join(ROOT, 'generate_excel.js'), '--book', file], { encoding: 'utf8' });
    if (r.status !== 0) throw new Error(r.stdout + r.stderr);
    return file;
}

async function main() {
    const acme = makeBook('acme', 'Acme LLC');
    makeBook('beta', 'Beta LLC');
    const siteFile = path.join(TMP, 'llc-site.toml');
    fs.writeFileSync(siteFile, `[site]\nport = ${PORT}\n\n[books]\nacme = "acme/acme.book.toml"\nbeta = "beta/beta.book.toml"\n`);

    const site = new Site(siteFile, { log: () => {} });
    try {
        const url = await site.start();
        check('Site starts on the configured port', url === `http://localhost:${PORT}/`, url);

        const state = JSON.parse((await get('/api/state')).body);
        check('State lists both LLCs', state.books.map(b => b.id).join(',') === 'acme,beta', JSON.stringify(state.books.map(b => b.id)));
        check('State has the site Version ID (version-branch-commit-codehash)', /^\d+\.\d+\.\d+-.+-[0-9a-z]+-[0-9a-f]{6}$/.test(state.versionId) && state.codeChanged === null, state.versionId);
        check('Both LLCs computed and pass their checks', state.books.every(b => b.hasResult && b.ok), JSON.stringify(state.books));

        check('Page served', (await get('/')).status === 200);
        check('Another Host name is refused (DNS rebinding guard)', (await get('/api/state', `evil.example:${PORT}`)).status === 403);
        check('127.0.0.1 Host is allowed', (await get('/api/state', `127.0.0.1:${PORT}`)).status === 200);
        const faq = await get('/api/faq');
        check('FAQ served from docs/faq.md', faq.status === 200 && faq.body.includes('How-to: set starting (opening) balances'));
        check('Unknown LLC is a 404', (await get('/api/book/nope')).status === 404);

        const detail = JSON.parse((await get('/api/book/acme')).body);
        check('Book detail has reports, checks and 1099 review', detail.current && Array.isArray(detail.current.pl) && Array.isArray(detail.current.checks) && Array.isArray(detail.current.review1099));
        check('No earlier result before any change', detail.previous === null);

        // Same numbers again (a comment added): "before" does not move.
        fs.appendFileSync(acme, '\n# a comment\n');
        await site.recompute('acme');
        check('A recompute with the same numbers keeps "before" empty', site.books.get('acme').previous === null);

        // A broken book: error shown, last good result kept.
        const good = fs.readFileSync(acme, 'utf8');
        fs.writeFileSync(acme, good.replace('type = "income"', 'type = "incom"'));
        await site.recompute('acme');
        const broken = site.books.get('acme');
        check('Broken book: config error with a suggested fix', broken.error && broken.error.kind === 'config' && broken.error.message.includes('Did you mean "income"'), JSON.stringify(broken.error));
        check('Broken book: last good result still shown', broken.current !== null);
        const other = JSON.parse((await get('/api/state')).body).books.find(b => b.id === 'beta');
        check('Broken book: the other LLC is unaffected', other.ok && !other.error);
        fs.writeFileSync(acme, good);
        await site.recompute('acme');
        check('Fixed book: error cleared', site.books.get('acme').error === null);

        // Site file edit: an LLC removed.
        fs.writeFileSync(siteFile, `[site]\nport = ${PORT}\n\n[books]\nacme = "acme/acme.book.toml"\n`);
        site.readSite();
        check('Site file reload removes an LLC', [...site.books.keys()].join(',') === 'acme');
        fs.writeFileSync(siteFile, '[site]\nport = "x"\n');
        site.readSite();
        check('Bad site file: error shown, current LLCs kept', site.siteError && site.books.has('acme'), site.siteError);
    } finally {
        site.stop();
    }

    if (!process.env.KEEP_TMP) fs.rmSync(TMP, { recursive: true, force: true });
    if (failures > 0) {
        console.error(`\n${failures} check(s) failed.`);
        process.exit(1);
    }
    console.log('\nAll site tests passed.');
    process.exit(0);
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});

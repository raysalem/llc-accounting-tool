// Watch mode: reruns the report every time you save the workbook (or the book,
// opening-balances or vendor file) and shows what changed since the last run,
// as a redline or side by side.
//
//   node watch.js --book val --all
//   node watch.js "My_Books_2026.xlsx" --pl --bs
//
// Takes the same arguments as report.js. It is print-only: --save is refused, and
// the report only ever reads the workbook, so the books are never changed.
// Keys: v redline/side by side · f changes only/full · b pin BEFORE · p BEFORE = previous run
//       r rerun · o open in the browser · q quit
const { spawn, exec } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { parseOptionValue, findPositional } = require('./lib/cli');
const { diffLines, countChanges, renderRedline, renderSideBySide } = require('./lib/watch/diff');
const { renderHtml } = require('./lib/watch/html');

const DEBOUNCE_MS = 1000;      // Excel writes a temp file, then renames it: wait for the save to settle
const RETRY_MS = 1500;         // a save still in progress reads as a broken file: try again
const MAX_RETRIES = 3;
const PARTIAL_SAVE = /Error reading file|EBUSY|end of central directory|Corrupted zip|invalid signature/i;

const args = process.argv.slice(2);
if (args.includes('--help') || args.length === 0) {
    console.log(fs.readFileSync(__filename, 'utf8').split('\n').filter(l => l.startsWith('//')).map(l => l.slice(3)).join('\n'));
    process.exit(0);
}
if (args.includes('--save')) {
    console.error('watch.js is print-only and never writes the workbook. Remove --save (run report.js --save once you are done).');
    process.exit(1);
}

// The files whose save triggers a rerun.
function watchedFiles() {
    const files = [];
    const ref = parseOptionValue(args, '--book');
    if (ref) {
        try {
            const book = require('./lib/book').loadBook(ref); // reads the book only, never the private file
            files.push(book.file, book.workbookPath);
            if (book.openingBalancesPath) files.push(book.openingBalancesPath);
        } catch (e) {
            // A broken book still gets watched, so fixing it reruns the report.
            if (ref.endsWith('.toml')) files.push(ref);
            console.error(`[watch] Could not read the book yet: ${e.message}`);
        }
    } else {
        files.push(findPositional(args) || 'LLC_Accounting_Template.xlsx');
    }
    for (const flag of ['--opening-balances', '--vendor-file']) {
        const v = parseOptionValue(args, flag);
        if (v) files.push(v);
    }
    return [...new Set(files.filter(Boolean).map(f => path.resolve(f)))];
}

const files = watchedFiles();
const runs = [];           // { n, time, lines, code }
let pinned = null;         // a run kept as BEFORE until p is pressed
let view = 'redline';
let full = false;
let htmlPath = null;       // set when o is pressed; the page is rewritten after every run
let running = false;
let queued = false;

function runReport(attempt = 1) {
    return new Promise(resolve => {
        const child = spawn(process.execPath, [path.join(__dirname, 'report.js'), ...args], {
            cwd: process.cwd(), env: { ...process.env, FORCE_COLOR: '0', NO_COLOR: '1' },
        });
        let out = '';
        child.stdout.on('data', d => { out += d; });
        child.stderr.on('data', d => { out += d; });
        child.on('close', code => resolve({ out, code }));
    }).then(r => {
        if (r.code !== 0 && PARTIAL_SAVE.test(r.out) && attempt < MAX_RETRIES) {
            return new Promise(res => setTimeout(res, RETRY_MS)).then(() => runReport(attempt + 1));
        }
        return r;
    });
}

function label(run, isPinned) {
    return `run #${run.n} ${run.time}${isPinned ? ' (pinned)' : ''}${run.code ? ` exit ${run.code}` : ''}`;
}

function show() {
    const after = runs[runs.length - 1];
    if (!after) return;
    const before = pinned || runs[runs.length - 2] || null;
    process.stdout.write('\x1b[2J\x1b[3J\x1b[H'); // clear screen and scrollback
    const watchedNames = files.map(f => path.basename(f)).join(', ');
    console.log(`\x1b[33m[watch] ${watchedNames}  ·  v view · f full · b pin · p previous · r rerun · o browser · q quit\x1b[0m`);
    if (!before) {
        console.log(`[watch] ${label(after)}: first run, nothing to compare yet. Edit and save to see changes.\n`);
        console.log(after.lines.join('\n'));
    } else {
        const ops = diffLines(before.lines, after.lines);
        const { removed, added } = countChanges(ops);
        console.log(`[watch] BEFORE ${label(before, before === pinned)}  →  AFTER ${label(after)}  ·  `
            + (removed + added ? `\x1b[31m-${removed}\x1b[0m \x1b[32m+${added}\x1b[0m lines` : 'no changes') + '\n');
        const opts = { full, columns: process.stdout.columns || 160 };
        const lines = view === 'redline' ? renderRedline(ops, opts) : renderSideBySide(ops, opts);
        if (!full && removed + added === 0) console.log('(no changes; press f for the full report)');
        else console.log(lines.join('\n'));
        writeHtml(before, after, ops);
    }
    if (after.code) console.log(`\n\x1b[31m[watch] The report stopped with errors (exit ${after.code}); see above.\x1b[0m`);
}

function writeHtml(before, after, ops) {
    if (!htmlPath) return;
    const page = renderHtml({
        title: `LLC watch: ${path.basename(files[0] || 'report')}`,
        ops: ops || diffLines(before ? before.lines : [], after.lines),
        beforeLabel: before ? `BEFORE ${label(before, before === pinned)}` : 'BEFORE (none yet)',
        afterLabel: `AFTER ${label(after)}`,
        status: `updated ${after.time}`,
    });
    fs.writeFileSync(htmlPath, page);
}

async function rerun() {
    if (running) { queued = true; return; }
    running = true;
    const { out, code } = await runReport();
    runs.push({ n: runs.length + 1, time: new Date().toLocaleTimeString(), lines: out.replace(/\r/g, '').trimEnd().split('\n'), code });
    if (runs.length > 50) runs.splice(0, runs.length - 50); // keep memory bounded; a pinned run is held separately
    running = false;
    show();
    if (queued) { queued = false; rerun(); }
}

// Watch each folder (Excel replaces the file on save, which a file watcher can lose).
let timer = null;
const byDir = new Map();
files.forEach(f => {
    const dir = path.dirname(f);
    if (!byDir.has(dir)) byDir.set(dir, new Set());
    byDir.get(dir).add(path.basename(f).toLowerCase());
});
for (const [dir, names] of byDir) {
    if (!fs.existsSync(dir)) { console.error(`[watch] Folder not found: ${dir}`); continue; }
    fs.watch(dir, (event, name) => {
        if (!name || !names.has(String(name).toLowerCase())) return;
        clearTimeout(timer);
        timer = setTimeout(rerun, DEBOUNCE_MS);
    });
}

if (process.stdin.isTTY) {
    process.stdin.setRawMode(true);
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', key => {
        const last = runs[runs.length - 1];
        if (key === 'q' || key === '\u0003') process.exit(0);
        else if (key === 'v') { view = view === 'redline' ? 'split' : 'redline'; show(); }
        else if (key === 'f') { full = !full; show(); }
        else if (key === 'b' && last) { pinned = last; show(); }
        else if (key === 'p') { pinned = null; show(); }
        else if (key === 'r') rerun();
        else if (key === 'o' && last) {
            const dir = path.join(os.tmpdir(), 'llc-watch');
            fs.mkdirSync(dir, { recursive: true });
            htmlPath = path.join(dir, 'watch.html');
            const before = pinned || runs[runs.length - 2] || null;
            writeHtml(before, last);
            exec(`start "" "${htmlPath}"`);
            show();
            console.log(`\n[watch] Browser view: ${htmlPath} (local file, updates on every run)`);
        }
    });
}

rerun();

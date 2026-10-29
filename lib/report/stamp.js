// The report stamp: what a report was made from. buildStamp() makes one object
// per run; the console header, the --save Report Info sheet and the PDF title
// page all render that same object through stampRows().
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const REPO_ROOT = path.join(__dirname, '..', '..');
const HASH_HEX_LENGTH = 12; // First 12 hex digits of the SHA-256
const MINUTES_PER_HOUR = 60;
const PAD_WIDTH = 2;
const UNKNOWN = 'unknown';

// Reads a git value from the tool's own checkout. Returns { value, error }.
function git(args) {
    try {
        return { value: execFileSync('git', args, { cwd: REPO_ROOT, stdio: ['ignore', 'pipe', 'pipe'] }).toString().trim(), error: null };
    } catch (e) {
        return { value: null, error: e.message.split('\n')[0] };
    }
}

// Short commit of the tool, with "+dirty" when the working tree has changes.
// Never throws: a missing git or .git gives commit "unknown" and the reason in error.
function readGitCommit() {
    if (!fs.existsSync(path.join(REPO_ROOT, '.git'))) return { commit: UNKNOWN, error: 'no .git folder' };
    const head = git(['rev-parse', '--short', 'HEAD']);
    if (head.error) return { commit: UNKNOWN, error: head.error };
    const status = git(['status', '--porcelain']);
    if (status.error) return { commit: UNKNOWN, error: status.error };
    return { commit: status.value ? `${head.value}+dirty` : head.value, error: null };
}

// ISO 8601 time in the local time zone, e.g. 2026-10-29T14:03:05-07:00.
function localIso(date) {
    const pad = n => String(Math.abs(n)).padStart(PAD_WIDTH, '0');
    const offset = -date.getTimezoneOffset();
    const sign = offset < 0 ? '-' : '+';
    const zone = `${sign}${pad(Math.trunc(Math.abs(offset) / MINUTES_PER_HOUR))}:${pad(Math.abs(offset) % MINUTES_PER_HOUR)}`;
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
        + `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}${zone}`;
}

// Hash and size of one input file. A file that cannot be read throws.
function fingerprint(role, filePath) {
    const bytes = fs.readFileSync(filePath);
    return {
        role,
        name: path.basename(filePath),
        path: path.resolve(filePath),
        sha256: crypto.createHash('sha256').update(bytes).digest('hex').slice(0, HASH_HEX_LENGTH),
        size: bytes.length,
    };
}

// Builds the stamp. inputs is [{ role, path }] for every input file read, in order.
function buildStamp({ args, taxYear, inputs, now = new Date() }) {
    const pkg = JSON.parse(fs.readFileSync(path.join(REPO_ROOT, 'package.json'), 'utf8'));
    const { commit, error } = readGitCommit();
    return {
        tool: 'LLC Accounting Tool',
        version: pkg.version,
        gitCommit: commit,
        gitError: error,
        runTime: localIso(now),
        args: [...args],
        taxYear,
        inputs: inputs.map(i => fingerprint(i.role, i.path)),
    };
}

// The stamp as [label, value] pairs, in display order.
function stampRows(stamp) {
    const rows = [
        ['Tool Version', stamp.version],
        ['Git Commit', stamp.gitCommit],
        ['Run Time', stamp.runTime],
        ['Flags', stamp.args.join(' ')],
        ['Tax Year', String(stamp.taxYear)],
    ];
    stamp.inputs.forEach(i => {
        rows.push([`Input: ${i.role}`, `${i.name}  sha256 ${i.sha256}  ${i.size.toLocaleString('en-US')} bytes`]);
    });
    return rows;
}

// Console form of stampRows(): labels padded to line up.
function stampLines(stamp) {
    const rows = stampRows(stamp);
    const width = Math.max(...rows.map(([label]) => label.length));
    return rows.map(([label, value]) => `  ${label.padEnd(width)} : ${value}`);
}

module.exports = { buildStamp, stampRows, stampLines, localIso };

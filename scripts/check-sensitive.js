// Blocks commits that would add personal financial data to the repo.
//
//   node scripts/check-sensitive.js           check every tracked file (used in CI)
//   node scripts/check-sensitive.js --staged  check staged files (used by .githooks/pre-commit)
//
// It fails on:
//   - data files (.csv, .xlsx, .xls, .pdf, .txt, .lnk) outside tests/ and examples/
//   - network-share paths to an IP address, e.g. \\10.0.0.5\share (sensitive-ok)
//   - values that look like real Social Security numbers
// To allow a specific line, add the text "sensitive-ok" to it.
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const staged = process.argv.includes('--staged');

const DATA_FILE = /\.(csv|xlsx|xls|pdf|txt|lnk)$/i;
const ALLOWED_DATA_DIRS = /^(tests|examples)\//;
const UNC_IP_PATH = /\\\\\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\\/;
const SSN = /\b(\d{3})-(\d{2})-(\d{4})\b/g;

// SSNs never start with 000, 666 or 9xx, or have 00 / 0000 groups; those are safe placeholders.
function looksLikeRealSsn([, area, group, serial]) {
    if (area === '000' || area === '666' || area[0] === '9') return false;
    if (group === '00' || serial === '0000') return false;
    return true;
}

function listFiles() {
    const args = staged ? ['diff', '--cached', '--name-only', '--diff-filter=ACMR'] : ['ls-files'];
    return execFileSync('git', args, { cwd: ROOT, encoding: 'utf8' }).split('\n').filter(Boolean);
}

function readFile(file) {
    if (staged) return execFileSync('git', ['show', `:${file}`], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
    return fs.readFileSync(path.join(ROOT, file), 'utf8');
}

const problems = [];
for (const file of listFiles()) {
    if (DATA_FILE.test(file) && !ALLOWED_DATA_DIRS.test(file)) {
        problems.push(`${file}: data file outside tests/ or examples/`);
        continue;
    }
    if (/\.(xlsx|xls|pdf|png|jpg)$/i.test(file) || file === 'package-lock.json') continue; // binary or generated
    const lines = readFile(file).split('\n');
    lines.forEach((line, i) => {
        if (line.includes('sensitive-ok')) return;
        if (UNC_IP_PATH.test(line)) problems.push(`${file}:${i + 1}: network path to an IP address`);
        for (const m of line.matchAll(SSN)) {
            if (looksLikeRealSsn(m)) problems.push(`${file}:${i + 1}: looks like a Social Security number`);
        }
    });
}

if (problems.length) {
    console.error('Possible personal data found:');
    problems.forEach(p => console.error(`  ${p}`));
    console.error('\nRemove it, or add "sensitive-ok" to the line if it is a harmless example.');
    process.exit(1);
}
console.log(`check-sensitive: ${staged ? 'staged' : 'tracked'} files are clean.`);

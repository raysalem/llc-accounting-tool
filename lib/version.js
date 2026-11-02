// The tool's Version ID: <package version>-<branch>-<git commit>-<code hash>,
// e.g. 3.0.142-feature/robustness-557496f-9f3c1a.
// The code hash covers the tool's own code (root *.js, package.json, lib/**), so two builds with
// uncommitted changes ("+dirty") still get different IDs. It is printed on the site page and on
// every report (console stamp, Report Info sheet, each Excel sheet's footer, every PDF page).
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const REPO_ROOT = path.join(__dirname, '..');
const CODE_HASH_LENGTH = 6;
const CODE_EXT = new Set(['.js', '.html', '.css', '.json']);

// Every code file of the tool, sorted, as paths relative to the repo root (forward slashes).
function codeFiles() {
    const files = fs.readdirSync(REPO_ROOT)
        .filter(f => f.endsWith('.js') || f === 'package.json')
        .filter(f => fs.statSync(path.join(REPO_ROOT, f)).isFile());
    const walk = dir => {
        for (const e of fs.readdirSync(path.join(REPO_ROOT, dir), { withFileTypes: true })) {
            const rel = `${dir}/${e.name}`;
            if (e.isDirectory()) walk(rel);
            else if (CODE_EXT.has(path.extname(e.name))) files.push(rel);
        }
    };
    walk('lib');
    return files.sort();
}

// Hash of the tool's code. Line endings are normalized so a CRLF checkout matches an LF one.
function codeHash() {
    const h = crypto.createHash('sha256');
    for (const rel of codeFiles()) {
        h.update(rel).update('\0');
        h.update(fs.readFileSync(path.join(REPO_ROOT, rel), 'utf8').replace(/\r\n/g, '\n')).update('\0');
    }
    return h.digest('hex').slice(0, CODE_HASH_LENGTH);
}

// The working branch of the tool's checkout. A detached checkout lists every local branch that
// contains the commit, joined with "+"; no git gives "nobranch".
function gitBranch() {
    const run = args => execFileSync('git', args, { cwd: REPO_ROOT, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
    try {
        const name = run(['rev-parse', '--abbrev-ref', 'HEAD']);
        if (name && name !== 'HEAD') return name;
        const list = run(['branch', '--contains', 'HEAD', '--format=%(refname:short)'])
            .split(/\r?\n/).filter(b => b && !b.startsWith('('));
        return list.length ? list.join('+') : 'detached';
    } catch {
        return 'nobranch';
    }
}

// commit: the short git commit ("557496f" or "557496f+dirty"; "unknown" without git).
function versionId(version, commit, branch = gitBranch()) {
    const c = String(commit || 'unknown').replace(/\+dirty$/, '');
    return `${version}-${branch}-${c}-${codeHash()}`;
}

module.exports = { versionId, codeHash, codeFiles, gitBranch };

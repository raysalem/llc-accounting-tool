// Line diff for watch.js: compares two report outputs (before / after) and
// renders them as a redline or side by side. Pure functions, no I/O.

// Lines that change on every run even when the books did not (the stamp's
// run time and file hashes). They are compared as equal so they never show as edits.
const VOLATILE = [
    /^\s*Run Time\s*:/,
    /^\s*Input: .*sha256 /,
    /^\s*Git Commit\s*:/,
];

const MAX_LCS_CELLS = 25e6; // 5000 x 5000 lines; larger outputs fall back to a plain before/after

function key(line) {
    return VOLATILE.some(re => re.test(line)) ? '\u0000volatile:' + line.split(':')[0].trim() : line;
}

// Returns ops: [{ op: 'same'|'del'|'add', a?: string, b?: string }] turning before into after.
function diffLines(before, after) {
    const A = before.map(key);
    const B = after.map(key);
    const n = A.length;
    const m = B.length;
    if (n * m > MAX_LCS_CELLS) {
        return [...before.map(a => ({ op: 'del', a })), ...after.map(b => ({ op: 'add', b }))];
    }
    // LCS lengths from the end, so the walk below goes forward.
    const width = m + 1;
    const L = new Uint32Array((n + 1) * width);
    for (let i = n - 1; i >= 0; i--) {
        for (let j = m - 1; j >= 0; j--) {
            L[i * width + j] = A[i] === B[j]
                ? L[(i + 1) * width + j + 1] + 1
                : Math.max(L[(i + 1) * width + j], L[i * width + j + 1]);
        }
    }
    const ops = [];
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
        if (A[i] === B[j]) { ops.push({ op: 'same', a: before[i], b: after[j] }); i++; j++; }
        else if (L[(i + 1) * width + j] >= L[i * width + j + 1]) ops.push({ op: 'del', a: before[i++] });
        else ops.push({ op: 'add', b: after[j++] });
    }
    while (i < n) ops.push({ op: 'del', a: before[i++] });
    while (j < m) ops.push({ op: 'add', b: after[j++] });
    return ops;
}

function countChanges(ops) {
    return {
        removed: ops.filter(o => o.op === 'del').length,
        added: ops.filter(o => o.op === 'add').length,
    };
}

// Side-by-side rows: [{ left, right, kind: 'same'|'change'|'del'|'add' }].
// A run of deletions followed by additions is paired row by row (a changed line).
function sideBySide(ops) {
    const rows = [];
    for (let k = 0; k < ops.length;) {
        if (ops[k].op === 'same') { rows.push({ left: ops[k].a, right: ops[k].b, kind: 'same' }); k++; continue; }
        const dels = [];
        const adds = [];
        while (k < ops.length && ops[k].op === 'del') dels.push(ops[k++].a);
        while (k < ops.length && ops[k].op === 'add') adds.push(ops[k++].b);
        for (let r = 0; r < Math.max(dels.length, adds.length); r++) {
            const left = r < dels.length ? dels[r] : null;
            const right = r < adds.length ? adds[r] : null;
            rows.push({ left, right, kind: left !== null && right !== null ? 'change' : left !== null ? 'del' : 'add' });
        }
    }
    return rows;
}

// Keeps only rows within `context` rows of a change; gaps become { gap: count }.
function onlyChanges(rows, isChange, context = 3) {
    const keep = new Array(rows.length).fill(false);
    rows.forEach((r, idx) => {
        if (!isChange(r)) return;
        for (let k = Math.max(0, idx - context); k <= Math.min(rows.length - 1, idx + context); k++) keep[k] = true;
    });
    const out = [];
    let gap = 0;
    rows.forEach((r, idx) => {
        if (keep[idx]) {
            if (gap) out.push({ gap });
            gap = 0;
            out.push(r);
        } else gap++;
    });
    if (gap) out.push({ gap });
    return out;
}

const ANSI = { red: '\x1b[31m', green: '\x1b[32m', dim: '\x1b[2m', yellow: '\x1b[33m', reset: '\x1b[0m', strike: '\x1b[9m' };

// Console redline: removed lines red and struck through, added lines green.
function renderRedline(ops, { full = false, color = true } = {}) {
    const c = color ? ANSI : Object.fromEntries(Object.keys(ANSI).map(k => [k, '']));
    const rows = full ? ops : onlyChanges(ops, o => o.op !== 'same');
    return rows.map(r => {
        if (r.gap) return `${c.dim}   ... ${r.gap} unchanged line${r.gap === 1 ? '' : 's'} ...${c.reset}`;
        if (r.op === 'same') return `${c.dim}  ${c.reset}${r.b}`;
        if (r.op === 'del') return `${c.red}- ${c.strike}${r.a}${c.reset}`;
        return `${c.green}+ ${r.b}${c.reset}`;
    });
}

function fit(text, width) {
    const s = text === null ? '' : text;
    return s.length > width ? s.slice(0, width - 1) + '…' : s.padEnd(width);
}

// Console split screen: BEFORE on the left, AFTER on the right.
function renderSideBySide(ops, { full = false, color = true, columns = 160 } = {}) {
    const c = color ? ANSI : Object.fromEntries(Object.keys(ANSI).map(k => [k, '']));
    const w = Math.max(20, Math.floor((columns - 3) / 2));
    const all = sideBySide(ops);
    const rows = full ? all : onlyChanges(all, r => r.kind !== 'same');
    const head = `${c.yellow}${fit('BEFORE', w)} | AFTER${c.reset}`;
    return [head, '-'.repeat(w) + '-+-' + '-'.repeat(w), ...rows.map(r => {
        if (r.gap) return `${c.dim}${fit(`   ... ${r.gap} unchanged ...`, w)} |${c.reset}`;
        if (r.kind === 'same') return `${fit(r.left, w)} | ${fit(r.right, w)}`;
        const left = r.left === null ? fit('', w) : `${c.red}${fit(r.left, w)}${c.reset}`;
        const right = r.right === null ? '' : `${c.green}${fit(r.right, w)}${c.reset}`;
        return `${left} | ${right}`;
    })];
}

module.exports = { diffLines, countChanges, sideBySide, onlyChanges, renderRedline, renderSideBySide, VOLATILE };

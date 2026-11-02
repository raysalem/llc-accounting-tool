// watch.js diff: before/after lines, volatile stamp lines, redline and side-by-side rendering.
const assert = require('assert');
const { diffLines, countChanges, sideBySide, renderRedline, renderSideBySide } = require('../lib/watch/diff');
const { renderHtml } = require('../lib/watch/html');

const before = [
    'Report stamp:',
    '  Run Time            : 2026-11-01T10:00:00-07:00',
    '  Input: Workbook     : Val.xlsx  sha256 aaaaaaaaaaaa  100 bytes',
    'Rent Income      1,000.00',
    'Repairs            200.00',
    'Net Income         800.00',
];
const after = [
    'Report stamp:',
    '  Run Time            : 2026-11-01T10:05:00-07:00',
    '  Input: Workbook     : Val.xlsx  sha256 bbbbbbbbbbbb  104 bytes',
    'Rent Income      1,000.00',
    'Repairs            250.00',
    'Net Income         750.00',
    '[OK] (A = L + E)',
];

// Identical outputs: no changes.
assert.deepStrictEqual(countChanges(diffLines(before, before)), { removed: 0, added: 0 });

// Run time and file hash lines never count as changes.
const ops = diffLines(before, after);
assert.deepStrictEqual(countChanges(ops), { removed: 2, added: 3 });
assert.ok(ops.filter(o => o.op !== 'same').every(o => !/Run Time|sha256/.test(o.a || o.b)));

// Side by side pairs the changed lines on one row.
const rows = sideBySide(ops);
const repairs = rows.find(r => r.left && r.left.startsWith('Repairs'));
assert.strictEqual(repairs.kind, 'change');
assert.ok(repairs.right.includes('250.00'));
assert.ok(rows.some(r => r.kind === 'add' && r.right === '[OK] (A = L + E)'));

// Redline without color: '-' for removed, '+' for added; changes-only keeps context.
const red = renderRedline(ops, { color: false });
assert.ok(red.includes('- Repairs            200.00'));
assert.ok(red.includes('+ Repairs            250.00'));
assert.ok(red.includes('+ [OK] (A = L + E)'));

// Changes-only collapses long unchanged runs.
const long = Array.from({ length: 40 }, (_, i) => `line ${i}`);
const long2 = [...long];
long2[20] = 'line 20 changed';
const collapsed = renderRedline(diffLines(long, long2), { color: false });
assert.ok(collapsed.some(l => /\.\.\. 17 unchanged lines/.test(l)));
assert.ok(collapsed.length < 15);
assert.strictEqual(renderRedline(diffLines(long, long2), { color: false, full: true }).length, 41);

// Split screen: BEFORE | AFTER header, lines truncated to the column width.
const split = renderSideBySide(ops, { color: false, columns: 60, full: true });
assert.ok(split[0].startsWith('BEFORE'));
assert.ok(split.every(l => l.length <= 60));

// HTML escapes the report text.
const html = renderHtml({ title: 't', ops: diffLines(['<a>'], ['<b>']), beforeLabel: 'B', afterLabel: 'A', status: 's' });
assert.ok(html.includes('&lt;a&gt;') && !html.includes('<a>'));

console.log('test_watch_diff: all checks passed');

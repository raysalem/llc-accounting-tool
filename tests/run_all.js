// Runs every test file and exits non-zero if any of them fail.
// Usage: npm test
const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const tests = [
    'tests/test_accounting_lib.js',
    'tests/run_integration_test.js',
    ...fs.readdirSync(__dirname)
        .filter(f => /^test_.*\.js$/.test(f) && f !== 'test_accounting_lib.js')
        .sort()
        .map(f => `tests/${f}`),
];

const results = [];
for (const test of tests) {
    console.log(`\n===== ${test} =====`);
    const r = spawnSync(process.execPath, [test], { cwd: root, stdio: 'inherit' });
    results.push({ test, ok: r.status === 0, code: r.status });
}

console.log('\n===== SUMMARY =====');
results.forEach(r => console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.test}${r.ok ? '' : ` (exit ${r.code})`}`));

const failed = results.filter(r => !r.ok).length;
if (failed > 0) {
    console.error(`\n${failed} of ${results.length} test files failed.`);
    process.exit(1);
}
console.log(`\nAll ${results.length} test files passed.`);

const { execSync } = require('child_process');

console.log('--- TEST: --details Flag Support ---');

try {
    // Requires the test file "tests/temp_target.xlsx" to exist (created in previous tests)
    // We will query for 'Office Supplies' which should exist in the standard template logic
    const cmd = `node update_financials.js "tests/temp_target.xlsx" --details "office supplies"`;
    console.log(`Running: ${cmd}`);
    const output = execSync(cmd).toString();
    console.log(output);

    if (output.includes('DETAILS: "office supplies"') && output.includes('TOTAL')) {
        console.log('✅ PASS: Details output detected with Total.');
    } else {
        console.error('❌ FAIL: Expected details output missing.');
        process.exit(1);
    }

} catch (e) {
    console.error('❌ FAIL: Execution error:', e.message);
    process.exit(1);
}

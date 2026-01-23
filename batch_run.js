const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const { glob } = require('glob');

// Usage: node batch_run.js "<glob_pattern>" [arg1] [arg2] ...
// Example: node batch_run.js "./taxes/2025/*.xlsx" --pl --bs --save

const args = process.argv.slice(2);
if (args.length < 1) {
    console.error('Usage: node batch_run.js "<glob_pattern>" [report_args...]');
    process.exit(1);
}

const pattern = args[0];
const passThroughArgsList = args.slice(1).map(arg => {
    // Quote args that contain spaces to ensure they remain single arguments when executed
    return arg.includes(' ') ? `"${arg}"` : arg;
});
const passThroughArgs = passThroughArgsList.join(' ');

console.log(`[Batch Runner] Searching for files matching: "${pattern}"`);

(async () => {
    try {
        const files = await glob(pattern, { windowsPathsNoEscape: true });

        if (files.length === 0) {
            console.log('No files found matching pattern.');
            // Debug: show if pattern looks like a UNC path with issues
            if (pattern.startsWith('\\\\')) {
                console.log('Tip: Ensure UNC paths (\\\\server\\share) are correctly quoted in your shell.');
            }
            return;
        }

        console.log(`[Batch Runner] Found ${files.length} file(s):`);
        files.forEach(f => console.log(`  - ${f}`));
        console.log('\nStarting processing...\n');

        const runNext = (index) => {
            if (index >= files.length) {
                console.log('\n[Batch Runner] All files processed.');
                return;
            }

            const file = files[index];
            const absPath = path.resolve(file); // Ensure absolute path for robustness

            // User requested to ALWAYS ignore vendor.xlsx for batch runs BY DEFAULT,
            // but if they specify a custom --vendor-file, we should use it.
            const skipVendors = !passThroughArgs.includes('--vendor-file') && !passThroughArgs.includes('--ignore-vendors');
            const command = `node report.js "${absPath}" ${passThroughArgs}${skipVendors ? ' --ignore-vendors' : ''}`;

            console.log(`\n>>> [${index + 1}/${files.length}] Processing: ${file}`);
            console.log(`    Cmd: ${command}`);

            exec(command, { maxBuffer: 1024 * 1024 * 10 }, (error, stdout, stderr) => {
                // Print output first so we see what happened
                console.log(stdout);
                if (stderr) console.error(stderr);

                if (error) {
                    console.error(`\n[BATCH FATAL] Failed/Warning in ${file}. Exit code: ${error.code}`);
                    console.error(`Stopping batch run as requested.`);
                    process.exit(1);
                }

                runNext(index + 1);
            });
        };

        runNext(0);
    } catch (err) {
        console.error('Error finding files:', err);
    }
})();

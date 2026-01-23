const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const glob = require('glob');

// Usage: node batch_run.js "<glob_pattern>" [arg1] [arg2] ...
// Example: node batch_run.js "./taxes/2025/*.xlsx" --pl --bs --save

const args = process.argv.slice(2);
if (args.length < 1) {
    console.error('Usage: node batch_run.js "<glob_pattern>" [report_args...]');
    process.exit(1);
}

const pattern = args[0];
const passThroughArgs = args.slice(1).join(' ');

console.log(`[Batch Runner] Searching for files matching: "${pattern}"`);

glob(pattern, (err, files) => {
    if (err) {
        console.error('Error finding files:', err);
        return;
    }

    if (files.length === 0) {
        console.log('No files found matching pattern.');
        return;
    }

    console.log(`[Batch Runner] Found ${files.length} files. Starting processing...\n`);

    let completed = 0;

    const runNext = (index) => {
        if (index >= files.length) {
            console.log('\n[Batch Runner] All files processed.');
            return;
        }

        const file = files[index];
        const absPath = path.resolve(file); // Ensure absolute path for robustness
        const command = `node report.js "${absPath}" ${passThroughArgs}`;

        console.log(`\n>>> [${index + 1}/${files.length}] Processing: ${file}`);
        console.log(`    Cmd: ${command}`);

        const child = exec(command, { maxBuffer: 1024 * 1024 * 10 }, (error, stdout, stderr) => {
            if (error) {
                console.error(`    [ERROR] Failed to process ${file}. Exit code: ${error.code}`);
                // warning: we don't stop, we continue to next file
            }

            // Print a summary or the output? User asked to "run the report". 
            // Usually we want to see the output.
            console.log(stdout);
            if (stderr) console.error(stderr);

            runNext(index + 1);
        });
    };

    runNext(0);
});

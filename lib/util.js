// Resolves Windows .lnk and .url shortcuts to the file they point to.
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function resolveShortcut(filePath) {
    try {
        const ext = path.extname(filePath).toLowerCase();
        if (ext === '.lnk') {
            const escapedPath = filePath.replace(/'/g, "''");
            const command = `powershell -NoProfile -Command "(New-Object -ComObject WScript.Shell).CreateShortcut('${escapedPath}').TargetPath"`;
            const target = execSync(command).toString().trim();
            if (target) return target;
        } else if (ext === '.url') {
            const content = fs.readFileSync(filePath, 'utf8');
            const match = content.match(/^URL=(.*)$/m);
            if (match && match[1]) {
                let target = match[1].trim();
                if (target.startsWith('file:///')) target = target.replace('file:///', '');
                else if (target.startsWith('file://')) target = target.replace('file://', '');
                return decodeURIComponent(target);
            }
        }
    } catch (e) {
        console.error(`Warning: Failed to resolve shortcut '${filePath}': ${e.message}`);
    }
    return filePath;
}

module.exports = { resolveShortcut };

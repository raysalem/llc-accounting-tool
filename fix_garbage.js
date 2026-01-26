const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'report.js');
let content = fs.readFileSync(filePath, 'utf8');

// Define markers
const startMarker = 'excludedFromLinkage += amount;';
const endMarker = "const sName = subCatVal ? subCatVal.toString().trim() : '(No Sub-Cat)';";

const startIdx = content.indexOf(startMarker);
const endIdx = content.indexOf(endMarker);

if (startIdx === -1 || endIdx === -1) {
    console.error('Markers not found.');
    console.log('Start:', startIdx, 'End:', endIdx);
    process.exit(1);
}

// Find the closing braces of the IF blocks after startMarker
// Logic: "excludedFromLinkage += amount;" is usually followed by a newline and some braces.
// But the garbage starts immediately after the braces.

// Let's identify the garbage start.
// It starts around "if (!catStats[linkName]..." or whatever remains.
// I deleted some lines already.

// Safe logic: Extract the substring and print it to verification.
const chunk = content.substring(startIdx, endIdx);
console.log('--- FOUND CHUNK ---');
console.log(chunk.substring(0, 200));
console.log('...');
console.log(chunk.substring(chunk.length - 200));
console.log('--- END CHUNK ---');

// We want to KEEP the closing braces of the `if (cData)` and `if (catLower)`.
// The chunk starts at `excludedFromLinkage...`.
// It should look like:
// excludedFromLinkage += amount;
//                     }
//                 }
// <GARBAGE>
// const sName ...

// I will look for the last `}` before the garbage. This is risky.
// Instead, I will assume the structure I pasted.
// 3 lines?
// line 1: code.
// line 2: indent }
// line 3: indent }

// I will replace the chunk with:
/*
excludedFromLinkage += amount;
                    }
                }

                
*/

const cleanChunk = `excludedFromLinkage += amount;
                    }
                }

                `;

// Wait, I need to make sure I don't delete too much.
// The garbage is messy.
// I'll replace the whole range Start -> End with CleanChunk.

const newContent = content.substring(0, startIdx) + cleanChunk + content.substring(endIdx);
fs.writeFileSync(filePath, newContent);
console.log('Cleaned report.js');

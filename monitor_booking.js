const https = require('https');
const { exec } = require('child_process');

const CHECK_INTERVAL_MS = 5000;
const url = 'https://bookingsus.newbook.cloud/online/havasupai';

console.log(`Starting INTELLIGENT monitor for: ${url}`);
console.log(`Checking every ${CHECK_INTERVAL_MS / 1000} seconds...`);
console.log('Press Ctrl+C to stop.');

function beep(count) {
    if (count <= 0) return;
    process.stdout.write('\x07');
    setTimeout(() => beep(count - 1), 500);
}

function checkSite() {
    const req = https.get(url, (res) => {
        const timestamp = new Date().toLocaleTimeString();
        let data = '';

        // Collect the full response body
        res.on('data', (chunk) => {
            data += chunk;
        });

        res.on('end', () => {
            // Check status code first
            if (res.statusCode === 200) {
                // INTELLIGENT CHECK:
                // Sometimes error pages (like "Maintenance" or Cloudflare) return Status 200.
                // We want to verify this is the REAL booking page.

                const isMaintenance = data.toLowerCase().includes("maintenance") ||
                    data.toLowerCase().includes("under heavy load") ||
                    data.includes("502 Bad Gateway"); // Sometimes 502 text acts as 200 body

                // If it's a 200 but looks like a maintenance page, ignore it
                if (isMaintenance) {
                    console.log(`[${timestamp}] Status: 200 (False Positive - Maintenance Page Detected)`);
                }
                // If it looks like a Cloudflare/security challenge
                else if (data.includes("Just a moment") || data.includes("Challenge")) {
                    console.log(`[${timestamp}] Status: 200 (Blocked by Bot Protection/Cloudflare)`);
                }
                // REAL SUCCESS CASE
                else {
                    console.log(`\n[${timestamp}] !!! POTENTIAL SUCCESS !!! Status: ${res.statusCode}`);
                    console.log(`Page Title/Preview: ${data.substring(0, 100).replace(/\n/g, ' ')}...`);
                    console.log('GO GO GO! CHECK YOUR BROWSER NOW!');
                    beep(5);
                    exec(`start ${url}`);
                }
            } else {
                console.log(`[${timestamp}] Status: ${res.statusCode} - Still down`);
            }
        });
    });

    req.on('error', (e) => {
        console.log(`[${new Date().toLocaleTimeString()}] Connection Error: ${e.message}`);
    });

    req.end();
}

setInterval(checkSite, CHECK_INTERVAL_MS);
checkSite();

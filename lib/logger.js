// Captures console output so it can be written to the report's "Processing Log"
// sheet, and counts warnings/errors for the final status.
const util = require('util');

// The real console methods, kept so output can bypass the capture.
const originalConsole = { log: console.log, warn: console.warn, error: console.error };

global.globalWarningCount = 0;
const logBuffer = [];
const consoleLogger = {
    log: (...args) => {
        const msg = util.format(...args);
        originalConsole.log(msg); // Print to terminal using SAFE original console
        logBuffer.push(msg); // Store for Excel
    },
    error: (...args) => {
        const msg = util.format(...args);
        originalConsole.error(msg);
        logBuffer.push('[ERROR] ' + msg);
        global.globalWarningCount = (global.globalWarningCount || 0) + 1;
    },
    warn: (...args) => {
        const msg = util.format(...args);
        originalConsole.warn(msg);
        logBuffer.push('[WARN] ' + msg);
        global.globalWarningCount = (global.globalWarningCount || 0) + 1;
    }
};

// Routes console.log/warn/error through the capturing logger.
function captureConsole() {
    console.log = consoleLogger.log;
    console.warn = consoleLogger.warn;
    console.error = consoleLogger.error;
}

module.exports = { originalConsole, logBuffer, consoleLogger, captureConsole };

const fs = require('fs');

function normalizeAccountName(value) {
    return String(value || '').trim().toLowerCase()
        .replace(/[^a-z0-9]+/g, '_')
        .replace(/^_+|_+$/g, '');
}

function parseNumericLiteral(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (typeof value === 'string') {
        const s = value.trim();
        if (s === '') return null;
        const n = Number(s);
        return Number.isFinite(n) ? n : null;
    }
    return null;
}

function parseScalar(raw) {
    if (raw === undefined) return raw;
    const value = String(raw).trim();
    if (value === '') return '';
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        return value.slice(1, -1);
    }
    if (/^(true|false)$/i.test(value)) return value.toLowerCase() === 'true';
    const numeric = parseNumericLiteral(value);
    if (numeric !== null) return numeric;
    return value;
}

function parseToml(text) {
    const result = { accounts: {} };
    let currentSection = null;
    const lines = text.split(/\r?\n/);

    for (const line of lines) {
        const raw = line.trim();
        if (!raw || raw.startsWith('#')) continue;
        const noComment = raw.replace(/#.*$/, '').trim();
        if (!noComment) continue;

        if (/^\[.*\]$/.test(noComment)) {
            const sectionName = noComment.slice(1, -1).trim();
            if (!sectionName) {
                currentSection = null;
                continue;
            }
            if (/^(accounts|balances|opening_balances|opening-balances|opening|root)$/i.test(sectionName)) {
                currentSection = 'accounts';
            } else {
                currentSection = sectionName.toLowerCase();
            }
            continue;
        }

        const match = noComment.match(/^([A-Za-z0-9_.\-\s]+)\s*=\s*(.+)$/);
        if (!match) continue;

        const key = match[1].trim();
        const value = parseScalar(match[2]);
        const target = currentSection === 'accounts' || currentSection === null ? result : result[currentSection] || (result[currentSection] = {});
        if (currentSection === null || currentSection === 'accounts') {
            if (key.toLowerCase() === 'year') result.year = value;
            else if (['as_of_date', 'as-of-date', 'asofdate', 'date'].includes(key.toLowerCase())) result.asOfDate = value;
            else if (key.toLowerCase() === 'as_of') result.asOfDate = value;
            else result.accounts[key] = value;
        } else {
            target[key] = value;
        }
    }

    if (!result.accounts || Object.keys(result.accounts).length === 0) {
        const candidateKeys = Object.keys(result).filter(k => !['year', 'asOfDate', 'as_of_date'].includes(k));
        if (candidateKeys.length > 0) {
            result.accounts = { ...result };
            delete result.accounts.year;
            delete result.accounts.asOfDate;
            delete result.accounts.as_of_date;
        }
    }

    if (result.accounts && !result.year && result.year !== 0) {
        const maybeYear = result.accounts.year;
        if (maybeYear !== undefined) result.year = maybeYear;
    }

    return result;
}

function parseJson(text) {
    return JSON.parse(text);
}

function parseXml(text) {
    const cleaned = text.replace(/<\?xml[^>]*>\s*/gi, '');
    const accounts = {};
    const yearMatch = cleaned.match(/<year\b[^>]*>([\s\S]*?)<\/year>/i);
    if (yearMatch) {
        const yearValue = yearMatch[1].replace(/<[^>]+>/g, '').trim();
        if (yearValue) accounts.__year = yearValue;
    }

    const asOfDateMatch = cleaned.match(/<(?:as_of_date|as-of-date|asofdate|date)\b[^>]*>([\s\S]*?)<\/\1>/i);
    if (asOfDateMatch) {
        const value = asOfDateMatch[1].replace(/<[^>]+>/g, '').trim();
        if (value) accounts.__asOfDate = value;
    }

    const accountsMatch = cleaned.match(/<accounts\b[^>]*>([\s\S]*?)<\/accounts>/i) || cleaned.match(/<opening-balances\b[^>]*>([\s\S]*?)<\/opening-balances>/i) || cleaned;
    const content = accountsMatch && typeof accountsMatch[1] === 'string' ? accountsMatch[1] : accountsMatch;
    const tagRegex = /<([A-Za-z0-9_.\-]+)\b[^>]*>([\s\S]*?)<\/\1>/g;
    let match;
    while ((match = tagRegex.exec(content)) !== null) {
        const tagName = match[1];
        const value = match[2].replace(/<[^>]+>/g, '').trim();
        if (!tagName || !value) continue;
        if (['year', 'as_of_date', 'as-of-date', 'asofdate', 'date', 'accounts', 'opening-balances', 'opening_balances'].includes(tagName.toLowerCase())) {
            continue;
        }
        accounts[tagName] = value;
    }

    const attributeRegex = /<([A-Za-z0-9_.\-]+)\b[^>]*\b(?:name|key|account)=["']([^"']+)["'][^>]*>([\s\S]*?)<\/\1>/g;
    while ((match = attributeRegex.exec(content)) !== null) {
        const tagName = match[1];
        const key = match[2];
        const value = match[3].replace(/<[^>]+>/g, '').trim();
        if (!key || !value) continue;
        if (['accounts', 'opening-balances', 'opening_balances'].includes(tagName.toLowerCase())) continue;
        accounts[key] = value;
    }

    if (yearMatch) {
        const yearValue = yearMatch[1].replace(/<[^>]+>/g, '').trim();
        if (yearValue) accounts.year = yearValue;
    }
    if (asOfDateMatch) {
        const value = asOfDateMatch[1].replace(/<[^>]+>/g, '').trim();
        if (value) accounts.asOfDate = value;
    }
    if (accounts.__year !== undefined) {
        accounts.year = accounts.__year;
        delete accounts.__year;
    }
    if (accounts.__asOfDate !== undefined) {
        accounts.asOfDate = accounts.__asOfDate;
        delete accounts.__asOfDate;
    }

    return { accounts };
}

function parseOpeningBalancesText(text, filePath = 'opening-balances.toml') {
    const lower = String(filePath).toLowerCase();
    if (lower.endsWith('.json')) return parseJson(text);
    if (lower.endsWith('.xml')) return parseXml(text);
    if (lower.endsWith('.toml')) return parseToml(text);

    const json = (() => { try { return { parsed: parseJson(text), format: 'json' }; } catch (e) { return null; } })();
    if (json) return json.parsed;

    const toml = (() => { try { return { parsed: parseToml(text), format: 'toml' }; } catch (e) { return null; } })();
    if (toml) return toml.parsed;

    const xml = (() => { try { return { parsed: parseXml(text), format: 'xml' }; } catch (e) { return null; } })();
    if (xml) return xml.parsed;

    throw new Error(`Unsupported opening-balance file format for '${filePath}'. Use TOML, JSON, or XML.`);
}

function parseOpeningBalancesFile(filePath) {
    const text = fs.readFileSync(filePath, 'utf8');
    return parseOpeningBalancesText(text, filePath);
}

function extractAccountsObject(data) {
    if (!data || typeof data !== 'object') {
        throw new Error('Opening balances file is empty or not an object.');
    }

    if (data.accounts && typeof data.accounts === 'object') return data.accounts;
    if (data.balances && typeof data.balances === 'object') return data.balances;
    if (data.openingBalances && typeof data.openingBalances === 'object') return data.openingBalances;
    if (data.opening_balances && typeof data.opening_balances === 'object') return data.opening_balances;

    const reserved = new Set(['year', 'asOfDate', 'as_of_date', 'as_of', 'date', 'notes', 'source', 'meta']);
    const entries = {};
    for (const [key, value] of Object.entries(data)) {
        if (!reserved.has(key)) entries[key] = value;
    }
    if (Object.keys(entries).length > 0) return entries;
    throw new Error('Opening balances file does not contain any account entries.');
}

function validateOpeningBalances(data, validCategories = null) {
    if (!data || typeof data !== 'object') {
        throw new Error('Opening balances file is empty or invalid.');
    }

    const yearValue = data.year ?? data.Year ?? data.taxYear;
    const year = Number(yearValue);
    if (!Number.isInteger(year) || String(yearValue).length !== 4) {
        throw new Error(`Opening balances file must include a 4-digit 'year' value. Received: ${String(yearValue)}`);
    }

    const accounts = extractAccountsObject(data);
    const result = [];
    const seen = new Set();
    const errors = [];

    for (const [rawKey, rawValue] of Object.entries(accounts)) {
        const key = String(rawKey).trim();
        if (!key) {
            errors.push('Found an empty account name.');
            continue;
        }
        const normalized = normalizeAccountName(key);
        if (seen.has(normalized)) {
            errors.push(`Duplicate opening-balance account after normalization: "${key}".`);
            continue;
        }
        seen.add(normalized);

        const amount = parseNumericLiteral(rawValue);
        if (amount === null) {
            errors.push(`Account "${key}" has a non-numeric value: ${String(rawValue)}.`);
            continue;
        }

        const entry = {
            key,
            normalized,
            amount,
            source: rawValue,
        };

        if (validCategories && !(validCategories.has(normalized) || validCategories.has(key.toLowerCase()))) {
            const categoryNames = Array.from(validCategories).slice(0, 5).join(', ');
            errors.push(`Account "${key}" is not defined in the Setup category list. Known examples: ${categoryNames}`);
            continue;
        }

        result.push(entry);
    }

    if (errors.length > 0) {
        throw new Error(`Opening balance validation failed:\n- ${errors.join('\n- ')}`);
    }

    if (result.length === 0) {
        throw new Error('Opening balances file contains no valid accounts.');
    }

    return {
        year,
        asOfDate: data.asOfDate || data.as_of_date || data.as_of || data.date || null,
        accounts: result,
    };
}

module.exports = {
    normalizeAccountName,
    parseNumericLiteral,
    parseOpeningBalancesText,
    parseOpeningBalancesFile,
    validateOpeningBalances,
    extractAccountsObject,
};


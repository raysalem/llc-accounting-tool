// Closest-match suggestions for validation messages ("Did you mean ...?").
// A small Levenshtein distance; no dependency.

// A candidate is suggested when its distance is at most this share of the longer
// string's length (rounded up). 0.4 catches a typo or a transposed pair
// ("bnak" -> "bank") but not an unrelated word.
const SUGGEST_MAX_RATIO = 0.4;

// Number of single-character edits (insert, delete, replace) between a and b.
function levenshtein(a, b) {
    let previous = Array.from({ length: b.length + 1 }, (_, i) => i);
    for (let i = 1; i <= a.length; i++) {
        const current = [i];
        for (let j = 1; j <= b.length; j++) {
            const cost = a[i - 1] === b[j - 1] ? 0 : 1;
            current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + cost);
        }
        previous = current;
    }
    return previous[b.length];
}

// The candidate closest to word (case-insensitive), or null if none is close enough.
function closest(word, candidates) {
    const w = String(word).trim().toLowerCase();
    let best = null;
    let bestDistance = Infinity;
    for (const candidate of candidates) {
        const c = String(candidate).trim().toLowerCase();
        const distance = levenshtein(w, c);
        const limit = Math.ceil(Math.max(w.length, c.length) * SUGGEST_MAX_RATIO);
        if (distance <= limit && distance < bestDistance) {
            best = candidate;
            bestDistance = distance;
        }
    }
    return best;
}

// ' Did you mean "x"?' (with a leading space), or '' when nothing is close.
function didYouMean(word, candidates) {
    const hit = closest(word, candidates);
    return hit === null ? '' : ` Did you mean "${hit}"?`;
}

module.exports = { levenshtein, closest, didYouMean, SUGGEST_MAX_RATIO };

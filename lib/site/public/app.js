// LLC Books page. Routes: #/ (all LLCs), #/<id>?l=pl&ld=brief&r=bs&rd=brief&cmp=1, and #faq (docs/faq.md).
// All text from the books is set with textContent (never as HTML).
'use strict';

const VIEWS = { pl: 'Profit & Loss', bs: 'Balance Sheet', vendors: 'Vendors', customers: 'Customers', t1099: '1099 Review' };
const ONE_LEVEL = new Set(['t1099']);
const EPS = 0.005;
let state = null;
let detail = null;

const el = (tag, attrs = {}, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
        if (k === 'class') e.className = v;
        else if (k === 'text') e.textContent = v;
        else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
        else e.setAttribute(k, v);
    }
    kids.flat(Infinity).forEach(k => { if (k !== null && k !== undefined && k !== false) e.append(k instanceof Node ? k : document.createTextNode(String(k))); });
    return e;
};
const money = n => (n === null || n === undefined || Number.isNaN(n)) ? '' : Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const timeOf = iso => iso ? new Date(iso).toLocaleTimeString() : '';

function route() {
    const [p, q] = location.hash.replace(/^#\/?/, '').split('?');
    const params = new URLSearchParams(q || '');
    return {
        id: decodeURIComponent(p || ''),
        l: VIEWS[params.get('l')] ? params.get('l') : 'pl',
        ld: params.get('ld') === 'detailed' ? 'detailed' : 'brief',
        r: VIEWS[params.get('r')] ? params.get('r') : (params.get('r') === 'none' ? 'none' : 'bs'),
        rd: params.get('rd') === 'detailed' ? 'detailed' : 'brief',
        cmp: params.get('cmp') === '1',
    };
}
function go(change) {
    const r = { ...route(), ...change };
    const q = new URLSearchParams({ l: r.l, ld: r.ld, r: r.r, rd: r.rd, cmp: r.cmp ? '1' : '0' });
    location.hash = `#/${encodeURIComponent(r.id)}?${q}`;
}

async function getJson(url) {
    const res = await fetch(url, { cache: 'no-store' });
    return res.json();
}

const isFaq = () => location.hash === '#faq';
let faqText = null;

async function refresh() {
    state = await getJson('/api/state');
    if (isFaq()) {
        const res = await fetch('/api/faq', { cache: 'no-store' });
        // An older site.js (started before the FAQ existed) answers 404: say so instead of showing a blank page.
        faqText = res.ok ? await res.text()
            : '# FAQ not available\n\nThe running site.js is older than this page. Stop it (Ctrl+C) and start it again: `node site.js`.';
        detail = null;
        render();
        return;
    }
    const r = route();
    detail = r.id ? await getJson(`/api/book/${encodeURIComponent(r.id)}`) : null;
    render();
}

function statusBadge(b) {
    if (b.running) return el('span', { class: 'badge run', text: 'updating…' });
    if (b.error) return el('span', { class: 'badge bad', text: b.error.kind === 'config' ? 'book file error' : 'report error' });
    if (!b.hasResult) return el('span', { class: 'badge run', text: 'no result yet' });
    return b.ok ? el('span', { class: 'badge ok', text: 'all checks pass' })
        : el('span', { class: 'badge bad', text: `${b.failedChecks.length} check${b.failedChecks.length === 1 ? '' : 's'} failing` });
}

function renderTabs() {
    const r = route();
    const tabs = document.getElementById('tabs');
    tabs.replaceChildren(...state.books.map(b => {
        const color = b.running ? '#7aa2ff' : (b.ok ? '#3ddc84' : '#e5484d');
        return el('a', { href: `#/${encodeURIComponent(b.id)}`, class: r.id === b.id ? 'on' : '' },
            el('span', { class: 'dot', style: `background:${color}` }), b.name);
    }));
}

function renderHome(main) {
    main.append(el('h1', { text: 'All LLCs' }));
    if (state.siteError) main.append(el('div', { class: 'banner bad', text: `Site file error (showing the last good list):\n${state.siteError}` }));
    if (!state.books.length) main.append(el('p', { class: 'muted', text: `No LLCs in ${state.siteFile}. Add them under [books].` }));
    main.append(el('div', { class: 'cards' }, state.books.map(b => el('a', { class: 'card', href: `#/${encodeURIComponent(b.id)}` },
        el('h2', { text: b.name }),
        el('div', { class: 'muted', text: `${b.id} · tax year ${b.taxYear || '?'}` }),
        el('div', { class: 'big', text: b.netIncome === null ? '—' : `Net income ${money(b.netIncome)}` }),
        statusBadge(b),
        b.error ? el('div', { class: 'banner bad', style: 'margin-top:8px', text: b.error.message.split('\n')[0] }) : null,
        b.failedChecks.length ? el('ul', {}, b.failedChecks.map(c => el('li', { text: `${c.title} (${c.count})` }))) : null,
        el('div', { class: 'muted', style: 'margin-top:6px;font-size:12px', text: b.updated ? `Updated ${timeOf(b.updated)}` : '' }),
    ))));
}

// ---- tables -------------------------------------------------------------

// cols: [{ head, get(row) -> number|string, num }]; rows: [{ key, label, cls, values: row }]
// With compare: the first numeric column gets Before and Change columns from prevMap.
function table(cols, rows, { compare = false, prevMap = null } = {}) {
    const mainCol = cols.findIndex(c => c.num && c.compare !== false);
    const head = el('tr', {}, el('th', { text: '' }), cols.map((c, i) => [
        compare && i === mainCol ? el('th', { text: `Before` }) : null,
        el('th', { text: compare && i === mainCol ? `${c.head} (after)` : c.head }),
        compare && i === mainCol ? el('th', { text: 'Change' }) : null,
    ]));
    const body = rows.map(r => {
        const tr = el('tr', { class: r.cls || '' });
        tr.append(el('td', { text: r.label }));
        let changed = false;
        cols.forEach((c, i) => {
            const v = r.values ? c.get(r.values) : null;
            if (compare && i === mainCol) {
                const before = prevMap && r.key !== undefined ? prevMap.get(r.key) : undefined;
                const hasBefore = before !== undefined;
                tr.append(el('td', { class: 'before', text: hasBefore ? money(before) : (r.values && prevMap ? '—' : '') }));
                tr.append(el('td', { text: c.num ? money(v) : (v || '') }));
                const delta = (typeof v === 'number' ? v : 0) - (hasBefore ? before : 0);
                const show = r.values && prevMap && Math.abs(delta) > EPS;
                if (show) changed = true;
                tr.append(el('td', { class: show ? (delta > 0 ? 'up' : 'down') : '', text: show ? (delta > 0 ? '+' : '') + money(delta) : '' }));
                if (r.values && prevMap && r.key !== undefined && !hasBefore) tr.classList.add('added');
            } else {
                tr.append(el('td', { text: c.num ? money(v) : (v === null || v === undefined ? '' : v) }));
            }
        });
        if (changed) tr.classList.add('changed');
        return tr;
    });
    return el('table', {}, el('thead', {}, head), el('tbody', {}, body));
}

const sum = rows => rows.reduce((a, r) => a + (r.value || 0), 0);

// Rows of the previous result, keyed like the current ones, for before/after.
function prevMapFor(view, prev) {
    if (!prev) return null;
    const m = new Map();
    const addRows = (sec, rows) => rows.forEach(r => {
        m.set(`${sec}:${r.label}`, r.value);
        Object.entries(r.subCats || {}).forEach(([s, v]) => m.set(`${sec}:${r.label}/${s}`, v));
    });
    if (view === 'pl') { addRows('pl', prev.pl); m.set('pl:=total', prev.netIncome); }
    if (view === 'bs') {
        addRows('a', prev.assets); addRows('l', prev.liabilities); addRows('e', prev.equity);
        m.set('a:=total', sum(prev.assets)); m.set('l:=total', sum(prev.liabilities)); m.set('e:=total', sum(prev.equity));
        m.set('le:=total', sum(prev.liabilities) + sum(prev.equity));
    }
    if (view === 'vendors') { addRows('v', prev.vendors); m.set('v:=total', sum(prev.vendors)); }
    if (view === 'customers') { addRows('c', prev.customers); m.set('c:=total', sum(prev.customers)); }
    if (view === 't1099') prev.review1099.forEach(r => m.set(`t:${r.vendor}`, r.paid));
    return m;
}

function categoryRows(sec, rows, detailed, sheets, extra = () => ({})) {
    const out = [];
    rows.forEach(r => {
        out.push({ key: `${sec}:${r.label}`, label: r.label, values: { ...r, ...extra(r) } });
        if (detailed) {
            Object.entries(r.subCats || {}).filter(([s]) => s).sort().forEach(([s, v]) => {
                out.push({ key: `${sec}:${r.label}/${s}`, label: s, cls: 'sub', values: { value: v, sheets: {} } });
            });
        }
    });
    return out;
}

function sheetCols(d, detailed) {
    if (!detailed) return [];
    return d.sheets.map(s => ({ head: s.short, num: true, compare: false, get: r => (r.sheets && r.sheets[s.name] ? r.sheets[s.name].total : null) }));
}

function viewPL(d, detailed, cmp, prev) {
    const cols = [{ head: 'Amount', num: true, get: r => r.value }, ...sheetCols(d, detailed)];
    const rows = categoryRows('pl', d.pl, detailed);
    rows.push({ key: 'pl:=total', label: 'Net income', cls: 'total', values: { value: d.netIncome } });
    return table(cols, rows, { compare: cmp, prevMap: cmp ? prevMapFor('pl', prev) : null });
}

function viewBS(d, detailed, cmp, prev) {
    const cols = [
        { head: 'Opening', num: true, compare: false, get: r => r.opening },
        { head: 'End', num: true, get: r => r.value },
        ...sheetCols(d, detailed),
    ];
    // "End" is the compared column: put it first among numeric columns that compare.
    const rows = [];
    const section = (title, sec, list) => {
        rows.push({ label: title, cls: 'sec' });
        rows.push(...categoryRows(sec, list, detailed));
        rows.push({ key: `${sec}:=total`, label: `Total ${title.toLowerCase()}`, cls: 'total', values: { value: sum(list), opening: sum(list.map(r => ({ value: r.opening }))) } });
    };
    section('Assets', 'a', d.assets);
    section('Liabilities', 'l', d.liabilities);
    section('Equity', 'e', d.equity);
    rows.push({ key: 'le:=total', label: 'Total liabilities + equity', cls: 'total', values: { value: sum(d.liabilities) + sum(d.equity), opening: null } });
    return table(cols, rows, { compare: cmp, prevMap: cmp ? prevMapFor('bs', prev) : null });
}

function viewNames(sec, list, d, detailed, cmp, prev, withForm) {
    const forms = new Map(d.review1099.map(r => [r.vendor, r]));
    const cols = [{ head: 'Total', num: true, get: r => r.value }];
    if (withForm) {
        cols.push({ head: '1099 form', get: r => (forms.get(r.label) || {}).form || '—' });
        cols.push({ head: '1099 required', get: r => ((forms.get(r.label) || {}).reportable ? 'YES' : '') });
    }
    cols.push(...sheetCols(d, detailed));
    const rows = categoryRows(sec, list, detailed);
    rows.push({ key: `${sec}:=total`, label: 'Total', cls: 'total', values: { value: sum(list) } });
    return table(cols, rows, { compare: cmp, prevMap: cmp ? prevMapFor(sec === 'v' ? 'vendors' : 'customers', prev) : null });
}

function view1099(d, cmp, prev) {
    const rows = d.review1099
        .slice()
        .sort((a, b) => (b.reportable - a.reportable) || (b.paid - a.paid))
        .map(r => ({ key: `t:${r.vendor}`, label: r.vendor, cls: r.missing.length ? 'changed' : '', values: r }));
    const cols = [
        { head: 'Paid', num: true, get: r => r.paid },
        { head: 'Form', get: r => r.form || 'no decision' },
        { head: 'Threshold', num: true, compare: false, get: r => r.threshold },
        { head: '1099 required', get: r => (r.reportable ? 'YES' : '') },
        { head: 'Tax ID', get: r => (r.reportable || r.form ? (r.hasTaxId ? 'on file' : 'MISSING') : '') },
        { head: 'Missing', get: r => r.missing.join(', ') },
    ];
    const box = el('div');
    const due = d.review1099.filter(r => r.reportable).length;
    const gaps = d.review1099.filter(r => r.missing.length).length;
    box.append(el('p', { class: gaps ? 'no' : 'muted', text: `${due} vendor${due === 1 ? '' : 's'} need a 1099. ${gaps ? `${gaps} with missing details: fix before filing.` : 'All details complete.'} Tax IDs are never shown here, only whether one is on file.` }));
    box.append(table(cols, rows, { compare: cmp, prevMap: cmp ? prevMapFor('t1099', prev) : null }));
    return box;
}

function panel(view, level, d, cmp, prev) {
    const detailed = level === 'detailed' && !ONE_LEVEL.has(view);
    const title = `${VIEWS[view]}${ONE_LEVEL.has(view) ? '' : ` · ${detailed ? 'Detailed' : 'Brief'}`}`;
    let body;
    if (view === 'pl') body = viewPL(d, detailed, cmp, prev);
    else if (view === 'bs') body = viewBS(d, detailed, cmp, prev);
    else if (view === 'vendors') body = viewNames('v', d.vendors, d, detailed, cmp, prev, true);
    else if (view === 'customers') body = viewNames('c', d.customers, d, detailed, cmp, prev, false);
    else body = view1099(d, cmp, prev);
    return el('section', { class: 'panel' }, el('h3', { text: title }), body);
}

function seg(options, current, onPick) {
    return el('span', { class: 'seg' }, options.map(([v, label]) =>
        el('button', { class: v === current ? 'on' : '', onclick: () => onPick(v), text: label })));
}

function renderBook(main) {
    const r = route();
    const b = detail;
    if (!b || !b.id) {
        main.append(el('p', { text: `No LLC "${r.id}".` }), el('a', { href: '#/', text: 'All LLCs' }));
        return;
    }
    main.append(el('h1', {}, b.name, ' ', statusBadge(b)));
    if (b.error) {
        main.append(el('div', { class: 'banner bad', text: `${b.error.kind === 'config' ? 'Book file error' : 'The report stopped'}${b.current ? ' (showing the last good result)' : ''}:\n${b.error.message}` }));
    }
    const d = b.current;
    if (!d) { main.append(el('p', { class: 'muted', text: b.running ? 'Computing…' : 'No result yet.' })); return; }

    const viewOpts = Object.entries(VIEWS);
    const levelOpts = [['brief', 'Brief'], ['detailed', 'Detailed']];
    main.append(el('div', { class: 'toolbar' },
        el('span', {}, el('label', { text: 'Left' }), seg(viewOpts, r.l, v => go({ l: v })), ' ', seg(levelOpts, r.ld, v => go({ ld: v }))),
        el('span', {}, el('label', { text: 'Right' }), seg([...viewOpts, ['none', 'None']], r.r, v => go({ r: v })), ' ', r.r === 'none' ? null : seg(levelOpts, r.rd, v => go({ rd: v }))),
        el('span', {}, el('label', { text: 'Before / after' }), seg([['0', 'Off'], ['1', 'Last change']], r.cmp ? '1' : '0', v => go({ cmp: v === '1' }))),
        el('button', { class: 'act', onclick: () => fetch(`/api/book/${encodeURIComponent(b.id)}/refresh`, { method: 'POST' }), text: 'Recompute' }),
        el('span', { class: 'muted', text: `Updated ${timeOf(b.updated)} · tax year ${d.taxYear}` }),
    ));
    if (r.cmp) {
        main.append(el('div', { class: 'banner info', text: b.previous
            ? `Before = the result before the last change (${b.previous.stamp ? timeOf(b.previous.stamp.runTime) : ''}); after = now. Changed rows are highlighted.`
            : 'No earlier result yet in this session: save a change in Excel or in the book file to compare.' }));
    }

    // Page order (Ray, 2026-11-02): the chosen reports, then the checker, then the report log.
    // The checker always shows; exports are hard-gated on it (step 5).
    const failing = d.checks.filter(c => !c.ok);
    const checker = el('details', { class: 'checks', ...(failing.length ? { open: '' } : {}) },
        el('summary', { class: failing.length ? 'no' : 'yes', text: failing.length ? `Checker: ${failing.length} failing — reports and 1099s cannot be exported until these are fixed` : 'Checker: all checks pass' }),
        el('ul', {}, d.checks.map(c => el('li', { class: c.ok ? 'ok' : 'bad' }, `${c.ok ? '✓' : '✗'} ${c.title}`,
            c.ok ? null : el('ul', {}, c.items.slice(0, 50).map(i => el('li', { text: i })), c.items.length > 50 ? el('li', { text: `… and ${c.items.length - 50} more` }) : null)))),
    );

    const two = r.r !== 'none';
    main.append(el('div', { class: `panels${two ? ' two' : ''}` },
        panel(r.l, r.ld, d, r.cmp, b.previous),
        two ? panel(r.r, r.rd, d, r.cmp, b.previous) : null));
    main.append(checker);

    main.append(el('details', { class: 'log' }, el('summary', { text: 'Report log' }), el('pre', { text: d.log.join('\n') })));
}

// ---- FAQ ------------------------------------------------------------------
// A small Markdown reader for docs/faq.md: headings, paragraphs, lists, tables, code, **bold**, `code`, links.
// Builds DOM nodes (no HTML strings). Links to other repo docs show as text; web links open in a new tab.
function inline(text) {
    const out = [];
    const re = /\*\*(.+?)\*\*|`([^`]+)`|\[([^\]]+)\]\(([^)]+)\)/g;
    let last = 0;
    let m;
    while ((m = re.exec(text))) {
        if (m.index > last) out.push(text.slice(last, m.index));
        if (m[1] !== undefined) out.push(el('strong', {}, inline(m[1])));
        else if (m[2] !== undefined) out.push(el('code', { text: m[2] }));
        else if (/^https?:\/\//.test(m[4])) out.push(el('a', { href: m[4], target: '_blank', rel: 'noopener', text: m[3] }));
        else out.push(el('em', { text: m[3] }));
        last = re.lastIndex;
    }
    if (last < text.length) out.push(text.slice(last));
    return out;
}

const LIST = /^\s*(-|\d+\.)\s+/;
const TABLE = /^\|/;

function markdown(md) {
    const lines = md.replace(/\r\n/g, '\n').split('\n');
    const out = [];
    let i = 0;
    const para = [];
    const flush = () => { if (para.length) { out.push(el('p', {}, inline(para.join(' ')))); para.length = 0; } };
    while (i < lines.length) {
        const line = lines[i];
        const h = line.match(/^(#{1,4})\s+(.*)/);
        if (line.startsWith('```')) {
            flush();
            const code = [];
            for (i++; i < lines.length && !lines[i].startsWith('```'); i++) code.push(lines[i]);
            out.push(el('pre', {}, el('code', { text: code.join('\n') })));
            i++;
        } else if (h) {
            flush();
            out.push(el(`h${h[1].length}`, {}, inline(h[2])));
            i++;
        } else if (TABLE.test(line)) {
            flush();
            const rows = [];
            for (; i < lines.length && TABLE.test(lines[i]); i++) {
                if (/^\|[\s|:-]+\|$/.test(lines[i])) continue;
                rows.push(lines[i].replace(/^\||\|$/g, '').split('|').map(c => c.trim()));
            }
            out.push(el('table', {}, el('tr', {}, rows[0].map(c => el('th', {}, inline(c)))),
                rows.slice(1).map(r => el('tr', {}, r.map(c => el('td', {}, inline(c)))))));
        } else if (LIST.test(line)) {
            flush();
            const ordered = /^\s*\d+\./.test(line);
            const items = [];
            for (; i < lines.length && LIST.test(lines[i]); i++) items.push(lines[i].replace(LIST, ''));
            out.push(el(ordered ? 'ol' : 'ul', {}, items.map(t => el('li', {}, inline(t)))));
        } else if (!line.trim()) {
            flush();
            i++;
        } else {
            para.push(line.trim());
            i++;
        }
    }
    flush();
    return out;
}

function renderFaq(main) {
    main.append(el('div', { class: 'faq' }, faqText === null ? el('p', { class: 'muted', text: 'Loading…' }) : markdown(faqText)));
}

function render() {
    renderTabs();
    document.getElementById('faqlink').classList.toggle('on', isFaq());
    const main = document.getElementById('main');
    main.replaceChildren();
    if (state.codeChanged) {
        main.append(el('div', { class: 'banner bad', text: `The tool's code changed since site.js started (running ${state.versionId}, on disk ${state.codeChanged}). Stop site.js (Ctrl+C) and start it again.` }));
    }
    if (isFaq()) renderFaq(main);
    else if (route().id) renderBook(main); else renderHome(main);
    renderFooter();
}

// Version IDs: the site's own, and on an LLC page the one stamped on its report run.
function renderFooter() {
    // No versionId = a site.js started before Version IDs existed: say what to do, never "unknown".
    const parts = [state.versionId ? `Site Version ID ${state.versionId}` : 'Site Version ID missing: this site.js is older than the page. Stop it (Ctrl+C) and run node site.js again.'];
    const stamp = detail && detail.current && detail.current.stamp;
    if (stamp && stamp.versionId) parts.push(`report Version ID ${stamp.versionId} (run ${timeOf(stamp.runTime)})`);
    document.getElementById('foot').textContent = parts.join(' · ');
}

window.addEventListener('hashchange', refresh);
function listen() {
    const live = document.getElementById('live');
    const es = new EventSource('/api/events');
    es.onopen = () => live.classList.remove('off');
    es.onmessage = () => refresh();
    es.onerror = () => { live.classList.add('off'); live.title = 'Not connected: is site.js still running?'; };
}
// ?print: a still page (no live updates), for printing and screenshots.
const PRINT = new URLSearchParams(location.search).has('print');
refresh().then(() => { if (!PRINT) listen(); });

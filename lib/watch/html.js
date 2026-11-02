// HTML page for watch.js: the same before/after diff as the console, with
// Redline and Side-by-side views. A local file only; it reloads itself every
// few seconds and keeps the chosen view and scroll position.
const { sideBySide } = require('./diff');

const esc = s => String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function renderHtml({ title, ops, beforeLabel, afterLabel, status }) {
    const redline = ops.map(o => {
        if (o.op === 'same') return `<div class="l same">${esc(o.b) || '&nbsp;'}</div>`;
        if (o.op === 'del') return `<div class="l del">${esc(o.a) || '&nbsp;'}</div>`;
        return `<div class="l add">${esc(o.b) || '&nbsp;'}</div>`;
    }).join('\n');
    const split = sideBySide(ops).map(r => `<tr class="${r.kind}">`
        + `<td class="${r.left === null ? 'empty' : ''}">${esc(r.left) || '&nbsp;'}</td>`
        + `<td class="${r.right === null ? 'empty' : ''}">${esc(r.right) || '&nbsp;'}</td></tr>`).join('\n');
    return `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>${esc(title)}</title>
<style>
body { font: 13px/1.35 Consolas, Menlo, monospace; margin: 0; background: #fafafa; color: #222; }
header { position: sticky; top: 0; background: #fff; border-bottom: 1px solid #ccc; padding: 8px 12px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
header b { font-size: 14px; }
button { font: inherit; padding: 3px 10px; border: 1px solid #999; background: #eee; cursor: pointer; }
button.on { background: #333; color: #fff; }
.status { color: #555; }
.l { white-space: pre; padding: 0 8px; }
.del { background: #fde2e2; color: #a00; text-decoration: line-through; }
.add { background: #e0f5e0; color: #060; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
td { white-space: pre; overflow: hidden; text-overflow: ellipsis; padding: 0 8px; vertical-align: top; border-right: 1px solid #ddd; }
th { text-align: left; padding: 4px 8px; background: #eee; }
tr.change td:first-child, tr.del td:first-child { background: #fde2e2; color: #a00; }
tr.change td:last-child, tr.add td:last-child { background: #e0f5e0; color: #060; }
td.empty { background: #f0f0f0 !important; }
body.changes .same:not(.ctx) { display: none; }
#split, #redline { padding: 6px 0; }
</style></head>
<body>
<header>
  <b>${esc(title)}</b>
  <button id="bRed">Redline</button><button id="bSplit">Side by side</button>
  <button id="bChg">Changes only</button>
  <span class="status">${esc(status)}</span>
</header>
<div id="redline">${redline}</div>
<div id="split"><table><thead><tr><th>${esc(beforeLabel)}</th><th>${esc(afterLabel)}</th></tr></thead><tbody>
${split}
</tbody></table></div>
<script>
const st = JSON.parse(localStorage.getItem('llcWatch') || '{"view":"red","changes":false}');
function markContext() {
  for (const box of [document.querySelectorAll('#redline .l'), document.querySelectorAll('#split tbody tr')]) {
    const els = [...box];
    els.forEach((el, i) => {
      if (el.classList.contains('same')) return;
      for (let k = Math.max(0, i - 3); k <= Math.min(els.length - 1, i + 3); k++) els[k].classList.add('ctx');
    });
  }
}
function apply() {
  document.getElementById('redline').style.display = st.view === 'red' ? '' : 'none';
  document.getElementById('split').style.display = st.view === 'split' ? '' : 'none';
  document.body.classList.toggle('changes', st.changes);
  bRed.classList.toggle('on', st.view === 'red'); bSplit.classList.toggle('on', st.view === 'split');
  bChg.classList.toggle('on', st.changes);
  localStorage.setItem('llcWatch', JSON.stringify(st));
}
bRed.onclick = () => { st.view = 'red'; apply(); };
bSplit.onclick = () => { st.view = 'split'; apply(); };
bChg.onclick = () => { st.changes = !st.changes; apply(); };
markContext(); apply();
window.scrollTo(0, Number(localStorage.getItem('llcWatchY') || 0));
setTimeout(() => { localStorage.setItem('llcWatchY', String(window.scrollY)); location.reload(); }, 3000);
</script>
</body></html>
`;
}

module.exports = { renderHtml };

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const FREE_NAMES = 24;
const FREE_TPLS = 3;
const STORE_BILLING = 'https://store.microsoft.com/billing';
const PRO_SKU = 'namepress_pro';
const PRO_IDS = [PRO_SKU];
let STORE_URL = 'https://apps.microsoft.com/detail/9NRRLVDK5TH3';
let proAvailable = false; // Pro limits apply only when the Store actually sells the add-on
let proSku = PRO_SKU;
let isPro = false;
try { isPro = localStorage.getItem('namepress-pro') === '1'; } catch (_) {}
// ---------- Import a list from Excel, CSV or the clipboard (shared by the Press apps) ----------
// importList({ title, fields: [{ key, label, required, guess: /regex/ }], onDone(rows) })
// rows = [{ key: value, ... }] – only non-empty rows.
const DataIn = (() => {
  const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  let xlsxReady = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxReady) return xlsxReady;
    xlsxReady = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_URL; s.onload = () => res(window.XLSX); s.onerror = () => { xlsxReady = null; rej(new Error('Excel reader could not be loaded – connect to the internet once, or save the sheet as CSV.')); };
      document.head.appendChild(s);
    });
    return xlsxReady;
  }
  function parseDelimited(text) {
    text = String(text || '').replace(/^﻿/, '').replace(/\r\n?/g, '\n');
    const first = text.split('\n').slice(0, 5).join('\n');
    const count = (ch) => (first.match(new RegExp(ch === '\t' ? '\t' : '\\' + ch, 'g')) || []).length;
    const delim = ['\t', ';', ','].map((d) => [d, count(d)]).sort((a, b) => b[1] - a[1])[0];
    const d = delim[1] ? delim[0] : '\t';
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c;
      } else if (c === '"' && cell === '') q = true;
      else if (c === d) { row.push(cell); cell = ''; }
      else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += c;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.map((r) => r.map((x) => x.trim())).filter((r) => r.some((x) => x));
  }
  async function readFile(file) {
    if (/\.(xlsx|xlsm|xls|ods)$/i.test(file.name)) {
      const X = await loadXlsx();
      const wb = X.read(await file.arrayBuffer(), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return X.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }).map((r) => r.map((x) => String(x == null ? '' : x).trim())).filter((r) => r.some((x) => x));
    }
    const buf = await file.arrayBuffer();
    let text = new TextDecoder('utf-8').decode(buf);
    if (text.includes('�')) text = new TextDecoder('windows-1252').decode(buf); // old Excel CSV
    return parseDelimited(text);
  }
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function css() {
    if (document.getElementById('dataInCss')) return;
    const st = document.createElement('style'); st.id = 'dataInCss';
    st.textContent = `
.di-dlg{max-width:760px;width:calc(100vw - 32px)}
.di-dlg h2{margin:0 0 4px}
.di-src{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:10px 0}
.di-dlg textarea{width:100%;min-height:90px;font:13px/1.4 Consolas,monospace;padding:8px;border-radius:8px;border:1px solid #8886;background:transparent;color:inherit;box-sizing:border-box}
.di-map{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;margin:10px 0}
.di-map label{display:flex;flex-direction:column;gap:3px;font-size:13px;font-weight:600}
.di-map select{padding:6px;border-radius:8px;border:1px solid #8886;background:transparent;color:inherit}
.di-map select option{color:#111}
.di-prev{max-height:220px;overflow:auto;border:1px solid #8885;border-radius:8px}
.di-prev table{border-collapse:collapse;width:100%;font-size:13px}
.di-prev th,.di-prev td{border-bottom:1px solid #8884;padding:4px 8px;text-align:left;white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis}
.di-prev th{position:sticky;top:0;background:Canvas}
.di-msg{font-size:13px;opacity:.8;margin:6px 0}
.di-err{color:#dc2626}
.di-btns{display:flex;gap:8px;justify-content:flex-end;margin-top:12px}`;
    document.head.appendChild(st);
  }
  function importList(opts) {
    css();
    let dlg = document.getElementById('dataInDlg');
    if (dlg) dlg.remove();
    dlg = document.createElement('dialog'); dlg.id = 'dataInDlg'; dlg.className = 'dlg di-dlg';
    dlg.innerHTML = `<h2>${esc(opts.title || 'Import a list')}</h2>
      <p class="di-msg">Open an Excel or CSV file, or copy the cells in Excel / Google Sheets and paste them below.</p>
      <div class="di-src"><button type="button" class="primary" data-a="file">📂 Open Excel or CSV file…</button><span class="di-msg" data-r="fname"></span></div>
      <textarea data-r="paste" placeholder="…or paste here (Ctrl+V)"></textarea>
      <div data-r="mapBox" hidden>
        <label class="check"><input type="checkbox" data-r="head" checked> The first row has the column names</label>
        <div class="di-map" data-r="map"></div>
        <div class="di-prev" data-r="prev"></div>
        <p class="di-msg" data-r="count"></p>
      </div>
      <p class="di-msg di-err" data-r="err"></p>
      <div class="di-btns"><button type="button" class="ghost" data-a="cancel">Cancel</button><button type="button" class="primary" data-a="ok" disabled>Import</button></div>
      <input type="file" accept=".xlsx,.xls,.xlsm,.ods,.csv,.txt,.tsv" hidden data-r="file">`;
    document.body.appendChild(dlg);
    const R = (k) => dlg.querySelector(`[data-r="${k}"]`), A = (k) => dlg.querySelector(`[data-a="${k}"]`);
    let table = [];
    const cols = () => { const w = Math.max(0, ...table.map((r) => r.length)); return Array.from({ length: w }, (_, i) => (R('head').checked && table[0] && table[0][i]) || `Column ${String.fromCharCode(65 + (i % 26))}`); };
    function guessMap() {
      const names = cols(), used = new Set();
      R('map').innerHTML = opts.fields.map((f) => {
        let pick = -1;
        if (R('head').checked && f.guess) pick = names.findIndex((n, i) => !used.has(i) && f.guess.test(n));
        if (pick < 0 && f.required) pick = names.findIndex((_, i) => !used.has(i));
        if (pick >= 0) used.add(pick);
        return `<label>${esc(f.label)}${f.required ? ' *' : ''}<select data-k="${f.key}"><option value="-1">${f.required ? '– choose –' : '(not used)'}</option>${names.map((n, i) => `<option value="${i}"${i === pick ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select></label>`;
      }).join('');
      R('map').querySelectorAll('select').forEach((s) => (s.onchange = preview));
      preview();
    }
    function mapped() {
      const m = {}; R('map').querySelectorAll('select').forEach((s) => (m[s.dataset.k] = Number(s.value)));
      const body = R('head').checked ? table.slice(1) : table;
      return body.map((r) => { const o = {}; opts.fields.forEach((f) => { o[f.key] = m[f.key] >= 0 ? (r[m[f.key]] || '').trim() : ''; }); return o; })
        .filter((o) => opts.fields.some((f) => f.required && o[f.key]));
    }
    function preview() {
      const rows = mapped(), shown = opts.fields.filter((f) => Number(R('map').querySelector(`[data-k="${f.key}"]`).value) >= 0);
      R('prev').innerHTML = `<table><tr>${shown.map((f) => `<th>${esc(f.label)}</th>`).join('')}</tr>${rows.slice(0, 8).map((o) => `<tr>${shown.map((f) => `<td>${esc(o[f.key])}</td>`).join('')}</tr>`).join('')}</table>`;
      R('count').textContent = rows.length ? `${rows.length} row${rows.length > 1 ? 's' : ''} will be imported${rows.length > 8 ? ' (the first 8 are shown)' : ''}.` : 'No rows yet – choose the right columns.';
      A('ok').disabled = !rows.length;
    }
    function setTable(t, label) {
      table = t; R('err').textContent = '';
      if (!table.length) { R('err').textContent = 'No rows were found.'; R('mapBox').hidden = true; A('ok').disabled = true; return; }
      R('fname').textContent = label || '';
      // a first row that looks like headings
      const h = table[0].join(' ');
      R('head').checked = opts.fields.some((f) => f.guess && f.guess.test(h)) || /name|address|street|city|zip|email|title|date|company/i.test(h);
      R('mapBox').hidden = false; guessMap();
    }
    A('file').onclick = () => R('file').click();
    R('file').onchange = async () => {
      const f = R('file').files[0]; if (!f) return;
      R('fname').textContent = 'Reading…';
      try { setTable(await readFile(f), f.name); } catch (e) { R('fname').textContent = ''; R('err').textContent = e.message || 'That file could not be read.'; }
      R('file').value = '';
    };
    R('paste').oninput = () => { if (R('paste').value.trim()) setTable(parseDelimited(R('paste').value), ''); };
    R('head').onchange = guessMap;
    A('cancel').onclick = () => dlg.close();
    A('ok').onclick = () => { const rows = mapped(); dlg.close(); opts.onDone(rows); };
    dlg.addEventListener('close', () => setTimeout(() => dlg.remove(), 50));
    dlg.showModal();
  }
  return { importList, parseDelimited, readFile };
})();
const unlocked = () => isPro || !proAvailable;

async function billing() {
  if (!('getDigitalGoodsService' in window)) return null;
  try { return await window.getDigitalGoodsService(STORE_BILLING); } catch (_) { return null; }
}
function setPro(v) {
  isPro = v;
  try { localStorage.setItem('namepress-pro', v ? '1' : '0'); } catch (_) {}
  const b = $('proBtn');
  b.textContent = v ? '★ Pro' : '★ Get Pro';
  b.classList.toggle('is-pro', v);
  b.hidden = !proAvailable && !v;
  document.body.classList.toggle('unlocked', unlocked());
  if (typeof renderAll === 'function') renderAll();
}
async function checkPro() {
  const svc = await billing();
  if (!svc) return false;
  try {
    const details = await svc.getDetails(PRO_IDS).catch(() => []);
    proAvailable = details.length > 0;
    if (details[0]) { proSku = details[0].itemId; if (!PRO_IDS.includes(proSku)) PRO_IDS.push(proSku); }
    const list = await svc.listPurchases();
    const owned = list.some((p) => PRO_IDS.includes(p.itemId));
    setPro(owned);
    return owned;
  } catch (_) { return isPro; }
}
function proMessage(text) { const m = $('proMsg'); m.hidden = !text; m.textContent = text || ''; }
async function openPro(reason) {
  proMessage(reason || '');
  const svc = await billing();
  if (!svc) {
    $('buyBtn').textContent = 'Get NamePress on Microsoft Store';
    $('restoreBtn').hidden = true;
  } else {
    $('restoreBtn').hidden = false;
    try {
      const [d] = await svc.getDetails([proSku]);
      const price = d && d.price ? new Intl.NumberFormat(undefined, { style: 'currency', currency: d.price.currency }).format(Number(d.price.value)) : '';
      $('buyBtn').textContent = price ? `Unlock Pro – ${price}` : 'Unlock Pro';
    } catch (_) { $('buyBtn').textContent = 'Unlock Pro'; }
  }
  if (!$('proDialog').open) $('proDialog').showModal();
}
async function buyPro() {
  const svc = await billing();
  if (!svc) { window.open(STORE_URL, '_blank'); return; }
  const methods = [{ supportedMethods: STORE_BILLING, data: { sku: proSku } }];
  try {
    let req;
    try { req = new PaymentRequest(methods); }
    catch (_) { req = new PaymentRequest(methods, { total: { label: 'Total', amount: { currency: 'USD', value: '0' } } }); }
    const res = await req.show();
    await res.complete('success');
  } catch (e) {
    if (e && e.name === 'AbortError') return;
    proMessage('Purchase could not be completed: ' + (e.message || e));
    return;
  }
  if (await checkPro()) { proMessage('Thank you! Pro is unlocked. ★'); setTimeout(() => $('proDialog').close(), 1500); }
}
function needsPro(reason) {
  if (unlocked()) return false;
  openPro(reason);
  return true;
}


// ---------- options (millimetres) ----------
const SHEETS = [
  { id: 'us8', mode: 'tag', name: 'US Letter – 8 name badges, 2⅓ × 3⅜ in (Avery 5395 size)', pw: 215.9, ph: 279.4, cols: 2, rows: 4, w: 85.73, h: 59.27, left: 17.46, top: 15.08, px: 95.25, py: 63.5 },
  { id: 'us6', mode: 'tag', name: 'US Letter – 6 name badges, 3 × 4 in (Avery 5392 size)', pw: 215.9, ph: 279.4, cols: 2, rows: 3, w: 101.6, h: 76.2, left: 6.35, top: 25.4, px: 101.6, py: 76.2 },
  { id: 'a48', mode: 'tag', name: 'A4 – 8 name badges, 90 × 60 mm', pw: 210, ph: 297, cols: 2, rows: 4, w: 90, h: 60, left: 12.5, top: 21, px: 95, py: 65 },
  { id: 'usc4', mode: 'card', name: 'US Letter – 4 place cards, 3½ × 2 in folded', pw: 215.9, ph: 279.4, cols: 2, rows: 2, w: 88.9, h: 101.6, left: 19.05, top: 38.1, px: 88.9, py: 101.6 },
  { id: 'usc2', mode: 'card', name: 'US Letter – 2 large table cards, 7 × 4¼ in folded', pw: 215.9, ph: 279.4, cols: 1, rows: 2, w: 177.8, h: 127, left: 19.05, top: 12.7, px: 177.8, py: 127 },
  { id: 'a4c4', mode: 'card', name: 'A4 – 4 place cards, 90 × 50 mm folded', pw: 210, ph: 297, cols: 2, rows: 2, w: 90, h: 100, left: 15, top: 48.5, px: 90, py: 100 },
  { id: 'a4c2', mode: 'card', name: 'A4 – 2 large table cards, 180 × 65 mm folded', pw: 210, ph: 297, cols: 1, rows: 2, w: 180, h: 130, left: 15, top: 18.5, px: 180, py: 130 },
];
const TPLS = [{ id: 'hello', name: 'Hello' }, { id: 'classic', name: 'Classic' }, { id: 'simple', name: 'Simple' }, { id: 'elegant', name: 'Elegant' }, { id: 'modern', name: 'Modern' }, { id: 'stars', name: 'Stars' }];
const COLORS = ['#7c3aed', '#b91c1c', '#1e3a8a', '#14532d', '#9a6a00', '#0f766e', '#be185d', '#1f2937'];
const FONTS = { modern: '"Segoe UI", Arial, sans-serif', classic: 'Georgia, "Times New Roman", serif', script: '"Segoe Script", "Lucida Handwriting", cursive' };
const tplIndex = (id) => Math.max(0, TPLS.findIndex((t) => t.id === id));
const tplLocked = (id) => !unlocked() && tplIndex(id) >= FREE_TPLS;

// ---------- state ----------
let st = { mode: 'tag', sheet: 'us8', cardSheet: 'usc4', tpl: 'hello', color: COLORS[0], font: 'modern', top: 'HELLO my name is', bottom: '', names: '', start: 1, dx: 0, dy: 0, both: true, logo: '' };
try { Object.assign(st, JSON.parse(localStorage.getItem('namepress') || '{}')); } catch (_) {}
let saveTimer;
function save() { clearTimeout(saveTimer); saveTimer = setTimeout(() => { try { localStorage.setItem('namepress', JSON.stringify(st)); } catch (_) { toast('Could not save – the logo picture may be too large'); } }, 250); }
const num = (v, lo, hi, d) => { const n = parseFloat(String(v).replace(',', '.')); return isFinite(n) && n >= lo && n <= hi ? n : d; };
function people() { // "Name, detail" – the text after the first comma becomes the second line
  return String(st.names || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean).map((l) => { const i = l.indexOf(','); return i > 0 ? { name: l.slice(0, i).trim(), detail: l.slice(i + 1).trim() } : { name: l, detail: '' }; });
}
const sheet = () => SHEETS.find((s) => s.id === (st.mode === 'card' ? st.cardSheet : st.sheet) && s.mode === st.mode) || SHEETS.find((s) => s.mode === st.mode);

// ---------- text measuring ----------
let mctx = null;
function textW(s, fs, family, weight) {
  if (typeof document !== 'undefined') { try { if (!mctx) mctx = document.createElement('canvas').getContext('2d'); mctx.font = `${weight || 400} 100px ${family}`; return mctx.measureText(s).width * fs / 100; } catch (_) {} }
  return String(s).length * fs * (weight >= 700 ? 0.6 : 0.55);
}
const fit = (s, fs, maxW, fam, wt) => { const w = textW(s, fs, fam, wt); return w > maxW ? fs * maxW / w : fs; };
const f2 = (n) => n.toFixed(2);

// ---------- one face (name tag or one side of a place card); x,y top-left ----------
function face(x, y, w, h, p, screen) {
  const c = st.color || COLORS[0], k = Math.min(w / 86, h / 55), F = FONTS[st.font] || FONTS.modern, gold = '#c9992e';
  const T = (t, tx, ty, fs, extra, fam) => t ? `<text x="${f2(tx)}" y="${f2(ty)}" font-size="${f2(fs)}" text-anchor="middle" font-family='${fam || FONTS.modern}'${extra || ''}>${esc(t)}</text>` : '';
  let s = '', bandH = 0, topInk = c, nameTop = y;
  const cx = x + w / 2, pad = 4 * k;
  if (st.tpl === 'hello') { bandH = h * 0.32; s += `<rect x="${f2(x)}" y="${f2(y)}" width="${f2(w)}" height="${f2(bandH)}" fill="${c}"/><rect x="${f2(x)}" y="${f2(y + h - h * 0.08)}" width="${f2(w)}" height="${f2(h * 0.08)}" fill="${c}"/>`; topInk = '#fff'; }
  else if (st.tpl === 'classic') s += `<rect x="${f2(x + 2.5 * k)}" y="${f2(y + 2.5 * k)}" width="${f2(w - 5 * k)}" height="${f2(h - 5 * k)}" fill="none" stroke="${c}" stroke-width="${f2(0.9 * k)}"/><rect x="${f2(x + 4 * k)}" y="${f2(y + 4 * k)}" width="${f2(w - 8 * k)}" height="${f2(h - 8 * k)}" fill="none" stroke="${c}" stroke-width="${f2(0.25 * k)}"/>`;
  else if (st.tpl === 'elegant') { s += `<rect x="${f2(x + 3 * k)}" y="${f2(y + 3 * k)}" width="${f2(w - 6 * k)}" height="${f2(h - 6 * k)}" fill="none" stroke="${gold}" stroke-width="${f2(0.5 * k)}"/>`; [[x + 3 * k, y + 3 * k], [x + w - 3 * k, y + 3 * k], [x + 3 * k, y + h - 3 * k], [x + w - 3 * k, y + h - 3 * k]].forEach(([a, b]) => { s += `<circle cx="${f2(a)}" cy="${f2(b)}" r="${f2(1.2 * k)}" fill="${c}"/>`; }); }
  else if (st.tpl === 'modern') { bandH = 0; s += `<rect x="${f2(x)}" y="${f2(y)}" width="${f2(w * 0.035)}" height="${f2(h)}" fill="${c}"/><rect x="${f2(x)}" y="${f2(y + h - 2 * k)}" width="${f2(w)}" height="${f2(2 * k)}" fill="${c}" opacity=".25"/>`; }
  else if (st.tpl === 'stars') { const star = (sx, sy, r, col) => { let d = ''; for (let i = 0; i < 10; i++) { const a = Math.PI * i / 5 - Math.PI / 2, rr = i % 2 ? r * 0.42 : r; d += `${i ? 'L' : 'M'}${f2(sx + rr * Math.cos(a))},${f2(sy + rr * Math.sin(a))}`; } return `<path d="${d}Z" fill="${col}"/>`; }; const cols = [c, '#f59e0b', '#10b981', '#3b82f6', '#ef4444']; for (let i = 0; i < 6; i++) { s += star(x + (i + 0.5) * w / 6, y + 4.5 * k, 2.2 * k, cols[i % 5]) + star(x + (i + 0.5) * w / 6, y + h - 4.5 * k, 2.2 * k, cols[(i + 2) % 5]); } }
  else s += `<rect x="${f2(x + w * 0.3)}" y="${f2(y + h * 0.83)}" width="${f2(w * 0.4)}" height="${f2(0.5 * k)}" fill="${c}"/>`; // simple
  // top line
  const top = String(st.top || '').trim();
  if (st.tpl === 'hello' && /^hello\b/i.test(top)) {
    const rest = top.replace(/^hello\s*/i, '');
    s += T('HELLO', cx, y + bandH * 0.55, fit('HELLO', bandH * 0.5, w - 2 * pad, FONTS.modern, 800), ` font-weight="800" fill="#fff" letter-spacing="${f2(0.6 * k)}"`);
    s += T(rest, cx, y + bandH * 0.88, fit(rest, bandH * 0.22, w - 2 * pad, FONTS.modern), ' fill="#fff"');
  } else if (top) {
    const ty = bandH ? y + bandH * 0.62 : y + (st.tpl === 'stars' ? 13 : 11) * k;
    s += T(top, cx, ty, fit(top, (bandH ? bandH * 0.32 : 4.4 * k), w - 2 * pad, FONTS.modern, 600), ` font-weight="600" fill="${topInk}"`);
  }
  if (st.logo) { const ls = Math.min(h * 0.2, 12 * k); s += `<image href="${st.logo}" x="${f2(x + pad * 0.8)}" y="${f2(y + (bandH ? bandH + 1.5 * k : pad))}" width="${f2(ls)}" height="${f2(ls)}" preserveAspectRatio="xMidYMid meet"/>`; }
  // name and detail
  nameTop = bandH ? y + bandH : y + (top ? 14 : 6) * k;
  const bottom = String(st.bottom || '').trim(), detail = p.detail, lower = y + h - (st.tpl === 'hello' ? h * 0.08 : 0) - (bottom ? 7 * k : 3 * k);
  const room = lower - nameTop - (detail ? 7 * k : 0), wt = st.font === 'script' ? 400 : 700;
  const name = p.name || '';
  if (name) { const fs = fit(name, Math.min(room * 0.62, 14 * k), w * 0.86 - (st.logo ? 12 * k : 0), F, wt); s += T(name, cx, nameTop + room * 0.5 + fs * 0.36, fs, ` font-weight="${wt}" fill="#111"`, F); }
  else if (screen) s += T('Name', cx, nameTop + room * 0.5 + 3 * k, 9 * k, ' fill="#c3c7cc"', F);
  if (detail) s += T(detail, cx, nameTop + room + 5 * k, fit(detail, 4.4 * k, w - 2 * pad, FONTS.modern), ' fill="#444"');
  if (bottom) s += T(bottom, cx, lower + 5 * k, fit(bottom, 3.4 * k, w - 2 * pad, FONTS.modern), ` fill="${st.tpl === 'hello' ? '#555' : c}"`);
  return s;
}

// ---------- pages ----------
function pages(screen, test) {
  const S = sheet(), per = S.cols * S.rows, list = people();
  const first = S.mode === 'tag' ? Math.max(0, Math.min(per - 1, Math.round(num(st.start, 1, per, 1)) - 1)) : 0;
  const slots = test ? [] : Array(first).fill(null).concat(list.length ? list : [{ name: '', detail: '' }]);
  const dx = num(st.dx, -20, 20, 0), dy = num(st.dy, -20, 20, 0), out = [];
  const nPages = test ? 1 : Math.max(1, Math.ceil(slots.length / per));
  for (let pg = 0; pg < nPages; pg++) {
    let s = '', g = '';
    for (let i = 0; i < per; i++) {
      const x = S.left + (i % S.cols) * S.px, y = S.top + Math.floor(i / S.cols) * S.py, p = slots[pg * per + i];
      if (S.mode === 'tag') {
        if (screen || test) g += `<rect x="${f2(x)}" y="${f2(y)}" width="${S.w}" height="${S.h}" rx="3" fill="none" stroke="${test ? '#111' : '#c7ccd1'}" stroke-width="${test ? 0.25 : 0.3}"/>`;
        if (screen && pg === 0 && i < first) g += `<text x="${f2(x + S.w / 2)}" y="${f2(y + S.h / 2)}" font-size="4" fill="#b0b5ba" text-anchor="middle" font-family="Segoe UI, Arial">used</text>`;
        if (p) s += `<g>${face(x, y, S.w, S.h, p, screen)}</g>`;
      } else {
        if (!p) continue;
        const hh = S.h / 2;
        g += `<rect x="${f2(x)}" y="${f2(y)}" width="${S.w}" height="${S.h}" fill="none" stroke="#9aa0a6" stroke-width=".2" stroke-dasharray="1.5 1.2"/><line x1="${f2(x)}" y1="${f2(y + hh)}" x2="${f2(x + S.w)}" y2="${f2(y + hh)}" stroke="#c7ccd1" stroke-width=".2" stroke-dasharray="3 2"/>`;
        s += `<g>${face(x, y + hh, S.w, hh, p, screen)}</g>`;
        if (st.both) s += `<g transform="rotate(180 ${f2(x + S.w / 2)} ${f2(y + hh / 2)})">${face(x, y, S.w, hh, p, screen)}</g>`;
      }
    }
    const shift = S.mode === 'tag' ? ` transform="translate(${dx} ${dy})"` : '';
    out.push({ w: S.pw, h: S.ph, svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${S.pw} ${S.ph}" width="${S.pw}mm" height="${S.ph}mm" fill="#111"><rect width="${S.pw}" height="${S.ph}" fill="#fff"/><g${shift}>${g}${s}</g></svg>` });
  }
  return out;
}
function limitReason() {
  if (unlocked()) return '';
  if (tplLocked(st.tpl)) return `The "${TPLS[tplIndex(st.tpl)].name}" design is part of Pro. The first ${FREE_TPLS} designs are free.`;
  if (st.logo) return 'Adding your own logo is part of Pro. Remove the logo to print with the free version.';
  if (people().length > FREE_NAMES) return `The free version prints up to ${FREE_NAMES} names at a time. You can print as many times as you like.`;
  return '';
}
if (typeof module !== 'undefined') module.exports = { pages, face, people, sheet, limitReason, SHEETS, TPLS, getSt: () => st, setSt: (o) => { Object.assign(st, o); }, setProState: (a, p) => { proAvailable = a; isPro = p; } };

if (typeof document !== 'undefined') (function () {
  let pgs = [], pageNo = 0, toastTimer;
  window.toast = function (msg) { const t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 6000); };
  function renderPage() {
    const p = pgs[pageNo];
    $('page').innerHTML = p.svg.replace(/ width="[\d.]+mm" height="[\d.]+mm"/, ` style="aspect-ratio:${p.w}/${p.h};width:min(100%, calc((100vh - 170px) * ${p.w / p.h}))"`);
    $('pager').textContent = pgs.length > 1 ? `Sheet ${pageNo + 1} of ${pgs.length}` : 'Preview';
    $('prev').disabled = pageNo === 0; $('next').disabled = pageNo >= pgs.length - 1;
  }
  function thumb(id) { const was = st.tpl; st.tpl = id; const s = face(0, 0, 86, 55, { name: 'Anna', detail: '' }, false); st.tpl = was; return s; }
  function renderTpls() {
    $('tpls').innerHTML = TPLS.map((t) => `<button data-t="${t.id}" class="${t.id === st.tpl ? 'on' : ''}" title="${t.name}" aria-label="${t.name} design"><svg viewBox="0 0 86 55"><rect width="86" height="55" fill="#fff"/>${thumb(t.id)}</svg>${tplLocked(t.id) ? '<span class="badge">PRO</span>' : ''}</button>`).join('');
    $('colors').innerHTML = COLORS.map((c) => `<button data-c="${c}" class="${c === st.color ? 'on' : ''}" style="background:${c}" aria-label="Colour"></button>`).join('');
  }
  window.renderAll = function () {
    document.querySelectorAll('#modes button').forEach((b) => b.classList.toggle('on', b.dataset.m === st.mode));
    const tag = st.mode === 'tag';
    $('sheet').innerHTML = ''; SHEETS.filter((s) => s.mode === st.mode).forEach((s) => $('sheet').add(new Option(s.name, s.id)));
    $('sheet').value = sheet().id;
    ['font', 'top', 'bottom', 'names', 'start', 'dx', 'dy'].forEach((k) => { if (document.activeElement !== $(k)) $(k).value = st[k]; });
    $('both').checked = !!st.both; $('bothF').hidden = tag; $('startF').hidden = !tag; $('offF').hidden = !tag; $('testBtn').hidden = !tag;
    $('logoDel').hidden = !st.logo; $('logoBtn').textContent = st.logo ? 'Change logo…' : 'Add a logo picture…';
    renderTpls();
    pgs = pages(true); pageNo = Math.min(pageNo, pgs.length - 1);
    const n = people().length, unit = tag ? 'name badge' : 'place card';
    $('hint').textContent = n ? `${n} ${unit}${n === 1 ? '' : 's'} on ${pgs.length} sheet${pgs.length === 1 ? '' : 's'}.` : `No names yet: one empty ${unit} is shown. Type one name on each line.`;
    renderPage();
  };
  const change = () => { save(); renderAll(); };
  function printPages(list) {
    const p = list[0];
    $('pageRule').textContent = `@page { size: ${p.w}mm ${p.h}mm; margin: 0; } @media print { #printArea svg { width: ${p.w}mm !important; height: ${p.h}mm !important; } }`;
    $('printArea').innerHTML = list.map((x) => x.svg).join('');
    window.print();
  }
  function printNow() { const r = limitReason(); if (r) { openPro(r); return; } printPages(pages(false)); }
  function testPage() { printPages(pages(false, true)); toast('Hold the test page behind a badge sheet against a window to check that the boxes match.'); }
  function pickLogo() {
    if (needsPro('Adding your own logo is part of Pro.')) return;
    const i = $('fileIn'); i.value = '';
    i.onchange = () => { const f = i.files[0]; if (!f) return; const img = new Image(), url = URL.createObjectURL(f);
      img.onload = () => { const m = 400 / Math.max(img.width, img.height, 400), cv = document.createElement('canvas'); cv.width = Math.round(img.width * m); cv.height = Math.round(img.height * m); cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); st.logo = cv.toDataURL('image/png'); URL.revokeObjectURL(url); change(); };
      img.onerror = () => { URL.revokeObjectURL(url); toast('This file is not a picture that can be used.'); };
      img.src = url; };
    i.click();
  }
  document.querySelectorAll('#modes button').forEach((b) => { b.onclick = () => { st.mode = b.dataset.m; pageNo = 0; change(); }; });
  $('sheet').onchange = () => { if (st.mode === 'card') st.cardSheet = $('sheet').value; else st.sheet = $('sheet').value; pageNo = 0; change(); };
  $('font').onchange = () => { st.font = $('font').value; change(); };
  ['top', 'bottom', 'names', 'start', 'dx', 'dy'].forEach((k) => { $(k).oninput = () => { st[k] = $(k).value; change(); }; });
  $('both').onchange = () => { st.both = $('both').checked; change(); };
  $('tpls').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) { st.tpl = b.dataset.t; change(); if (tplLocked(st.tpl)) toast('This design is part of Pro. You can preview it here.'); } });
  $('colors').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) { st.color = b.dataset.c; change(); } });
  $('logoBtn').onclick = pickLogo; $('logoDel').onclick = () => { st.logo = ''; change(); };
  $('prev').onclick = () => { if (pageNo > 0) { pageNo--; renderPage(); } };
  $('next').onclick = () => { if (pageNo < pgs.length - 1) { pageNo++; renderPage(); } };
  $('printBtn').onclick = printNow; $('testBtn').onclick = testPage;
  $('moreBtn').onclick = (e) => { e.stopPropagation(); $('menu').hidden = !$('menu').hidden; };
  document.addEventListener('click', () => { $('menu').hidden = true; });
  $('importBtn').onclick = () => DataIn.importList({
    title: 'Import names from Excel or CSV',
    fields: [
      { key: 'n', label: 'Name', required: true, guess: /^(full ?)?name|guest|attendee|tên|participant/i },
      { key: 'l', label: 'Second line (optional)', guess: /table|company|organi[sz]ation|title|role|job|team|bàn|công ty|chức/i },
      { key: 'l2', label: 'Add to the second line (optional)', guess: /department|city|country/i },
    ],
    onDone: (rows) => {
      st.names = rows.map((o) => { const two = [o.l, o.l2].filter(Boolean).join(' · '); return (o.n.replace(/,/g, ' ') + (two ? ', ' + two.replace(/,/g, ' ') : '')).trim(); }).join('\n');
      pageNo = 0; change(); toast(`${rows.length} name${rows.length === 1 ? '' : 's'} imported.`);
    },
  });
  $('ex1').onclick = () => { Object.assign(st, { mode: 'tag', tpl: 'hello', top: 'HELLO my name is', bottom: 'Riverside Book Club', names: 'Margaret Ellis\nDavid Brooks\nSophia Nguyen\nRobert Miller\nHelen Carter\nFrank Murphy' }); pageNo = 0; change(); };
  $('ex2').onclick = () => { Object.assign(st, { mode: 'card', tpl: 'elegant', font: 'script', top: '', bottom: '', both: true, names: 'Emily Johnson, Table 1\nMichael Carter, Table 1\nLaura Bennett, Table 2\nJames Wilson, Table 2' }); pageNo = 0; change(); };
  $('resetBtn').onclick = () => { if (confirm('Clear all names and texts and start again?')) { try { localStorage.removeItem('namepress'); } catch (_) {} location.reload(); } };
  $('proBtn').onclick = () => { if (!isPro) openPro(''); };
  $('buyBtn').onclick = buyPro;
  $('restoreBtn').onclick = async () => { proMessage((await checkPro()) ? 'Pro is unlocked. ★' : 'No purchase was found on this Microsoft account.'); };
  $('closePro').onclick = () => $('proDialog').close();
  window.addEventListener('resize', () => renderPage());
  setPro(isPro);
  checkPro();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();

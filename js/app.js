import { extractTables, mergeAcrossPages, pdfTextToItems } from './extract.js';
import { parseNumber } from './numbers.js';
import { readPdf, readImage, renderPdfPage, isPdf } from './reader.js';

const $ = (id) => document.getElementById(id);
const state = { tables: [], textLines: [], docs: new Map(), current: 0, nextId: 1 };

// ---------- Utilidades de interfaz ----------
function setStatus(text, fraction) {
  $('status').hidden = false;
  $('status-text').textContent = text;
  if (fraction != null) $('bar-fill').style.width = Math.round(fraction * 100) + '%';
}
const opt = () => ({
  ocr: $('opt-ocr').value,
  lang: $('opt-lang').value,
  decimal: $('opt-decimal').value,
  merge: $('opt-merge').checked,
  numbers: $('opt-numbers').checked,
  text: $('opt-text').checked,
});

// ---------- Lectura de archivos ----------
async function processPdf(file, o, progress) {
  const { pdf, pages } = await readPdf(file, o, progress);
  const fileId = state.nextId++;
  state.docs.set(fileId, { kind: 'pdf', pdf });
  const found = [];
  for (const { page, pageW, pageH, items, source } of pages) {
    const { tables, textLines } = extractTables(items);
    for (const t of tables) found.push({ ...t, page, pageW, pageH, source });
    state.textLines.push(...textLines.map((t) => [file.name, page, t]));
  }
  return found.map((t) => ({ ...t, fileId, fileName: file.name }));
}

async function processImage(file, o, progress) {
  const { bitmap, pages } = await readImage(file, o, progress);
  const fileId = state.nextId++;
  state.docs.set(fileId, { kind: 'image', bitmap });
  const { page, pageW, pageH, items } = pages[0];
  const { tables, textLines } = extractTables(items);
  state.textLines.push(...textLines.map((t) => [file.name, 1, t]));
  return tables.map((t) => ({ ...t, page: 1, pageW, pageH, source: 'ocr', fileId, fileName: file.name }));
}

async function handleFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  const o = opt();
  state.tables = []; state.textLines = []; state.docs.clear(); state.current = 0;
  $('results').hidden = true; $('empty').hidden = true;
  try {
    let all = [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      const prefix = files.length > 1 ? `(${i + 1}/${files.length}) ` : '';
      const progress = (txt, frac) => setStatus(prefix + txt, (i + frac) / files.length);
      let found = isPdf(f) ? await processPdf(f, o, progress) : await processImage(f, o, progress);
      if (o.merge) found = mergeAcrossPages(found);
      all = all.concat(found);
    }
    state.tables = all.map((t, i) => ({ ...t, id: i, include: true }));
    setStatus(`Listo: ${state.tables.length} tabla(s) encontrada(s).`, 1);
    if (state.tables.length) { renderAll(); $('results').hidden = false; } else { $('empty').hidden = false; }
  } catch (err) {
    console.error(err);
    setStatus('No pude leer el archivo: ' + (err.message || err) + '. Si es un PDF con contraseña, quítale la clave primero.', 0);
  }
}

// ---------- Editor ----------
let sel = { r: 0, c: 0 };

function renderAll() { renderTabs(); renderGrid(); renderPreview(); }

function renderTabs() {
  const nav = $('tabs');
  nav.replaceChildren();
  state.tables.forEach((t, i) => {
    const b = document.createElement('button');
    b.className = 'tab' + (i === state.current ? ' active' : '');
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = t.include; cb.title = 'Incluir en el Excel';
    cb.addEventListener('click', (e) => { e.stopPropagation(); t.include = cb.checked; });
    const span = document.createElement('span');
    const pages = t.lastPage && t.lastPage !== t.page ? `págs. ${t.page}–${t.lastPage}` : `pág. ${t.page}`;
    span.innerHTML = '<b></b><small></small>';
    span.querySelector('b').textContent = t.title ? `${i + 1}. ${t.title.slice(0, 28)}` : `Tabla ${i + 1}`;
    span.querySelector('small').textContent = `${t.fileName} · ${pages}${t.source === 'ocr' ? ' · OCR' : ''}`;
    b.append(cb, span);
    b.addEventListener('click', () => { state.current = i; sel = { r: 0, c: 0 }; renderAll(); });
    nav.append(b);
  });
}

function renderGrid() {
  const t = state.tables[state.current];
  const table = $('grid');
  table.replaceChildren();
  if (!t) return;
  t.rows.forEach((row, r) => {
    const tr = document.createElement('tr');
    row.forEach((val, c) => {
      const td = document.createElement('td');
      td.contentEditable = 'plaintext-only';
      td.textContent = val;
      td.addEventListener('input', () => { t.rows[r][c] = td.textContent; });
      td.addEventListener('focus', () => { sel = { r, c }; });
      tr.append(td);
    });
    table.append(tr);
  });
}

const ops = {
  'add-row': (t) => t.rows.splice(sel.r + 1, 0, t.rows[0].map(() => '')),
  'add-col': (t) => t.rows.forEach((row) => row.splice(sel.c + 1, 0, '')),
  'del-row': (t) => { if (t.rows.length > 1) t.rows.splice(sel.r, 1); },
  'del-col': (t) => { if (t.rows[0].length > 1) t.rows.forEach((row) => row.splice(sel.c, 1)); },
  'merge-down': (t) => {
    if (sel.r >= t.rows.length - 1) return;
    t.rows[sel.r] = t.rows[sel.r].map((v, c) => [v, t.rows[sel.r + 1][c]].filter(Boolean).join(' '));
    t.rows.splice(sel.r + 1, 1);
  },
  'merge-right': (t) => {
    if (sel.c >= t.rows[0].length - 1) return;
    t.rows.forEach((row) => { row[sel.c] = [row[sel.c], row[sel.c + 1]].filter(Boolean).join(' '); row.splice(sel.c + 1, 1); });
  },
};

document.querySelector('.toolbar').addEventListener('click', (e) => {
  const b = e.target.closest('button[data-op]');
  const t = state.tables[state.current];
  if (!b || !t) return;
  ops[b.dataset.op](t);
  sel.r = Math.min(sel.r, t.rows.length - 1);
  sel.c = Math.min(sel.c, t.rows[0].length - 1);
  renderGrid();
});

async function renderPreview() {
  const t = state.tables[state.current];
  const canvas = $('preview-canvas');
  const hl = $('preview-hl');
  hl.style.display = 'none';
  if (!t) return;
  const doc = state.docs.get(t.fileId);
  let src;
  if (doc.kind === 'pdf') src = await renderPdfPage(doc.pdf, t.page, 1.5);
  else { src = document.createElement('canvas'); src.width = doc.bitmap.width; src.height = doc.bitmap.height; src.getContext('2d').drawImage(doc.bitmap, 0, 0); }
  if (state.tables[state.current] !== t) return; // el usuario cambió de tabla mientras se dibujaba
  canvas.width = src.width; canvas.height = src.height;
  canvas.getContext('2d').drawImage(src, 0, 0);
  const b = t.bbox;
  const pad = 4;
  hl.style.cssText = `display:block;left:${((b.x0 - pad) / t.pageW) * 100}%;top:${((b.y0 - pad) / t.pageH) * 100}%;` +
    `width:${((b.x1 - b.x0 + 2 * pad) / t.pageW) * 100}%;height:${((b.y1 - b.y0 + 2 * pad) / t.pageH) * 100}%`;
}

// ---------- Exportar ----------
function buildSheet(rows, o) {
  const aoa = rows.map((row) => row.map((val) => {
    const v = (val ?? '').trim();
    if (!o.numbers) return v;
    const n = parseNumber(v, o.decimal);
    if (!n) return v;
    return n.format ? { t: 'n', v: n.value, z: n.format } : { t: 'n', v: n.value };
  }));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const widths = [];
  rows.forEach((row) => row.forEach((v, c) => { widths[c] = Math.max(widths[c] || 8, Math.min(60, String(v).length + 2)); }));
  ws['!cols'] = widths.map((wch) => ({ wch }));
  return ws;
}

function sheetName(base, used) {
  let name = base.replace(/[\[\]:*?/\\]/g, ' ').slice(0, 31).trim() || 'Hoja';
  let n = 2;
  while (used.has(name.toLowerCase())) name = name.slice(0, 28) + ' ' + n++;
  used.add(name.toLowerCase());
  return name;
}

function exportXlsx() {
  const o = opt();
  const chosen = state.tables.map((t, i) => ({ t, i })).filter(({ t }) => t.include);
  if (!chosen.length && !(o.text && state.textLines.length)) { alert('Marca al menos una tabla para incluirla.'); return; }
  const wb = XLSX.utils.book_new();
  const used = new Set();
  for (const { t, i } of chosen) {
    const pages = t.lastPage && t.lastPage !== t.page ? `p${t.page}-${t.lastPage}` : `p${t.page}`;
    const label = t.title ? `${i + 1} ${t.title}` : `Tabla ${i + 1} (${pages})`;
    XLSX.utils.book_append_sheet(wb, buildSheet(t.rows, o), sheetName(label, used));
  }
  if (o.text && state.textLines.length) {
    XLSX.utils.book_append_sheet(wb, buildSheet([['Archivo', 'Página', 'Texto'], ...state.textLines.map(([f, p, x]) => [f, String(p), x])], { ...o, numbers: false }), sheetName('Texto', used));
  }
  const first = chosen[0]?.t.fileName || 'tablas.pdf';
  XLSX.writeFile(wb, first.replace(/\.[^.]+$/, '') + '.xlsx');
}

function exportCsv() {
  const t = state.tables[state.current];
  if (!t) return;
  const esc = (v) => /[",\n\r;]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  const csv = '﻿' + t.rows.map((r) => r.map((v) => esc((v ?? '').trim())).join(',')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = t.fileName.replace(/\.[^.]+$/, '') + `-tabla${state.current + 1}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---------- Eventos ----------
const drop = $('drop');
$('file').addEventListener('change', (e) => { handleFiles(e.target.files); e.target.value = ''; });
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => handleFiles(e.dataTransfer.files));
drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
$('btn-xlsx').addEventListener('click', exportXlsx);
$('btn-csv').addEventListener('click', exportCsv);

// Funciona sin internet después de la primera visita (útil en laptops con red restringida).
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

// Para pruebas automáticas
window.__pdfAExcel = { state, handleFiles };

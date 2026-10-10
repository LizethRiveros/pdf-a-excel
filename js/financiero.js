import { readPdf, readImage, isPdf } from './reader.js';
import { parseDocument } from './fin-parse.js';
import { buildModel } from './fin-cruce.js';
import { createKit } from './fin-excel.js';

const $ = (id) => document.getElementById(id);
const kit = createKit(XLSX);
const st = { docs: [], errores: [], empresa: {}, fecha: {}, confirmed: {}, model: null };
const fmt = (v) => (v == null ? '' : Number(v).toLocaleString('es-PE', { maximumFractionDigits: 2 }));
const SEV = { alta: 0, media: 1, info: 2 };

function setStatus(text, frac) {
  $('status').hidden = false;
  $('status-text').textContent = text;
  if (frac != null) $('bar-fill').style.width = Math.round(frac * 100) + '%';
}

async function handleFiles(list) {
  const files = [...list];
  if (!files.length) return;
  const ocr = $('opt-ocr').value;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    if (st.docs.some((d) => d.pdf === f.name)) { st.errores.push({ severidad: 'media', tipo: 'Duplicidad', pdf: f.name, pagina: '', empresa: '', detalle: 'Ya se había cargado un archivo con este nombre; se omitió el repetido.' }); continue; }
    const prefix = files.length > 1 ? `(${i + 1}/${files.length}) ` : '';
    const progress = (txt, frac) => setStatus(prefix + txt, (i + frac) / files.length);
    try {
      const read = isPdf(f) ? await readPdf(f, { ocr, lang: 'spa+eng' }, progress) : await readImage(f, { lang: 'spa+eng' }, progress);
      st.docs.push(parseDocument({ pdf: f.name, pages: read.pages }));
    } catch (err) {
      console.error(err);
      st.errores.push({ severidad: 'alta', tipo: 'Lectura', pdf: f.name, pagina: '', empresa: '', detalle: `No se pudo leer el archivo: ${err.message || err}. Si tiene contraseña, quítela primero.` });
    }
  }
  setStatus(`Listo: ${st.docs.length} PDF procesado(s).`, 1);
  rebuild();
}

function rebuild() {
  if (!st.docs.length) { $('out').hidden = !st.errores.length; renderErrores(); return; }
  st.model = buildModel(st.docs, { empresaOverrides: st.empresa, fechaOverrides: st.fecha, confirmed: st.confirmed, tol: Number($('opt-tol').value) || 0 });
  st.model.incidencias.push(...st.errores);
  render();
}

function renderErrores() {
  const t = $('t-inc');
  t.innerHTML = '<tr><th>Severidad</th><th>PDF</th><th>Detalle</th></tr>' + st.errores.map((e) => `<tr><td class="alta">ALTA</td><td>${esc(e.pdf)}</td><td>${esc(e.detalle)}</td></tr>`).join('');
}
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function render() {
  const m = st.model;
  $('out').hidden = false;
  const cnt = (s) => m.incidencias.filter((i) => i.severidad === s).length;
  const incorporados = m.docs.filter((d) => !d.duplicado && d.fechaCorte).length;
  const dif = m.docs.flatMap((d) => d.lineas).filter((l) => l.validacion === 'diferencia').length;
  const altas = m.incidencias.filter((i) => i.severidad === 'alta').slice(0, 8);
  $('informe').innerHTML = `
    <h2>Informe antes de entregar</h2>
    <div class="kpis">
      <div class="kpi"><b>${m.docs.length}</b>PDF procesados (${incorporados} incorporados)</div>
      <div class="kpi"><b>${m.empresas.length}</b>empresas identificadas</div>
      <div class="kpi"><b>${m.cortes.length}</b>fechas de corte (${esc(m.cortes.join(', ') || '—')})</div>
      <div class="kpi"><b>${m.pares.filter((p) => p.confirmada).length}</b>pares eliminados de ${m.pares.length} encontrados</div>
      <div class="kpi"><b>${m.noConciliados.length}</b>saldos relacionados sin conciliar</div>
      <div class="kpi"><b class="${cnt('alta') ? 'alta' : ''}">${cnt('alta')}</b>incidencias ALTAS (${cnt('media')} medias)</div>
    </div>
    <ul class="rep">
      <li>Subtotales de los PDF que no cuadran con sus componentes: <b>${dif}</b>.</li>
      <li>Empresas: ${m.empresas.map((e) => esc(e.nombre) + (e.ruc ? ` (RUC ${e.ruc})` : '')).join(' · ') || '—'}.</li>
      ${altas.map((i) => `<li class="alta">${esc(i.tipo)}${i.empresa ? ' · ' + esc(i.empresa) : ''}: ${esc(i.detalle)}</li>`).join('')}
    </ul>`;

  // descargas
  const box = $('downloads');
  box.replaceChildren();
  const mk = (label, fn, primary) => { const b = document.createElement('button'); b.textContent = label; if (primary) b.className = 'primary'; b.addEventListener('click', fn); box.append(b); };
  mk('Descargar todo (.zip)', downloadZip, true);
  mk('Solo el maestro consolidado', () => save(kit.buildMasterWorkbook(st.model), 'CONSOLIDADO MAESTRO.xlsx'));
  for (const e of m.empresas) mk(`Excel: ${e.nombre}`, () => save(kit.buildEmpresaWorkbook(st.model, e), `${kit.fileSafe(e.nombre)}.xlsx`));

  // PDF procesados
  const t = $('t-docs');
  t.innerHTML = '<tr><th>PDF</th><th>Tipo</th><th>Empresa (editable)</th><th>Corte (AAAA-MM-DD)</th><th>Páginas</th><th>Líneas</th><th>Cobertura</th><th>Resultado</th></tr>';
  for (const d of m.docs) {
    const pg = d.paginas.filter((p) => !p.informativa);
    const det = pg.reduce((s, p) => s + p.totalNums, 0), cap = pg.reduce((s, p) => s + p.leidos + (p.omitidos || 0), 0);
    const altasDoc = m.incidencias.filter((i) => i.pdf === d.pdf && i.severidad === 'alta').length;
    const res = d.duplicado ? 'Duplicado: excluido' : !d.fechaCorte ? 'Sin fecha: excluido' : altasDoc || cap < det ? 'Con observaciones' : 'Procesado íntegramente';
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${esc(d.pdf)}</td><td>${d.tipo === 'PDT710' ? 'PDT 710' : d.tipo === 'EEFF' ? 'EEFF' : 'No reconocido'}</td>
      <td><input class="cell" data-k="empresa" value="${esc(d.empresaNombre)}"></td><td><input class="cell" data-k="fecha" value="${esc(d.fechaCorte)}" placeholder="2026-07-31"></td>
      <td>${d.paginas.length}</td><td>${d.lineas.length}</td><td>${det ? Math.round((cap / det) * 100) : 100}%</td><td class="${res.startsWith('Procesado') ? '' : 'alta'}">${res}</td>`;
    tr.querySelectorAll('input').forEach((inp) => inp.addEventListener('change', () => {
      (inp.dataset.k === 'empresa' ? st.empresa : st.fecha)[d.idx] = inp.value.trim();
      rebuild();
    }));
    t.append(tr);
  }

  // pares
  const tp = $('t-pares');
  tp.innerHTML = '<tr><th>Eliminar</th><th>Corte</th><th>Por cobrar</th><th>Por pagar</th><th>Importes</th><th>Sustento</th></tr>';
  if (!m.pares.length) tp.insertAdjacentHTML('beforeend', '<tr><td colspan="6">No se encontró ningún par recíproco con el mismo corte, moneda e importe.</td></tr>');
  for (const p of m.pares) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><input type="checkbox" ${p.confirmada ? 'checked' : ''}></td><td>${esc(p.corte)}</td><td>${esc(p.cobrar.empresa)}<br><small>${esc(p.cobrar.cuenta)}</small></td><td>${esc(p.pagar.empresa)}<br><small>${esc(p.pagar.cuenta)}</small></td>
      <td>${fmt(p.cobrar.importe)} / ${fmt(p.pagar.importe)}<br><small>dif. ${fmt(p.diff)}</small></td><td><b>${esc(p.nivel)}</b>: ${esc(p.estado)}<br><small>${esc(p.sustento.join('; '))}</small></td>`;
    tr.querySelector('input').addEventListener('change', (e) => { st.confirmed[p.id] = e.target.checked; rebuild(); });
    tp.append(tr);
  }
  const tn = $('t-nc');
  tn.innerHTML = '<tr><th>Corte</th><th>Empresa</th><th>Lado</th><th>Cuenta</th><th>Importe</th><th>Motivo</th></tr>';
  if (!m.noConciliados.length) tn.insertAdjacentHTML('beforeend', '<tr><td colspan="6">Sin saldos no conciliados.</td></tr>');
  for (const r of m.noConciliados) tn.insertAdjacentHTML('beforeend', `<tr><td>${esc(r.fechaCorte)}</td><td>${esc(r.empresa)}</td><td>${r.lado === 'cobrar' ? 'Por cobrar' : r.lado === 'pagar' ? 'Por pagar' : '?'}</td><td>${esc(r.cuenta)}</td><td>${fmt(r.importe)}</td><td>${esc(r.motivo)}</td></tr>`);

  // incidencias
  const ti = $('t-inc');
  ti.innerHTML = '<tr><th>Severidad</th><th>Tipo</th><th>Empresa</th><th>PDF · pág.</th><th>Detalle</th></tr>';
  for (const i of [...m.incidencias].sort((a, b) => SEV[a.severidad] - SEV[b.severidad])) {
    ti.insertAdjacentHTML('beforeend', `<tr><td class="${i.severidad === 'alta' ? 'alta' : ''}">${i.severidad === 'info' ? 'Info' : i.severidad.toUpperCase()}</td><td>${esc(i.tipo)}</td><td>${esc(i.empresa)}</td><td>${esc(i.pdf)}${i.pagina ? ' · p.' + esc(i.pagina) : ''}</td><td>${esc(i.detalle)}</td></tr>`);
  }
}

// ---------- descargas ----------
function blobDownload(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function save(wb, name) {
  blobDownload(new Blob([kit.toBytes(wb)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
}
async function downloadZip() {
  const zip = new JSZip();
  for (const e of st.model.empresas) zip.file(`Por empresa/${kit.fileSafe(e.nombre)}.xlsx`, kit.toBytes(kit.buildEmpresaWorkbook(st.model, e)));
  zip.file('CONSOLIDADO MAESTRO.xlsx', kit.toBytes(kit.buildMasterWorkbook(st.model)));
  blobDownload(await zip.generateAsync({ type: 'blob' }), 'consolidado-financiero.zip');
}

// ---------- eventos ----------
const drop = $('drop');
$('file').addEventListener('change', (e) => { handleFiles(e.target.files); e.target.value = ''; });
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => handleFiles(e.dataTransfer.files));
drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
$('opt-tol').addEventListener('change', () => st.docs.length && rebuild());
$('btn-clear').addEventListener('click', () => { st.docs = []; st.errores = []; st.empresa = {}; st.fecha = {}; st.confirmed = {}; st.model = null; $('out').hidden = true; $('status').hidden = true; });

if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(() => {});
window.__fin = st;

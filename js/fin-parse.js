// Lector de documentos financieros: PDT 710 (SUNAT) y estados financieros (EEFF) en PDF.
// Entrada: páginas con fragmentos de texto (ver reader.js). Salida: líneas con cuenta, código, importe,
// página de origen, validación de subtotales e incidencias. No usa IA: reglas y posiciones.
import { analyzeLines } from './extract.js';
import { parseNumber } from './numbers.js';
import { findOverlaps } from './visibility.js';

export const norm = (s) =>
  String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase()
    .replace(/[^A-Z0-9%/.,\-() :]+/g, ' ').replace(/\s+/g, ' ').trim();

const MESES = { ENERO: 1, FEBRERO: 2, MARZO: 3, ABRIL: 4, MAYO: 5, JUNIO: 6, JULIO: 7, AGOSTO: 8, SETIEMBRE: 9, SEPTIEMBRE: 9,
  OCTUBRE: 10, NOVIEMBRE: 11, DICIEMBRE: 12, ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, SET: 9, SEP: 9, OCT: 10, NOV: 11, DIC: 12 };
const MES_ABR = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
const pad = (n) => String(n).padStart(2, '0');
const lastDay = (y, m) => new Date(y, m, 0).getDate();

// ---------- clasificación de celdas ----------
function kindOf(str) {
  const t = str.trim();
  const n = parseNumber(t, '.');
  if (n) return n.format ? { k: 'pct', v: n.value } : { k: 'num', v: n.value, int3: /^\d{3}$/.test(t) };
  if (/^N\.?\s?\d{1,3}$/i.test(t) || /^0\d$/.test(t)) return { k: 'code', v: t.replace(/\s/g, '') };
  return { k: 'text' };
}

const FURNITURE = [
  /^ESTADO DE /, /^\(?EXPRESADO EN/, /^(AL|DEL) \d/, /^[A-Z]{3}-\d{2}$/, /^PAGINA \d+/, /^REPORTE DEFINITIVO/,
  /^FORMULARIO 710/, /^TERCERA CATEGORIA/, /^ESTADOS FINANCIEROS$/, /RUC ?:|RAZON SOCIAL|PERIODO TRIBUTARIO|NUMERO DE ORDEN|NUMERO DE FORMULARIO|FECHA PRESENTACION|FECHA:/,
];
const isFurniture = (label) => FURNITURE.some((r) => r.test(norm(label)));

const SUB_RE = /^(TOTAL\b|SUB ?TOTAL|SALDOS? GENERALES?)/;
const RES_RE = /^(UTILIDAD|PERDIDA|RESULTADO|MARGEN|EBITDA)/;
function isSubtotal(label, clase) {
  const n = norm(label);
  if (n === '') return true;
  if (SUB_RE.test(n)) return true;
  if (clase === 'Resultados' && (RES_RE.test(n) || n === 'VENTAS NETAS')) return true;
  if (clase === 'Detalle resultados' && n === 'TOTAL') return true;
  return false;
}

// Une líneas "partidas": un monto y su rótulo que quedaron en líneas casi iguales pero desfasadas.
function mergeOffsetLines(lines, pitch) {
  const info = (l) => {
    const ks = l.cells.map((c) => kindOf(c.str).k);
    return { hasText: ks.includes('text'), hasNum: ks.includes('num') || ks.includes('code'), pctOnly: ks.every((k) => k === 'pct') };
  };
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const a = lines[i], b = lines[i + 1];
    if (b && b.yc - a.yc < pitch * 0.75) {
      const ia = info(a), ib = info(b);
      if ((ia.hasNum && !ia.hasText && ib.hasText && !ib.hasNum) || (ia.hasText && !ia.hasNum && ib.hasNum && !ib.hasText)) {
        out.push({ yc: (a.yc + b.yc) / 2, cells: [...a.cells, ...b.cells].sort((p, q) => p.x - q.x) });
        i++;
        continue;
      }
    }
    out.push(a);
  }
  return out;
}

const lineText = (l) => l.cells.map((c) => c.str).join(' ');

// ---------- datos del documento ----------
function parseFecha(text) {
  const t = norm(text);
  let m = t.match(/\bAL (\d{1,2}) DE ([A-Z]+)\.? (?:DE |DEL )?(\d{4})/);
  if (m && MESES[m[2]]) return { y: +m[3], m: MESES[m[2]], d: +m[1] };
  m = t.match(/\bAL (\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return { y: +m[3], m: +m[2], d: +m[1] };
  m = t.match(/\bAL (\d{1,2}) DE ([A-Z]+)\.? ?DEL? ?(\d{4})/);
  if (m && MESES[m[2]]) return { y: +m[3], m: MESES[m[2]], d: +m[1] };
  return null;
}

const normEmpresa = (s) => norm(s)
  .replace(/\bSOCIEDAD ANONIMA CERRADA\b/g, 'SAC').replace(/\bSOCIEDAD ANONIMA ABIERTA\b/g, 'SAA').replace(/\bSOCIEDAD ANONIMA\b/g, 'SA')
  .replace(/\bS\.?A\.?C\.?\b/g, 'SAC').replace(/\bS\.?R\.?L\.?\b/g, 'SRL').replace(/\bSOCIEDAD COMERCIAL DE RESPONSABILIDAD LIMITADA\b/g, 'SRL')
  .replace(/[.,]/g, '').replace(/\s+/g, ' ').trim();
export { normEmpresa };

function pageClass(tipo, text) {
  const t = norm(text);
  if (tipo === 'PDT710') {
    if (/ESTADO DE SITUACION FINANCIERA/.test(t)) return 'Balance';
    if (/ESTADO DE RESULTADOS/.test(t)) return 'Resultados';
    if (/DETERMINACION DE LA DEUDA|IMPUESTO A LA RENTA DE TERCERA/.test(t)) return 'Determinación';
    return 'Informativa';
  }
  if (/ESTADO DE SITUACION FINANCIERA/.test(t)) return 'Balance';
  if (/ESTADO DE RESULTADOS/.test(t)) return 'Resultados';
  if (/\bDEL \d{2}\/\d{2}\/\d{4} AL \d{2}\/\d{2}\/\d{4}/.test(t)) return 'Detalle resultados';
  return 'Detalle balance';
}

function detectTipo(pagesText) {
  const all = norm(pagesText.join('\n'));
  if (/FORMULARIO 0?710|NUMERO DE FORMULARIO 0?710/.test(all)) return 'PDT710';
  if (/ESTADO DE SITUACION FINANCIERA|ESTADO DE RESULTADOS|SALDOS GENERALES/.test(all)) return 'EEFF';
  return 'DESCONOCIDO';
}

// ---------- análisis de una página ----------
function segmentsOf(cells) {
  // Una "fila" = rótulo(s) seguido(s) de números. Devuelve segmentos {labelCells, nums}.
  const segs = [];
  let cur = null;
  for (const c of cells) {
    const k = kindOf(c.str);
    if (k.k === 'text') {
      if (!cur || cur.nums.length > 0 || cur.codes.length > 0) { cur = { labelCells: [], nums: [], codes: [], pct: null }; segs.push(cur); }
      cur.labelCells.push(c);
    } else {
      if (!cur) { cur = { labelCells: [], nums: [], codes: [], pct: null }; segs.push(cur); }
      if (k.k === 'num') cur.nums.push({ ...c, v: k.v, int3: k.int3 });
      else if (k.k === 'code') cur.codes.push(c.str.trim());
      else cur.pct = k.v;
    }
  }
  return segs;
}

function amountCells(page) {
  // cuenta de números de la página para el control de cobertura
  return page.lines.reduce((n, l) => n + l.cells.filter((c) => kindOf(c.str).k === 'num').length, 0);
}

// ---------- parser EEFF ----------
function parseEeffPage(pg, ctx) {
  const rows = [];
  let seccion = '';
  let prevHeading = false, lastHeadY = -1e9;
  let recentText = [];
  let colHeads = null;
  let omitted = 0;
  let cur = ctx.clase;
  let sawData = false;
  const titulosTardios = [];
  const empresaNorm = ctx.empresa ? normEmpresa(ctx.empresa) : '';
  for (const line of pg.lines) {
    const segs = segmentsOf(line.cells);
    const label = segs.flatMap((s) => s.labelCells).map((c) => c.str).join(' ').trim();
    const nums = segs.flatMap((s) => s.nums);
    // Una página puede cambiar de bloque a mitad (fin de las notas del balance + inicio del detalle de resultados).
    const nl = norm(label);
    // Un título de estado solo cambia el bloque al inicio de la página; uno que aparece después de cifras
    // es texto encimado de otro documento y no debe reclasificar lo que ya se leyó.
    if (/^ESTADO DE (SITUACION FINANCIERA|RESULTADOS)/.test(nl) && sawData) { titulosTardios.push(label); continue; }
    if (/^ESTADO DE SITUACION FINANCIERA/.test(nl)) cur = 'Balance';
    else if (/^ESTADO DE RESULTADOS/.test(nl)) cur = 'Resultados';
    else if (/^AL \d{1,2} DE /.test(nl) && cur !== 'Balance') cur = 'Detalle balance';
    else if (/^DEL \d{2}\/\d{2}\/\d{4} AL/.test(nl) && cur !== 'Resultados') cur = 'Detalle resultados';
    if (isFurniture(label) || (empresaNorm && normEmpresa(label) === empresaNorm)) { omitted += nums.length; continue; }
    const pct = segs.map((s) => s.pct).find((p) => p != null) ?? null;
    const codigo = segs.flatMap((s) => s.codes)[0] ?? '';
    if (nums.length === 0) {
      if (label) {
        seccion = prevHeading && line.yc - lastHeadY <= pg.pitch * 1.6 ? `${seccion} ${label}` : label;
        prevHeading = true; lastHeadY = line.yc;
        recentText.push(line);
      }
      continue; // línea de solo porcentaje u otros: no aporta importes
    }
    prevHeading = false;
    sawData = true;
    if (nums.length > 1 && !colHeads) {
      colHeads = recentText.filter((l) => line.yc - l.yc <= pg.pitch * 4.5).flatMap((l) => l.cells.map((c) => ({ ...c, y: l.yc })));
    }
    if (nums.length <= 1) colHeads = null;
    const centers = nums.map((n) => n.x + n.w / 2);
    const entries = nums.map((n, j) => {
      let columna = '';
      if (nums.length > 1 && colHeads) {
        columna = colHeads.filter((h) => {
          const hc = h.x + h.w / 2;
          const d = centers.map((c) => Math.abs(c - hc));
          const nearest = d.indexOf(Math.min(...d));
          return nearest === j && d[j] < 100;
        }).sort((x, y) => x.y - y.y || x.x - y.x).map((h) => h.str).join(' ');
      }
      return { importe: n.v, columna: columna || (nums.length > 1 ? `Col. ${j + 1}` : '') };
    });
    recentText = [];
    rows.push({ label, codigo, pct, entries, seccion, y: line.yc, nusados: nums.length, clase: cur });
  }
  return { rows, omitted, titulosTardios };
}

// ---------- parser PDT 710 ----------
function parsePdtPage(pg, ctx) {
  const twoSided = ctx.clase === 'Balance';
  const mid = pg.pageW / 2;
  const sides = twoSided ? [0, 1] : [0];
  const segsAll = [];
  const strays = [];
  let omitted = 0;
  for (const line of pg.lines) {
    const joined = lineText(line);
    if (isFurniture(joined) || isFurniture(line.cells[0]?.str || '')) { omitted += line.cells.filter((c) => kindOf(c.str).k === 'num').length; continue; }
    for (const side of sides) {
      const cells = line.cells.filter((c) => !twoSided || ((c.x + c.w / 2 < mid) === (side === 0)));
      if (!cells.length) continue;
      for (const s of segmentsOf(cells)) {
        const nums = s.nums;
        const label = s.labelCells.map((c) => c.str).join(' ').trim();
        if (!nums.length && !s.codes.length) { if (label) strays.push({ side, y: line.yc, label, x: s.labelCells[0].x }); continue; }
        segsAll.push({ side, y: line.yc, labelCells: s.labelCells, label, nums });
      }
    }
  }
  // Pares (casilla, importe) por segmento.
  const rows = [];
  for (const sg of segsAll) {
    const nums = sg.nums;
    let i = 0;
    const pairs = [];
    while (i < nums.length) {
      if (nums[i].int3) {
        const next = nums[i + 1];
        if (next) { pairs.push({ casilla: String(nums[i].v), amount: next.v, used: 2 }); i += 2; } else { pairs.push({ casilla: String(nums[i].v), amount: null, used: 1 }); i += 1; }
      } else { pairs.push({ casilla: '', amount: nums[i].v, used: 1, odd: true }); i += 1; }
    }
    pairs.forEach((p, j) => rows.push({ side: sg.side, y: sg.y, label: j === 0 ? sg.label : '', labelCells: j === 0 ? sg.labelCells : [], ...p }));
  }
  // Rótulos partidos: filas sin rótulo toman los textos sueltos cercanos del mismo lado;
  // los textos que empiezan en minúscula continúan el rótulo de la fila más cercana.
  const used = new Set();
  for (const r of rows.filter((x) => !x.label)) {
    const near = strays.filter((s) => s.side === r.side && Math.abs(s.y - r.y) <= pg.pitch * 1.3);
    near.forEach((s) => used.add(s));
    r.label = near.sort((a, b) => a.y - b.y).map((s) => s.label).join(' ');
  }
  for (const s of strays) {
    if (used.has(s) || !/^[a-záéíóúñ]/.test(s.label)) continue;
    const cand = rows.filter((r) => r.side === s.side && r.casilla && Math.abs(r.y - s.y) <= pg.pitch * 1.3 && r.y <= s.y);
    const best = cand.sort((a, b) => b.y - a.y)[0];
    if (best) { best.label = (best.label + ' ' + s.label).trim(); used.add(s); }
  }
  // Sección = último texto suelto no usado por encima de cada fila (mismo lado).
  const heads = strays.filter((s) => !used.has(s));
  for (const r of rows) {
    const above = heads.filter((h) => h.side === r.side && h.y < r.y).sort((a, b) => b.y - a.y)[0];
    r.seccion = above ? above.label : '';
  }
  rows.sort((a, b) => a.side - b.side || a.y - b.y);
  return { omitted, rows: rows.map((r) => ({
    label: r.label, codigo: r.casilla, pct: null, seccion: r.seccion, y: r.y, nusados: r.used + 0,
    entries: [{ importe: r.amount, columna: '' }], odd: r.odd, blank: r.amount == null,
  })) };
}

// ---------- validación de subtotales ----------
function validar(lineas) {
  const groups = new Map();
  for (const l of lineas) {
    const k = l.clase + '|' + l.columna;
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(l);
  }
  for (const [, g] of groups) {
    if (g[0].clase === 'Determinación') { g.forEach((l) => { l.subtotal = false; }); continue; }
    const tol = (v) => Math.max(1.5, Math.abs(v) * 0.00005);
    g.forEach((l, idx) => {
      if (!l.subtotal) return;
      const val = (x) => x.importe ?? 0;
      const prev = g.slice(0, idx);
      const prevSubs = prev.filter((x) => x.subtotal);
      const lastSubIdx = prev.map((x) => x.subtotal).lastIndexOf(true);
      const since = (i) => prev.slice(i + 1).filter((x) => !x.subtotal);
      const cands = [];
      const c1 = since(lastSubIdx);
      if (c1.length) cands.push({ refs: c1.map((x) => x.n), tipo: 'detalle' });
      const c3 = prev.filter((x) => !x.subtotal);
      if (c3.length) cands.push({ refs: c3.map((x) => x.n), tipo: 'detalle acumulado' });
      for (let k = 1; k <= Math.min(6, prevSubs.length); k++) {
        const s = prevSubs[prevSubs.length - k];
        const sIdx = prev.indexOf(s);
        cands.push({ refs: [s.n, ...since(sIdx).map((x) => x.n)], tipo: 'subtotal previo + detalle' });
      }
      for (const m of [2, 3]) if (prevSubs.length >= m) cands.push({ refs: prevSubs.slice(-m).map((x) => x.n), tipo: 'suma de subtotales' });
      const byN = new Map(lineas.map((x) => [x.n, x]));
      let best = null;
      for (const c of cands) {
        const sum = c.refs.reduce((s, n) => s + val(byN.get(n)), 0);
        if (Math.abs(sum - val(l)) <= tol(val(l))) { best = { ...c, sum }; break; }
      }
      if (best) { l.calc = { refs: best.refs, tipo: best.tipo }; l.validacion = 'ok'; l.diff = val(l) - best.sum; }
      else if (val(l) === 0) { l.validacion = 'ok'; l.diff = 0; l.calc = null; }
      else if (cands.length) {
        const c = cands[0];
        const sum = c.refs.reduce((s, n) => s + val(byN.get(n)), 0);
        l.validacion = 'diferencia'; l.diff = val(l) - sum; l.calc = { refs: c.refs, tipo: c.tipo, noCuadra: true };
      } else l.validacion = 'sin-validar';
    });
  }
}

// ---------- documento completo ----------
const analyzedItems = (a) => a.items || [];

export function parseDocument({ pdf, pages }) {
  const incidencias = [];
  const inc = (severidad, page, detalle, tipo = 'Extracción') => incidencias.push({ severidad, tipo, pdf, pagina: page, detalle });
  const analyzed = pages.map((p) => {
    const a = analyzeLines(p.items);
    return { items: p.items, page: p.page, pageW: p.pageW, pageH: p.pageH, source: p.source, ocultos: p.ocultos || [], dudosos: p.dudosos || [], sinVerificar: !!p.sinVerificar, ...a, lines: a.lines };
  });
  const texts0 = analyzed.map((a) => a.lines.map(lineText).join('\n'));
  const tipoPre = detectTipo(texts0);
  // En el PDT cada importe va anclado a su casilla, así que no se unen líneas desfasadas.
  if (tipoPre !== 'PDT710') for (const a of analyzed) a.lines = mergeOffsetLines(a.lines, a.pitch || 1);
  const texts = analyzed.map((a) => a.lines.map(lineText).join('\n'));
  const tipo = detectTipo(texts);
  if (tipo === 'DESCONOCIDO') inc('alta', '', 'Formato no reconocido (ni PDT 710 ni estados financieros). Se intentó leer como estados financieros; revisar.', 'Formato');
  const tipoUse = tipo === 'DESCONOCIDO' ? 'EEFF' : tipo;

  // Empresa, RUC, fecha de corte
  let empresa = '', ruc = '';
  let fecha = null;
  let moneda = '';
  if (tipoUse === 'PDT710') {
    for (const a of analyzed) for (const l of a.lines) {
      const t = lineText(l);
      if (!empresa) { const m = t.match(/Raz[oó]n Social\s*:?\s*([A-ZÁÉÍÓÚÑ0-9&.,\- ]{6,})$/i); if (m && !/^ANONIMA CERRADA$/i.test(m[1].trim())) empresa = m[1].trim(); }
      if (!ruc) { const m = t.match(/\bRUC\s*:?\s*(\d{11})\b/i); if (m) ruc = m[1]; }
    }
    const per = norm(texts.join('\n')).match(/PERIODO(?: TRIBUTARIO)? ?:? ?(\d{4})(\d{2})/);
    const y = per ? +per[1] : null;
    for (const t of texts) { const f = parseFecha(t); if (f) { fecha = f; break; } }
    if (!fecha && y) fecha = { y, m: 12, d: 31 };
    moneda = 'PEN';
    if (y && fecha && fecha.y !== y) inc('media', '', `La fecha de corte (${fecha.d}/${fecha.m}/${fecha.y}) no coincide con el periodo tributario ${y}.`, 'Periodo');
  } else {
    const firsts = analyzed.map((a) => a.lines[0] && lineText(a.lines[0])).filter((t) => t && !/\d/.test(t) && t.length < 70);
    const freq = new Map();
    firsts.forEach((t) => freq.set(t, (freq.get(t) || 0) + 1));
    empresa = [...freq.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
    const cnt = new Map();
    for (const t of texts) { const f = parseFecha(t); if (f) { const k = `${f.y}-${f.m}-${f.d}`; cnt.set(k, { f, n: (cnt.get(k)?.n || 0) + 1 }); } }
    fecha = [...cnt.values()].sort((a, b) => b.n - a.n)[0]?.f ?? null;
    if (!fecha) for (const t of texts) { const m = norm(t).match(/\bAL (\d{2})\/(\d{2})\/(\d{4})\b/); if (m) { fecha = { d: +m[1], m: +m[2], y: +m[3] }; break; } }
    const mon = norm(texts.join('\n')).match(/EXPRESADO EN (NUEVOS SOLES|SOLES|DOLARES|US\$|DOLARES AMERICANOS)/);
    moneda = !mon ? '' : /SOLES/.test(mon[1]) ? 'PEN' : 'USD';
    if (!mon) inc('media', '', 'No se indica la moneda ("Expresado en ..."). Se asume soles; confirmar.', 'Moneda');
    if (!ruc) inc('info', '', 'El documento no trae RUC; la empresa se identifica por su nombre.', 'Empresa');
  }
  if (!empresa) { inc('alta', '', 'No se pudo identificar la empresa.', 'Empresa'); empresa = `(Empresa sin identificar - ${pdf})`; }
  if (!fecha) { inc('alta', '', 'No se pudo identificar la fecha de corte del documento.', 'Periodo'); }
  const fechaCorte = fecha ? `${fecha.y}-${pad(fecha.m)}-${pad(fecha.d)}` : '';
  const periodoEtiqueta = tipoUse === 'PDT710' ? `PDT 710 ${fecha?.y ?? ''}` : fecha ? `EEFF ${MES_ABR[fecha.m]}-${fecha.y}` : 'EEFF (sin fecha)';
  if (fecha && tipoUse === 'EEFF' && fecha.d !== lastDay(fecha.y, fecha.m)) inc('info', '', `La fecha de corte (${fecha.d}/${fecha.m}/${fecha.y}) no es fin de mes.`, 'Periodo');

  const lineas = [];
  const paginas = [];
  let n = 0;
  for (const a of analyzed) {
    const clase = pageClass(tipoUse, texts[a.page - 1]);
    const base = { page: a.page, clase, source: a.source, totalNums: amountCells(a), leidos: 0, ocultos: a.ocultos.length };
    paginas.push(base);
    if (a.ocultos.length) {
      const nums = a.ocultos.map((h) => h.str.trim()).filter((t) => parseNumber(t, '.') || /^\(?-?[\d.,]+\)?%?$/.test(t));
      const muestra = [...new Set(nums)].slice(0, 6).join(', ');
      inc('alta', a.page, `Esta página contiene ${a.ocultos.length} texto(s) que NO se ven en el documento (texto oculto o encimado)${muestra ? `, entre ellos las cifras ${muestra}${nums.length > 6 ? '…' : ''}` : ''}. Se ignoraron para no mezclar importes; verifique que lo que se muestra coincide con lo que ve en el PDF.`, 'Texto oculto');
    }
    const ov = clase === 'Informativa' ? [] : findOverlaps(analyzedItems(a).filter((i) => i.str.trim().length >= 3));
    if (ov.length) inc('alta', a.page, `Texto superpuesto en la página (${ov.length} cruce(s), p. ej. "${ov[0][0]}" sobre "${ov[0][1]}"): las cifras de esta página pueden estar mezcladas con otro contenido. Verifíquelas contra el PDF.`, 'Texto superpuesto');
    const dudNum = a.dudosos.filter((h) => /\d/.test(h.str));
    if (dudNum.length) inc('media', a.page, `${dudNum.length} cifra(s) con tinta muy tenue en el documento (${[...new Set(dudNum.map((h) => h.str.trim()))].slice(0, 5).join(', ')}): confirme que se leen bien.`, 'Legibilidad');
    if (a.source === 'ocr') inc('media', a.page, 'Página leída con OCR (escaneo/imagen): verificar cada cifra contra el original.', 'OCR');
    if (clase === 'Informativa') { base.informativa = true; continue; }
    const parsed = tipoUse === 'PDT710' ? parsePdtPage(a, { clase, empresa }) : parseEeffPage(a, { empresa, clase });
    const rows = parsed.rows;
    base.omitidos = parsed.omitted;
    if (parsed.titulosTardios?.length) inc('media', a.page, `Aparecen títulos de estado en medio de la página (${[...new Set(parsed.titulosTardios)].join('; ')}): probable texto encimado; se ignoraron. Verifique la página.`, 'Texto superpuesto');
    for (const r of rows) {
      r.entries.forEach((e, j) => {
        const cuenta = r.label || (tipoUse === 'PDT710' ? (r.codigo ? `Casilla ${r.codigo}` : '(sin rótulo)') : '(subtotal sin rótulo)');
        const claseRow = r.clase || clase;
        const l = {
          n: n++, page: a.page, orden: lineas.length + 1, clase: claseRow, seccion: r.seccion || '', cuenta, codigo: r.codigo || '',
          columna: e.columna || '', importe: e.importe, pct: r.pct, subtotal: isSubtotal(r.label, claseRow), calc: null, validacion: '', diff: 0,
          enBlanco: e.importe == null, source: a.source, moneda,
        };
        lineas.push(l);
        if (!r.label && tipoUse === 'PDT710') inc('media', a.page, `Casilla ${r.codigo || '?'} sin rótulo detectado (importe ${e.importe ?? 'en blanco'}).`, 'Extracción');
        if (r.odd) inc('media', a.page, `Número suelto sin casilla: ${e.importe} (${r.label || 'sin rótulo'}). Revisar.`, 'Extracción');
        if (/(^|\s)-?\(?\d{1,3}(,\d{3})+(\.\d+)?\)?(\s|$)/.test(r.label)) inc('alta', a.page, `Posible cifra pegada dentro del rótulo "${r.label}": revisar que no falte un importe.`, 'Cobertura');
      });
      base.leidos += r.nusados;
    }
  }
  if (analyzed.some((a) => a.sinVerificar)) inc('media', '', 'No se pudo verificar si hay texto oculto en este PDF (no se renderizó la página).', 'Texto oculto');
  validar(lineas);
  for (const l of lineas) {
    if (l.validacion === 'diferencia') inc('alta', l.page, `El subtotal "${l.cuenta}"${l.columna ? ' [' + l.columna + ']' : ''} (${l.importe}) no cuadra con sus componentes (diferencia ${Math.round(l.diff * 100) / 100}).`, 'Validación');
  }
  for (const p of paginas) {
    if (p.informativa) continue;
    if (p.leidos + (p.omitidos || 0) !== p.totalNums) inc('media', p.page, `Se detectaron ${p.totalNums} cifras en la página: ${p.leidos} capturadas y ${p.omitidos || 0} en encabezados/pie. Faltan ${p.totalNums - p.leidos - (p.omitidos || 0)} por ubicar.`, 'Cobertura');
  }
  // Total de páginas informativas
  const info = paginas.filter((p) => p.informativa).map((p) => p.page);
  if (info.length) inc('info', info.join(', '), 'Páginas informativas (identificación, datos del declarante, constancia): no se extraen cifras; se usan solo para identificar empresa y periodo.', 'Cobertura');

  return { pdf, tipo, empresa: { nombre: empresa, ruc }, fechaCorte, periodoEtiqueta, moneda, paginas, lineas, incidencias };
}

// Generación de los Excel (uno por empresa + un maestro consolidado) a partir del modelo de fin-cruce.js.
// Recibe la librería XLSX (xlsx-js-style) para funcionar igual en el navegador y en Node.
import { norm } from './fin-parse.js';

const NUM = '#,##0.00;[Red]-#,##0.00';
const SEV = { alta: 0, media: 1, info: 2 };
const SEV_TXT = { alta: 'ALTA', media: 'MEDIA', info: 'Informativa' };

export function createKit(XLSX) {
  const colL = (i) => XLSX.utils.encode_col(i);
  const addr = (r, c) => XLSX.utils.encode_cell({ r, c });
  const S = {
    title: { font: { bold: true, sz: 14 } },
    sub: { font: { italic: true, color: { rgb: '555555' } } },
    head: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1D6F42' } }, alignment: { wrapText: true, vertical: 'center' } },
    bold: { font: { bold: true } },
    sec: { font: { bold: true, sz: 12 }, fill: { fgColor: { rgb: 'E2EFDA' } } },
    sub_row: { font: { bold: true }, fill: { fgColor: { rgb: 'F2F2F2' } } },
    bad: { font: { bold: true, color: { rgb: 'C00000' } }, fill: { fgColor: { rgb: 'FDE9E7' } } },
    ok: { font: { color: { rgb: '1D6F42' } } },
    warn: { fill: { fgColor: { rgb: 'FFF2CC' } } },
  };
  const t = (v, s) => ({ t: 's', v: v == null ? '' : String(v), ...(s ? { s } : {}) });
  const n = (v, s, f) => ({ t: 'n', v: v ?? 0, z: NUM, ...(f ? { f } : {}), ...(s ? { s } : {}) });
  const pctc = (v) => (v == null ? t('') : { t: 'n', v, z: '0.00%' });
  const head = (arr) => arr.map((x) => t(x, S.head));
  const q = (name) => `'${name.replace(/'/g, "''")}'`;

  function finish(aoa, widths, opts = {}) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = widths.map((w) => ({ wch: w }));
    if (opts.freezeRow) ws['!freeze'] = { xSplit: opts.freezeCol || 0, ySplit: opts.freezeRow };
    if (opts.freezeRow) ws['!views'] = [{ state: 'frozen', xSplit: opts.freezeCol || 0, ySplit: opts.freezeRow }];
    if (opts.merges) ws['!merges'] = opts.merges;
    return ws;
  }
  const safeName = (name, used) => {
    let s = name.replace(/[\[\]:*?/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31) || 'Hoja';
    let k = 2;
    while (used.has(s.toLowerCase())) s = s.slice(0, 28) + ' ' + k++;
    used.add(s.toLowerCase());
    return s;
  };
  const fileSafe = (s) => s.replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
  const short = (d) => (d.tipo === 'PDT710' ? (d.fechaCorte || '').slice(0, 4) : d.periodoEtiqueta.replace('EEFF ', ''));
  const claseSheet = (d, clase) => {
    const e = short(d);
    if (d.tipo === 'PDT710') return { Balance: `PDT Balance ${e}`, Resultados: `PDT Resultados ${e}`, 'Determinación': `PDT Renta ${e}` }[clase] || `PDT ${clase} ${e}`;
    return { Balance: `Balance ${e}`, Resultados: `Resultados ${e}`, 'Detalle balance': `Notas balance ${e}`, 'Detalle resultados': `Notas resultados ${e}` }[clase] || `${clase} ${e}`;
  };
  const rangeExpr = (rows, col) => {
    const s = [...new Set(rows)].sort((a, b) => a - b);
    const parts = [];
    for (let i = 0; i < s.length;) {
      let j = i;
      while (j + 1 < s.length && s[j + 1] === s[j] + 1) j++;
      parts.push(j > i ? `${col}${s[i]}:${col}${s[j]}` : `${col}${s[i]}`);
      i = j + 1;
    }
    return parts.join(',');
  };

  // ---------- hoja de un documento: se ve como el PDF y permite calcular ----------
  // Estilo tipo estado financiero: Arial 10, secciones en negrita, cuentas con sangría, totales con línea superior.
  const st = ({ b, i, sz = 10, color, fill, top, dbl, ind, h, wrap } = {}) => ({
    font: { name: 'Arial', sz, bold: !!b, italic: !!i, ...(color ? { color: { rgb: color } } : {}) },
    ...(fill ? { fill: { fgColor: { rgb: fill } } } : {}),
    ...(top || dbl ? { border: { ...(top ? { top: { style: 'thin', color: { rgb: '000000' } } } : {}), ...(dbl ? { bottom: { style: 'double', color: { rgb: '000000' } } } : {}) } } : {}),
    alignment: { vertical: 'top', wrapText: !!wrap, ...(h ? { horizontal: h } : {}), ...(ind ? { indent: ind } : {}) },
  });
  const GRAND = /^(TOTAL (ACTIVO|PASIVO Y PATRIMONIO|PATRIMONIO Y PASIVO)|UTILIDAD DEL EJERCICIO|PERDIDA DEL EJERCICIO|RESULTADO DE EJERCICIO)/;

  function presSheet(model, d, clase, entradas, lineMap) {
    const pdt = d.tipo === 'PDT710';
    const intFmt = pdt ? '#,##0;(#,##0)' : '#,##0;-#,##0';
    const maxCols = Math.min(5, Math.max(1, ...entradas.filter((e) => e.tipo === 'd').map((e) => e.lineIds.length)));
    const hasPct = !pdt && entradas.some((e) => e.tipo === 'd' && e.pct != null);
    const cAmt = 2, cPct = hasPct ? cAmt + maxCols : -1;
    const cSp = (hasPct ? cPct : cAmt + maxCols - 1) + 1, cPdf = cSp + 1, cDif = cSp + 2, cCtl = cSp + 3, cPag = cSp + 4;
    const aoa = [];
    const rowOf = new Map();
    const push = (cells) => { aoa.push(cells); return aoa.length; }; // devuelve la fila de Excel (1-based)
    const amountCell = (v, style) => (v == null ? t('', style) : { t: 'n', v, z: Number.isInteger(v) ? intFmt : '#,##0.00;-#,##0.00', s: style });

    // Título: igual al encabezado del PDF
    const pages = new Set(entradas.filter((e) => e.tipo !== 't').map((e) => e.page));
    const tit = d.estructura.find((e) => e.tipo === 't' && pages.has(e.page));
    let colTitulo = tit?.colTitulo || '';
    const lineasTit = [];
    for (const x of tit?.textos || []) { if (/^[A-Za-z]{3,}-\d{2,4}$/.test(x)) colTitulo = x; else if (!lineasTit.some((y) => norm(y) === norm(x))) lineasTit.push(x); }
    if (!lineasTit.length) lineasTit.push(d.empresa_.nombre);
    lineasTit.forEach((x, k) => push([t(x, st({ b: k < 2, sz: k === 0 ? 14 : k === 1 ? 12 : 10, i: k > 1 }))]));
    push([t(`${d.empresa_.nombre} · RUC ${d.empresa_.ruc || 'no consta'} · Corte ${d.fechaCorte} · ${d.moneda === 'USD' ? 'US$' : 'Soles (S/)'} · Fuente: ${d.pdf}`, st({ i: true, sz: 8, color: '666666' }))]);
    push([]);
    const headRow = [t('Cuenta', st({ b: true, fill: 'D9E2F3', top: true })), t(pdt ? 'Casilla' : 'Nota', st({ b: true, fill: 'D9E2F3', top: true, h: 'center' }))];
    for (let j = 0; j < maxCols; j++) headRow.push(t(maxCols === 1 ? (colTitulo || 'Importe') : `Importe ${j + 1}`, st({ b: true, fill: 'D9E2F3', top: true, h: 'right' })));
    if (hasPct) headRow.push(t('%', st({ b: true, fill: 'D9E2F3', top: true, h: 'right' })));
    headRow[cPdf] = t('En el PDF', st({ b: true, sz: 8, color: '666666', h: 'right' }));
    headRow[cDif] = t('Diferencia', st({ b: true, sz: 8, color: '666666', h: 'right' }));
    headRow[cCtl] = t('Control del subtotal', st({ b: true, sz: 8, color: '666666' }));
    headRow[cPag] = t('Pág.', st({ b: true, sz: 8, color: '666666', h: 'right' }));
    const headerRow = push(headRow);

    let prevPage = null;
    let lastMultiKey = '';
    for (const e of entradas) {
      if (e.tipo === 't' || e.consumido) continue;
      if (prevPage != null && e.page !== prevPage) push([t(`— pág. ${e.page} del PDF —`, st({ i: true, sz: 8, color: '999999' }))]);
      prevPage = e.page;
      for (let g = 0; g < (e.gap || 0); g++) push([]);
      if (e.tipo === 'h') { push([t(e.texto, st({ b: true })), ...[]]); lastMultiKey = ''; continue; }
      const ls = e.lineIds.map((n) => lineMap.get(n)).filter(Boolean);
      if (!ls.length) continue;
      const sub = ls.some((l) => l.subtotal);
      const grand = sub && GRAND.test(norm(e.texto || ''));
      const sBase = st({ b: sub, top: sub, dbl: grand });
      if (ls.length > 1) { // tabla de varias columnas: se repite su encabezado una vez
        const key = ls.map((l) => l.columna).join('|');
        if (key !== lastMultiKey) { const hr = []; ls.forEach((l, j) => { hr[cAmt + j] = t(l.columna, st({ b: true, sz: 8, h: 'right', wrap: true })); }); push(hr); lastMultiKey = key; }
      } else lastMultiKey = '';
      const r = aoa.length + 1;
      const row = [t(e.texto || '', st({ b: sub, top: sub, dbl: grand, ind: e.lvl || 0 })), t(e.codigo || '', st({ sz: 9, color: '444444', h: 'center', b: sub, top: sub, dbl: grand }))];
      let ctl = '', ctlStyle = null, pdfVal = null, difF = null;
      ls.forEach((l, j) => {
        rowOf.set(l.n, { r, c: cAmt + j });
        const col = colL(cAmt + j);
        if (l.subtotal && l.calc && !l.calc.noCuadra) {
          const refs = l.calc.refs.map((x) => rowOf.get(x)).filter(Boolean);
          if (refs.length === l.calc.refs.length) {
            const byCol = new Map();
            refs.forEach(({ r: rr, c }) => (byCol.get(c) ?? byCol.set(c, []).get(c)).push(rr));
            const parts = [...byCol.entries()].map(([c, rows]) => rangeExpr(rows, colL(c)));
            const sum = l.calc.refs.reduce((acc, x) => acc + (lineMap.get(x)?.importe ?? 0), 0);
            row[cAmt + j] = { t: 'n', v: sum, z: Number.isInteger(sum) ? intFmt : '#,##0.00;-#,##0.00', f: `SUM(${parts.join(',')})`, s: sBase };
            if (ls.length === 1) { pdfVal = l.importe; difF = `${col}${r}-${colL(cPdf)}${r}`; }
            const dif = Math.abs(l.diff) > 0.005;
            if (!ctl) { ctl = dif ? `Cuadra (redondeo ${Math.round(l.diff * 100) / 100})` : 'Cuadra'; ctlStyle = st({ sz: 8, color: '1D6F42' }); }
            return;
          }
        }
        row[cAmt + j] = amountCell(l.enBlanco ? null : l.importe, sBase);
        if (l.subtotal) {
          if (l.validacion === 'diferencia') { ctl = `DIFERENCIA ${Math.round(l.diff * 100) / 100} frente a sus componentes`; ctlStyle = st({ b: true, sz: 8, color: 'C00000' }); }
          else if (!ctl) { ctl = l.importe === 0 ? 'Cero' : 'Sin validar'; ctlStyle = st({ sz: 8, color: '996600' }); }
        }
      });
      if (hasPct) row[cPct] = e.pct == null ? t('', sBase) : { t: 'n', v: e.pct, z: '0.00%', s: st({ b: sub, top: sub, dbl: grand, h: 'right' }) };
      if (pdfVal != null) { row[cPdf] = { t: 'n', v: pdfVal, z: Number.isInteger(pdfVal) ? intFmt : '#,##0.00;-#,##0.00', s: st({ sz: 8, color: '666666' }) }; row[cDif] = { t: 'n', v: 0, z: '#,##0.00;-#,##0.00', f: difF, s: st({ sz: 8, color: '666666' }) }; }
      if (ctl) row[cCtl] = t(ctl, ctlStyle);
      row[cPag] = { t: 'n', v: e.page, s: st({ sz: 8, color: '999999', h: 'right' }) };
      push(row);
    }
    push([]);
    push([t('Los totales (negrita, con línea superior) son fórmulas sobre las cuentas de arriba; "En el PDF" muestra la cifra original para comparar.', st({ i: true, sz: 8, color: '666666' }))]);
    const widths = [60, 11, ...Array(maxCols).fill(17), ...(hasPct ? [9] : []), 2, 15, 11, 30, 6];
    return { ws: finish(aoa, widths, { freezeRow: headerRow }), rowOf };
  }

  // ---------- libro de una empresa ----------
  function buildEmpresaWorkbook(model, emp) {
    const wb = XLSX.utils.book_new();
    const used = new Set();
    const docs = emp.docs.filter((d) => !d.duplicado).sort((a, b) => a.fechaCorte.localeCompare(b.fechaCorte) || a.tipo.localeCompare(b.tipo));
    const sheets = [];
    const keyRows = [];
    for (const d of docs) {
      const lines = model.lineas.filter((l) => l.doc === d.idx);
      const clases = [...new Set(lines.map((l) => l.clase))];
      for (const clase of ['Balance', 'Resultados', 'Detalle balance', 'Detalle resultados', 'Determinación']) {
        if (!clases.includes(clase)) continue;
        const cl = lines.filter((l) => l.clase === clase);
        const entradas = d.estructura.filter((e) => e.clase === clase && e.tipo !== 't');
        const { ws, rowOf } = presSheet(model, d, clase, entradas, new Map(cl.map((l) => [l.n, l])));
        const name = safeName(claseSheet(d, clase), used);
        sheets.push({ name, ws });
        if (clase === 'Balance' || clase === 'Resultados') {
          for (const l of cl) {
            const nn = norm(l.cuenta);
            const k = d.tipo === 'PDT710'
              ? { 390: 'Total activo', 412: 'Total pasivo', 425: 'Total patrimonio', 426: 'Total pasivo y patrimonio', 463: 'Ventas netas', 492: 'Resultado del ejercicio (utilidad)', 493: 'Resultado del ejercicio (pérdida)' }[l.codigo]
              : /^TOTAL ACTIVO$/.test(nn) ? 'Total activo' : /^TOTAL PASIVO$/.test(nn) ? 'Total pasivo' : /^TOTAL PATRIMONIO$/.test(nn) ? 'Total patrimonio'
              : /^TOTAL PASIVO Y PATRIMONIO$/.test(nn) ? 'Total pasivo y patrimonio' : (clase === 'Resultados' && /^UTILIDAD DEL EJERCICIO|^PERDIDA DEL EJERCICIO/.test(nn)) ? 'Resultado del ejercicio (estado de resultados)'
              : (clase === 'Balance' && /^RESULTADO DEL EJERCICIO/.test(nn)) ? 'Resultado del ejercicio (balance)' : null;
            if (k && l.importe != null) keyRows.push({ d, k, name, row: rowOf.get(l.n)?.r, col: colL(rowOf.get(l.n)?.c ?? 2), importe: l.importe });
          }
        }
      }
    }
    // Resumen
    const inc = model.incidencias.filter((i) => i.empresa === emp.nombre || docs.some((d) => d.pdf === i.pdf)).sort((a, b) => SEV[a.severidad] - SEV[b.severidad]);
    const aoa = [[t(emp.nombre, S.title)], [t(`RUC: ${emp.ruc || '(no consta)'}   ·   Importes en Soles (S/) salvo indicación   ·   Generado con PDF a Excel (sin IA)`, S.sub)], []];
    aoa.push([t('Documentos procesados', S.sec)]);
    aoa.push(head(['PDF', 'Tipo', 'Periodo', 'Fecha de corte', 'Páginas', 'Pág. con cifras', 'Pág. informativas', 'Líneas extraídas', 'Subtotales que cuadran', 'Subtotales con diferencia', 'Incidencias']));
    const r0 = aoa.length + 1;
    for (const d of docs) {
      const info = d.paginas.filter((p) => p.informativa).length;
      const ls = model.lineas.filter((l) => l.doc === d.idx);
      aoa.push([t(d.pdf), t(d.tipo === 'PDT710' ? 'PDT 710' : 'EEFF'), t(d.periodoEtiqueta), t(d.fechaCorte), { t: 'n', v: d.paginas.length }, { t: 'n', v: d.paginas.length - info }, { t: 'n', v: info },
        { t: 'n', v: ls.length }, { t: 'n', v: ls.filter((l) => l.subtotal && l.validacion === 'ok').length }, { t: 'n', v: ls.filter((l) => l.validacion === 'diferencia').length },
        { t: 'n', v: model.incidencias.filter((i) => i.pdf === d.pdf).length }]);
    }
    const r1 = aoa.length;
    aoa.push([t('Totales', S.bold), t(''), t(''), t(''), ...[4, 5, 6, 7, 8, 9, 10].map((c) => ({ t: 'n', v: docs.length ? aoa.slice(r0 - 1, r1).reduce((s, row) => s + (row[c]?.v || 0), 0) : 0, f: `SUM(${colL(c)}${r0}:${colL(c)}${r1})`, s: S.bold }))]);
    aoa.push([]);
    aoa.push([t('Hojas de este libro (clic para ir)', S.sec)]);
    aoa.push(head(['Hoja', 'Contenido']));
    for (const sh of sheets) {
      const desc = /Balance Jul|Balance \d|PDT Balance/.test(sh.name) && !/Notas/.test(sh.name) ? 'Estado de situación financiera' : /Notas balance/.test(sh.name) ? 'Detalle (notas) del balance' : /Notas resultados/.test(sh.name) ? 'Detalle (notas) del estado de resultados'
        : /Resultados/.test(sh.name) ? 'Estado de resultados' : /Renta/.test(sh.name) ? 'Determinación del impuesto a la renta (PDT)' : '';
      aoa.push([{ t: 's', v: sh.name, l: { Target: `#'${sh.name}'!A1` }, s: { font: { color: { rgb: '0563C1' }, underline: true } } }, t(desc)]);
    }
    aoa.push([{ t: 's', v: 'Datos', l: { Target: "#'Datos'!A1" }, s: { font: { color: { rgb: '0563C1' }, underline: true } } }, t('Todas las líneas en formato plano')]);
    aoa.push([]);
    aoa.push([t('Cifras clave (enlazadas a las hojas de detalle)', S.sec)]);
    aoa.push(head(['Concepto', 'Periodo', 'Importe', 'Hoja de origen']));
    for (const k of keyRows) aoa.push([t(k.k), t(k.d.periodoEtiqueta), n(k.importe, null, `${q(k.name)}!${k.col}${k.row}`), t(`${k.name} · fila ${k.row} · ${k.d.pdf}`)]);
    aoa.push([]);
    aoa.push([t('Incidencias de esta empresa', S.sec)]);
    aoa.push(head(['Severidad', 'Tipo', 'PDF', 'Página', 'Detalle']));
    for (const i of inc) aoa.push([t(SEV_TXT[i.severidad], i.severidad === 'alta' ? S.bad : null), t(i.tipo), t(i.pdf), t(i.pagina), t(i.detalle)]);
    if (!inc.length) aoa.push([t('Sin incidencias.')]);
    const resumen = finish(aoa, [46, 22, 24, 20, 12, 14, 16, 14, 18, 18, 12]);
    XLSX.utils.book_append_sheet(wb, resumen, 'Resumen');
    for (const s of sheets) XLSX.utils.book_append_sheet(wb, s.ws, s.name);
    // Datos planos
    const dat = [head(['Empresa', 'RUC', 'Periodo', 'Fecha de corte', 'Documento', 'Estado', 'Sección', 'Cuenta (nombre original)', 'Cód.', 'Columna', 'Importe', 'Moneda', 'Subtotal', 'Validación', 'PDF', 'Pág.', 'Lectura'])];
    for (const l of model.lineas.filter((x) => x.empresaId === emp.id)) dat.push([t(l.empresa), t(emp.ruc), t(l.periodoEtiqueta), t(l.fechaCorte), t(l.tipo === 'PDT710' ? 'PDT 710' : 'EEFF'), t(l.clase), t(l.seccion), t(l.cuenta), t(l.codigo), t(l.columna),
      n(l.importe), t(l.moneda), t(l.subtotal ? 'Sí' : ''), t(l.validacion), t(l.pdf), { t: 'n', v: l.page }, t(l.source === 'ocr' ? 'OCR' : 'Texto del PDF')]);
    XLSX.utils.book_append_sheet(wb, finish(dat, [34, 14, 16, 12, 10, 18, 30, 50, 8, 22, 16, 8, 9, 11, 34, 6, 14], { freezeRow: 1 }), 'Datos');
    return wb;
  }

  // ---------- libro maestro ----------
  function buildMasterWorkbook(model) {
    const wb = XLSX.utils.book_new();
    const used = new Set(['resumen', 'inventario pdf', 'incidencias']);
    const ordered = [...model.incidencias].sort((a, b) => SEV[a.severidad] - SEV[b.severidad]);
    const emps = model.empresas;
    const sheetsOut = [];

    // Consolidados por grupo
    const consInfo = new Map();
    for (const g of model.grupos) {
      const name = safeName(`Consolidado ${g.tipo === 'PDT710' ? g.etiqueta.replace('PDT 710', 'PDT') : g.etiqueta}`, used);
      const es = g.empresas.map((id) => emps.find((e) => e.id === id));
      const aoa = [[t(`Consolidado grupal — ${g.etiqueta} — corte ${g.corte}`, S.title)],
        [t('Importes en Soles (S/), antes de eliminaciones intercompañía. Cada importe indica su PDF y página de origen. No se mezclan periodos.', S.sub)], [t(`Empresas incluidas: ${es.map((e) => e.nombre).join(' · ')}`, S.sub)], []];
      // Importes de todas las empresas juntos (se leen como un estado comparativo); el PDF y la página de origen van al final.
      const hdr = [t('Estado', S.head), t('Cód.', S.head), t('Cuenta (nombre original)', S.head)];
      es.forEach((e) => hdr.push(t(e.nombre, S.head)));
      hdr.push(t('Total grupo (antes de eliminar)', S.head), t('Observación', S.head), t('', S.head));
      es.forEach((e) => hdr.push(t(`Origen: ${e.nombre}`, S.head)));
      aoa.push(hdr);
      const first = aoa.length + 1;
      const rowOf = new Map();
      const amtCols = es.map((_, i) => 3 + i);
      g.rows.forEach((row, i) => {
        const r = first + i;
        rowOf.set(row.key, r);
        if (row.heading) {
          aoa.push([t(row.clase, st({ b: true, sz: 8, color: '666666' })), t(''), t(row.cuenta, st({ b: true }))]);
          return;
        }
        const sub = row.subtotal;
        const sb = st({ b: sub, top: sub, dbl: sub && GRAND.test(norm(row.cuenta)) });
        const line = [t(row.clase, st({ sz: 8, color: '666666' })), t(row.codigo, st({ sz: 9, color: '444444', h: 'center', b: sub, top: sub })), t(row.cuenta, st({ b: sub, top: sub, ind: row.lvl || 0, dbl: sub && GRAND.test(norm(row.cuenta)) }))];
        let tot = 0;
        const origenes = [];
        for (const e of es) {
          const v = row.vals.get(e.id);
          line.push(v ? { t: 'n', v: v.importe, z: Number.isInteger(v.importe) ? '#,##0;-#,##0' : '#,##0.00;-#,##0.00', s: sb } : t('', sb));
          origenes.push(t(v ? `${v.pdf} · p.${v.page}${v.enBlanco ? ' · (en blanco)' : ''}` : '', st({ sz: 8, color: '888888' })));
          tot += v ? v.importe : 0;
        }
        line.push({ t: 'n', v: tot, z: Number.isInteger(tot) ? '#,##0;-#,##0' : '#,##0.00;-#,##0.00', f: amtCols.map((c) => `N(${colL(c)}${r})`).join('+') || '0', s: st({ b: true, top: sub, dbl: sub && GRAND.test(norm(row.cuenta)) }) });
        const solo = row.vals.size === 1 && es.length > 1 && g.tipo !== 'PDT710' ? `Solo figura en ${emps.find((e) => e.id === [...row.vals.keys()][0]).nombre}` : '';
        const dif = [...row.vals.values()].some((v) => v.validacion === 'diferencia') ? 'Subtotal con diferencia en origen' : '';
        line.push(t([solo, dif].filter(Boolean).join(' · '), st({ sz: 8, color: '996600' })), t(''), ...origenes);
        aoa.push(line);
      });
      const totCol = 3 + es.length;
      // control de cuadre
      if (g.control) {
        const ra = rowOf.get(g.control.activo.key), rp = rowOf.get(g.control.pasivoPatrimonio.key);
        aoa.push([]);
        aoa.push([t('Control: total activo − total pasivo y patrimonio (debe ser 0)', S.bold), t(''), t(''), ...es.map(() => t('')),
          n((g.control.activo.vals.size ? [...g.control.activo.vals.values()].reduce((s, v) => s + v.importe, 0) : 0) - [...g.control.pasivoPatrimonio.vals.values()].reduce((s, v) => s + v.importe, 0), S.bold, `${colL(totCol)}${ra}-${colL(totCol)}${rp}`)]);
      }
      const widths = [11, 8, 50, ...es.map(() => 15), 18, 34, 2, ...es.map(() => 34)];
      consInfo.set(g.key, { name, rowOf, totCol, first });
      sheetsOut.push({ name, ws: finish(aoa, widths, { freezeRow: 5, freezeCol: 3 }) });
    }

    // Cruce CxC-CxP
    {
      const name = safeName('Cruce CxC-CxP', used);
      const aoa = [[t('Cruce de cuentas por cobrar y por pagar entre relacionadas', S.title)],
        [t('Solo se cruzan saldos de la misma fecha de corte y moneda. Una cuenta se elimina únicamente si coincide en monto con la cuenta recíproca de otra empresa y su naturaleza está sustentada.', S.sub)], []];
      aoa.push([t('1. Pares recíprocos encontrados', S.sec)]);
      aoa.push(head(['Corte', 'Empresa por cobrar', 'Cuenta por cobrar', 'Importe por cobrar', 'Empresa por pagar', 'Cuenta por pagar', 'Importe por pagar', 'Diferencia', 'Nivel de sustento', 'Estado', '¿Se elimina?', 'Sustento', 'Origen por cobrar', 'Origen por pagar']));
      if (!model.pares.length) aoa.push([t('No se encontró ningún par recíproco (ver sección 2 y la hoja Incidencias).', S.warn)]);
      for (const p of model.pares) {
        const r = aoa.length + 1;
        aoa.push([t(p.corte), t(p.cobrar.empresa), t(p.cobrar.cuenta), n(p.cobrar.importe), t(p.pagar.empresa), t(p.pagar.cuenta), n(p.pagar.importe), n(p.diff, null, `D${r}-G${r}`),
          t(p.nivel), t(p.estado), t(p.confirmada ? 'SÍ' : 'NO', p.confirmada ? S.ok : S.warn), t(p.sustento.join('; ')), t(`${p.cobrar.pdf} · p.${p.cobrar.page}`), t(`${p.pagar.pdf} · p.${p.pagar.page}`)]);
      }
      aoa.push([]);
      aoa.push([t('2. Todas las cuentas con relacionadas / accionistas detectadas', S.sec)]);
      aoa.push(head(['Corte', 'Empresa', 'Lado', 'Categoría', 'Cuenta (nombre original)', 'Sección / nota', 'Cód.', 'Importe', 'PDF', 'Pág.', 'Situación', 'Motivo']));
      const t0 = aoa.length + 1;
      const conf = new Set(model.pares.filter((p) => p.confirmada).flatMap((p) => [p.cobrar.gid, p.pagar.gid]));
      const nc = new Map(model.noConciliados.map((x) => [x.gid, x.motivo]));
      const cat = { relacionada: 'Relacionada', accionista: 'Accionista / socio / director', posible: 'Posible (revisar)' };
      for (const r of [...model.rel].sort((a, b) => a.fechaCorte.localeCompare(b.fechaCorte) || a.empresa.localeCompare(b.empresa))) {
        aoa.push([t(r.fechaCorte), t(r.empresa), t(r.lado === 'cobrar' ? 'Por cobrar' : r.lado === 'pagar' ? 'Por pagar' : 'Indeterminado'), t(cat[r.categoria]), t(r.cuenta), t(r.seccion), t(r.codigo), n(r.importe),
          t(r.pdf), { t: 'n', v: r.page }, t(conf.has(r.gid) ? 'Conciliada y eliminada' : 'NO CONCILIADA', conf.has(r.gid) ? S.ok : S.bad), t(conf.has(r.gid) ? '' : nc.get(r.gid) || '')]);
      }
      const t1 = aoa.length;
      if (!model.rel.length) aoa.push([t('No se detectaron cuentas con relacionadas ni accionistas.')]);
      aoa.push([]);
      aoa.push([t('3. Resumen por empresa y corte (fórmulas sobre la sección 2)', S.sec)]);
      aoa.push(head(['Corte', 'Empresa', 'Por cobrar relacionadas/accionistas', 'Por pagar relacionadas/accionistas', 'Neto']));
      for (const c of model.cortes) for (const e of emps) {
        if (!model.rel.some((r) => r.fechaCorte === c && r.empresaId === e.id)) continue;
        const r = aoa.length + 1;
        const rg = (col) => `$${col}$${t0}:$${col}$${t1}`;
        const sumif = (lado) => `SUMIFS(${rg('H')},${rg('A')},A${r},${rg('B')},B${r},${rg('C')},"${lado}")`;
        const v = (lado) => model.rel.filter((x) => x.fechaCorte === c && x.empresaId === e.id && x.lado === lado).reduce((s, x) => s + x.importe, 0);
        aoa.push([t(c), t(e.nombre), n(v('cobrar'), null, sumif('Por cobrar')), n(v('pagar'), null, sumif('Por pagar')), n(v('cobrar') - v('pagar'), null, `C${r}-D${r}`)]);
      }
      sheetsOut.push({ name, ws: finish(aoa, [12, 36, 44, 18, 36, 44, 18, 14, 16, 34, 26, 60, 34, 34]) });
    }

    // Eliminaciones
    let elimName;
    {
      elimName = safeName('Eliminaciones', used);
      const aoa = [[t('Detalle de eliminaciones y diferencias intercompañía', S.title)],
        [t('Solo se eliminan importes recíprocos sustentados. Todo saldo no conciliado permanece en el consolidado y se lista abajo.', S.sub)], []];
      aoa.push([t('1. Asientos de eliminación', S.sec)]);
      aoa.push(head(['Par', 'Corte', 'Lado', 'Empresa', 'Cuenta eliminada', 'Contraparte', 'Importe eliminado', 'Fila de estado financiero afectada', 'PDF', 'Pág.']));
      const a0 = aoa.length + 1;
      if (!model.eliminaciones.length) aoa.push([t('No hay eliminaciones: ningún par recíproco quedó sustentado.', S.warn)]);
      model.eliminaciones.forEach((e, i) => aoa.push([{ t: 'n', v: model.pares.findIndex((p) => p.id === e.parId) + 1 }, t(e.corte), t(e.lado === 'cobrar' ? 'Por cobrar' : 'Por pagar'), t(e.empresa), t(e.cuenta), t(e.contraparte), n(e.importe), t(e.filaEstado), t(e.pdf), { t: 'n', v: e.page }]));
      const a1 = aoa.length;
      aoa.push([t('Total eliminado', S.bold), t(''), t(''), t(''), t(''), t(''), n(model.eliminaciones.reduce((s, e) => s + e.importe, 0), S.bold, model.eliminaciones.length ? `SUM(G${a0}:G${a1})` : '0')]);
      aoa.push([]);
      aoa.push([t('2. Diferencias y saldos NO conciliados (permanecen visibles en el consolidado)', S.sec)]);
      aoa.push(head(['Corte', 'Empresa', 'Lado', 'Cuenta', 'Importe', 'Categoría', 'Motivo', 'PDF', 'Pág.']));
      if (!model.noConciliados.length) aoa.push([t('Sin saldos no conciliados.')]);
      for (const r of model.noConciliados) aoa.push([t(r.fechaCorte), t(r.empresa), t(r.lado === 'cobrar' ? 'Por cobrar' : r.lado === 'pagar' ? 'Por pagar' : 'Indeterminado'), t(r.cuenta), n(r.importe, S.warn), t({ relacionada: 'Relacionada', accionista: 'Accionista / socio / director', posible: 'Posible (revisar)' }[r.categoria]), t(r.motivo), t(r.pdf), { t: 'n', v: r.page }]);
      const diffs = model.pares.filter((p) => Math.abs(p.diff) > 0.005);
      if (diffs.length) {
        aoa.push([]);
        aoa.push([t('Diferencias menores dentro de la tolerancia en pares eliminados', S.sec)]);
        aoa.push(head(['Par', 'Corte', 'Por cobrar', 'Por pagar', 'Diferencia']));
        for (const p of diffs) aoa.push([t(p.id), t(p.corte), n(p.cobrar.importe), n(p.pagar.importe), n(p.diff)]);
      }
      sheetsOut.push({ name: elimName, ws: finish(aoa, [10, 12, 14, 34, 44, 34, 18, 44, 34, 6]) });
    }

    // Efectos (auxiliar para "antes y después")
    const efName = safeName('Efecto eliminaciones', used);
    const efAoa = [[t('Efecto de las eliminaciones en cada fila consolidada (auxiliar de "Antes vs Después")', S.title)], [t('Cada eliminación reduce su cuenta y todos los subtotales que la incluyen.', S.sub)], [], head(['Grupo (tipo|corte)', 'Clave de fila', 'Cuenta', 'Empresa', 'Importe', 'Par'])];
    const e0 = efAoa.length + 1;
    for (const e of model.efectos) efAoa.push([t(e.grupo), t(e.rowKey), t(e.cuenta), t(e.empresa), n(e.importe), t(e.parId)]);
    if (!model.efectos.length) efAoa.push([t(''), t(''), t('(sin eliminaciones)'), t(''), n(0)]);
    const e1 = efAoa.length;
    const efWs = finish(efAoa, [24, 50, 50, 34, 16, 40]);

    // Antes vs Después
    {
      const name = safeName('Antes vs Después', used);
      const aoa = [[t('Consolidado grupal antes y después de eliminaciones intercompañía', S.title)], [t('Antes = hoja consolidada. Eliminaciones = suma de la hoja "Efecto eliminaciones". Después = Antes + Eliminaciones.', S.sub)], []];
      for (const g of model.grupos) {
        const info = consInfo.get(g.key);
        aoa.push([t(`${g.etiqueta} — corte ${g.corte}`, S.sec)]);
        aoa.push(head(['Estado', 'Cód.', 'Cuenta', 'Antes de eliminaciones', 'Eliminaciones', 'Después de eliminaciones']));
        const b0 = aoa.length + 1;
        const idx = new Map();
        g.rows.forEach((row, i) => {
          if (row.heading) { aoa.push([t(row.clase, st({ sz: 8, color: '666666' })), t(''), t(row.cuenta, st({ b: true }))]); return; }
          const r = aoa.length + 1;
          idx.set(row.key, r);
          const before = [...row.vals.values()].reduce((s, v) => s + v.importe, 0);
          const el = model.efectos.filter((e) => e.grupo === g.key && e.rowKey === row.key).reduce((s, e) => s + e.importe, 0);
          const sub = row.subtotal;
          const sb = st({ b: sub, top: sub, dbl: sub && GRAND.test(norm(row.cuenta)) });
          const nm = (v, f) => ({ t: 'n', v, z: Number.isInteger(v) ? '#,##0;-#,##0' : '#,##0.00;-#,##0.00', ...(f ? { f } : {}), s: sb });
          aoa.push([t(row.clase, st({ sz: 8, color: '666666' })), t(row.codigo, st({ sz: 9, color: '444444', h: 'center' })), t(row.cuenta, st({ b: sub, top: sub, ind: row.lvl || 0 })),
            nm(before, `${q(info.name)}!${colL(info.totCol)}${info.rowOf.get(row.key)}`),
            nm(el, `SUMIFS(${q(efName)}!$E$${e0}:$E$${e1},${q(efName)}!$A$${e0}:$A$${e1},"${g.key}",${q(efName)}!$B$${e0}:$B$${e1},"${row.key.replace(/"/g, '""')}")`),
            nm(before + el, `D${r}+E${r}`)]);
        });
        if (g.control) {
          const ra = idx.get(g.control.activo.key), rp = idx.get(g.control.pasivoPatrimonio.key);
          const r = aoa.length + 1;
          aoa.push([t('Control de cuadre (activo − pasivo y patrimonio)', S.bold), t(''), t(''), n(0, S.bold, `D${ra}-D${rp}`), n(0, S.bold, `E${ra}-E${rp}`), n(0, S.bold, `F${ra}-F${rp}`)]);
        }
        aoa.push([]);
      }
      if (!model.grupos.length) aoa.push([t('No hay consolidados.')]);
      sheetsOut.push({ name, ws: finish(aoa, [11, 8, 52, 22, 20, 24], { freezeRow: 3 }) });
    }

    // Comparativo
    {
      const name = safeName('Comparativo', used);
      const aoa = [[t('Comparativo entre periodos (solo cifras comparables)', S.title)],
        [t('Se compara el balance entre cortes distintos de la misma empresa. Entre PDT y EEFF la equivalencia de cuentas es aproximada y se marca. El estado de resultados solo se compara si los periodos tienen la misma duración.', S.sub)], []];
      const cmp = model.comparativo;
      if (!cmp.empresas.length) for (const x of cmp.notas) aoa.push([t(x, S.warn)]);
      for (const c of cmp.empresas) {
        aoa.push([t(`${c.empresa.nombre}: ${c.d1.periodoEtiqueta} (${c.d1.fechaCorte}) vs ${c.d2.periodoEtiqueta} (${c.d2.fechaCorte})`, S.sec)]);
        aoa.push([t(c.nota, S.sub)]);
        if (c.aproximado) aoa.push([t('Equivalencia de cuentas aproximada entre PDT y EEFF: revisar antes de usar.', S.warn)]);
        aoa.push(head(['Cuenta', c.d1.periodoEtiqueta, c.d2.periodoEtiqueta, 'Variación', 'Variación %', `Origen ${c.d1.periodoEtiqueta}`, `Origen ${c.d2.periodoEtiqueta}`]));
        for (const f of c.filas) {
          const r = aoa.length + 1, st = f.subtotal ? S.sub_row : null;
          aoa.push([t(f.cuenta, st), n(f.v1, st), n(f.v2, st), n(f.v2 - f.v1, st, `C${r}-B${r}`), { t: 'n', v: f.v1 ? (f.v2 - f.v1) / Math.abs(f.v1) : 0, z: '0.0%', f: `IF(B${r}=0,0,D${r}/ABS(B${r}))` }, t(f.o1), t(f.o2)]);
        }
        aoa.push([]);
      }
      if (cmp.grupo && cmp.empresas.length > 1) {
        aoa.push([t(`Grupo (solo empresas comparables: ${cmp.grupo.perimetro.map((e) => e.nombre).join(', ')})`, S.sec)]);
        aoa.push(head(['Cuenta', 'Periodo inicial', 'Periodo final', 'Variación', 'Variación %']));
        for (const f of cmp.grupo.filas) { const r = aoa.length + 1; aoa.push([t(f.cuenta), n(f.v1), n(f.v2), n(f.v2 - f.v1, null, `C${r}-B${r}`), { t: 'n', v: f.v1 ? (f.v2 - f.v1) / Math.abs(f.v1) : 0, z: '0.0%', f: `IF(B${r}=0,0,D${r}/ABS(B${r}))` }]); }
      }
      aoa.push([]);
      aoa.push([t('Cobertura de documentos por empresa y corte', S.sec)]);
      aoa.push(head(['Empresa', ...model.cortes]));
      for (const e of emps) aoa.push([t(e.nombre), ...model.cortes.map((c) => { const ds = e.docs.filter((d) => !d.duplicado && d.fechaCorte === c); return t(ds.length ? ds.map((d) => (d.tipo === 'PDT710' ? 'PDT 710' : 'EEFF')).join(' + ') : '— sin documento —', ds.length ? null : S.warn); })]);
      sheetsOut.push({ name, ws: finish(aoa, [48, 20, 20, 18, 14, 34, 34]) });
    }

    // Incidencias
    const incAoa = [[t('Incidencias y datos pendientes de validación', S.title)], [t('Ordenadas por severidad. La columna "Revisión" es para anotar la validación.', S.sub)], [], head(['N°', 'Severidad', 'Tipo', 'Empresa', 'PDF', 'Página', 'Detalle', 'Revisión'])];
    ordered.forEach((i, k) => incAoa.push([{ t: 'n', v: k + 1 }, t(SEV_TXT[i.severidad], i.severidad === 'alta' ? S.bad : i.severidad === 'media' ? S.warn : null), t(i.tipo), t(i.empresa), t(i.pdf), t(i.pagina), t(i.detalle), t('')]));
    const i0 = 5, i1 = incAoa.length;
    const incWs = finish(incAoa, [6, 12, 16, 34, 36, 10, 110, 30], { freezeRow: 4 });

    // Inventario
    const inv = [[t('Inventario de PDF y control de procesamiento', S.title)], [t('Cobertura = cifras capturadas + cifras de encabezado/pie, sobre el total de cifras detectadas en las páginas con datos.', S.sub)], [],
      head(['PDF', 'Empresa', 'RUC', 'Tipo', 'Periodo', 'Fecha de corte', 'Páginas', 'Pág. con cifras', 'Pág. informativas', 'Cifras detectadas', 'Cifras capturadas', 'Cobertura', 'Líneas', 'Subtotales que cuadran', 'Con diferencia', 'Incidencias ALTAS', 'Resultado'])];
    const v0 = inv.length + 1;
    for (const d of model.docs) {
      const r = inv.length + 1;
      const pg = d.paginas.filter((p) => !p.informativa);
      const det = pg.reduce((s, p) => s + p.totalNums, 0), cap = pg.reduce((s, p) => s + p.leidos + (p.omitidos || 0), 0);
      const ls = d.lineas;
      const altas = model.incidencias.filter((i) => i.pdf === d.pdf && i.severidad === 'alta').length;
      const dif = ls.filter((l) => l.validacion === 'diferencia').length;
      const res = d.duplicado ? 'Duplicado: excluido' : !d.fechaCorte ? 'Sin fecha de corte: excluido' : altas || dif || cap < det ? 'Procesado con observaciones' : 'Procesado íntegramente';
      inv.push([t(d.pdf), t(d.empresa_.nombre), t(d.empresa_.ruc), t(d.tipo === 'PDT710' ? 'PDT 710' : d.tipo === 'EEFF' ? 'EEFF' : 'No reconocido'), t(d.periodoEtiqueta), t(d.fechaCorte), { t: 'n', v: d.paginas.length }, { t: 'n', v: pg.length }, { t: 'n', v: d.paginas.length - pg.length },
        { t: 'n', v: det }, { t: 'n', v: cap }, { t: 'n', v: det ? cap / det : 1, z: '0%', f: `IF(J${r}=0,1,K${r}/J${r})` }, { t: 'n', v: ls.length }, { t: 'n', v: ls.filter((l) => l.subtotal && l.validacion === 'ok').length }, { t: 'n', v: dif }, { t: 'n', v: altas },
        t(res, res === 'Procesado íntegramente' ? S.ok : S.warn)]);
    }
    const v1 = inv.length;
    inv.push([t('Totales', S.bold), ...Array(5).fill(t('')), ...[6, 7, 8, 9, 10].map((c) => ({ t: 'n', v: model.docs.length ? inv.slice(v0 - 1, v1).reduce((s, row) => s + (row[c]?.v || 0), 0) : 0, f: `SUM(${colL(c)}${v0}:${colL(c)}${v1})`, s: S.bold }))]);
    const invWs = finish(inv, [48, 36, 14, 12, 16, 12, 9, 12, 12, 12, 12, 10, 8, 14, 10, 12, 28], { freezeRow: 4 });

    // Resumen
    const cnt = (s) => model.incidencias.filter((i) => i.severidad === s).length;
    const pdfsOk = model.docs.filter((d) => !d.duplicado && d.fechaCorte).length;
    const R = [[t('Consolidado financiero — Resumen y controles', S.title)], [t('Generado con PDF a Excel (reglas, sin IA). Revise la hoja Incidencias antes de usar las cifras.', S.sub)], [],
      [t('Procesamiento', S.sec)],
      [t('PDF recibidos'), { t: 'n', v: model.docs.length, f: `COUNTA(${q('Inventario PDF')}!A5:A${4 + model.docs.length})` }],
      [t('PDF incorporados a los consolidados'), { t: 'n', v: pdfsOk }],
      [t('PDF excluidos (duplicados o sin fecha de corte)'), { t: 'n', v: model.docs.length - pdfsOk }],
      [t('Empresas identificadas'), { t: 'n', v: emps.length }],
      [t('Fechas de corte'), t(model.cortes.join(' · ') || '—')],
      [],
      [t('Hallazgos', S.sec)],
      [t('Saldos con relacionadas/accionistas detectados'), { t: 'n', v: model.rel.length }],
      [t('Pares recíprocos encontrados'), { t: 'n', v: model.pares.length }],
      [t('Pares eliminados (sustentados)'), { t: 'n', v: model.pares.filter((p) => p.confirmada).length }],
      [t('Saldos NO conciliados (siguen en el consolidado)'), { t: 'n', v: model.noConciliados.length }],
      [t('Total eliminado'), n(model.eliminaciones.reduce((s, e) => s + e.importe, 0), null, `${q(elimName)}!G${5 + 2 + 0 + 0}`)],
      [],
      [t('Incidencias', S.sec)],
      [t('ALTAS (corregir o validar antes de usar)'), { t: 'n', v: cnt('alta'), f: `COUNTIF(${q('Incidencias')}!B${i0}:B${Math.max(i1, i0)},"ALTA")` }],
      [t('MEDIAS'), { t: 'n', v: cnt('media'), f: `COUNTIF(${q('Incidencias')}!B${i0}:B${Math.max(i1, i0)},"MEDIA")` }],
      [t('Informativas'), { t: 'n', v: cnt('info'), f: `COUNTIF(${q('Incidencias')}!B${i0}:B${Math.max(i1, i0)},"Informativa")` }],
      [],
      [t('Empresas', S.sec)],
      ...emps.map((e) => [t(e.nombre), t(e.ruc || 'sin RUC'), t(`${e.docs.filter((d) => !d.duplicado).length} documento(s)`)]),
      [],
      [t('Hojas', S.sec)],
      ...['Consolidado por periodo: una hoja por (tipo de documento, corte)', 'Cruce CxC-CxP: pares recíprocos y todas las cuentas con relacionadas', 'Eliminaciones: asientos y saldos no conciliados', 'Antes vs Después: consolidado antes y después de eliminar', 'Comparativo: solo cifras comparables', 'Incidencias: datos pendientes de validación', 'Inventario PDF: control de cobertura por archivo', 'Datos fuente: todas las líneas con PDF y página'].map((x) => [t(x)])];
    // El total eliminado se enlaza a la fila "Total eliminado" de la hoja Eliminaciones (fila fija = encabezado + asientos).
    const elimTotalRow = 4 + 1 + Math.max(1, model.eliminaciones.length) + 1;
    R.find((row) => row[0]?.v === 'Total eliminado')[1] = n(model.eliminaciones.reduce((s, e) => s + e.importe, 0), null, `${q(elimName)}!G${elimTotalRow}`);

    // Datos fuente
    const dat = [head(['Empresa', 'RUC', 'Periodo', 'Fecha de corte', 'Documento', 'Estado', 'Sección', 'Cuenta (nombre original)', 'Cód.', 'Columna', 'Importe', 'Moneda', 'Subtotal', 'Validación', 'PDF', 'Pág.', 'Lectura'])];
    for (const l of model.lineas) dat.push([t(l.empresa), t(emps.find((e) => e.id === l.empresaId)?.ruc), t(l.periodoEtiqueta), t(l.fechaCorte), t(l.tipo === 'PDT710' ? 'PDT 710' : 'EEFF'), t(l.clase), t(l.seccion), t(l.cuenta), t(l.codigo), t(l.columna), n(l.importe), t(l.moneda), t(l.subtotal ? 'Sí' : ''), t(l.validacion), t(l.pdf), { t: 'n', v: l.page }, t(l.source === 'ocr' ? 'OCR' : 'Texto del PDF')]);

    XLSX.utils.book_append_sheet(wb, finish(R, [58, 26, 26]), 'Resumen');
    for (const s of sheetsOut.filter((x) => x.name.startsWith('Consolidado'))) XLSX.utils.book_append_sheet(wb, s.ws, s.name);
    for (const nm of ['Cruce CxC-CxP', 'Eliminaciones', 'Antes vs Después', 'Comparativo']) { const s = sheetsOut.find((x) => x.name === nm); if (s) XLSX.utils.book_append_sheet(wb, s.ws, s.name); }
    XLSX.utils.book_append_sheet(wb, incWs, 'Incidencias');
    XLSX.utils.book_append_sheet(wb, invWs, 'Inventario PDF');
    XLSX.utils.book_append_sheet(wb, efWs, efName);
    XLSX.utils.book_append_sheet(wb, finish(dat, [34, 14, 16, 12, 10, 18, 30, 50, 8, 22, 16, 8, 9, 11, 34, 6, 14], { freezeRow: 1 }), 'Datos fuente');
    return wb;
  }

  const toBytes = (wb) => XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  return { buildEmpresaWorkbook, buildMasterWorkbook, toBytes, fileSafe };
}

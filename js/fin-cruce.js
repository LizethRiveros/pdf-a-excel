// Consolidación e intercompañía a partir de documentos ya leídos (fin-parse.js).
// Reglas, no IA: todo lo que no se pueda sustentar queda visible como "no conciliado" o como incidencia.
import { norm, normEmpresa } from './fin-parse.js';

const GENERIC = new Set(['SAC', 'SA', 'SRL', 'SAA', 'EIRL', 'GRUPO', 'INVERSIONES', 'SERVICIOS', 'COMERCIAL', 'CORPORACION', 'INMOBILIARIA',
  'HOLDING', 'PERU', 'CINE', 'CINES', 'CINEMAS', 'PREMIUM', 'DEL', 'DE', 'LA', 'EL', 'Y', 'E', 'LOS', 'LAS', 'CONSTRUCTORA', 'INVERSION', 'NEGOCIOS', 'EMPRESA', 'COMPANIA', 'INDUSTRIAL']);
const firstToken = (nombre) => normEmpresa(nombre).split(' ').find((t) => t.length >= 4 && !GENERIC.has(t)) || '';

// Casillas del PDT 710 que son cuentas con relacionadas / accionistas.
const PDT_REL = {
  362: { lado: 'cobrar', categoria: 'relacionada', nombre: 'Ctas. por cobrar comerciales - relacionadas' },
  363: { lado: 'cobrar', categoria: 'accionista', nombre: 'Ctas. por cobrar al personal, accionistas y directores' },
  365: { lado: 'cobrar', categoria: 'relacionada', nombre: 'Ctas. por cobrar diversas - relacionadas' },
  405: { lado: 'pagar', categoria: 'relacionada', nombre: 'Ctas. por pagar comerciales - relacionadas' },
  406: { lado: 'pagar', categoria: 'accionista', nombre: 'Ctas. por pagar accionistas, socios y directores' },
  408: { lado: 'pagar', categoria: 'relacionada', nombre: 'Ctas. por pagar diversas - relacionadas' },
};
const RE_STRONG = /RELACIONAD|VINCULAD|AFILIAD|MATRIZ|SUBSIDIAR|INTERCOMP|EMPRESAS? DEL GRUPO|PARTES? RELACIONADAS?/;
const RE_SOCIO = /ACCIONISTA|\bSOCIOS?\b|DIRECTORES|PARTICIPACIONISTA/;
const RE_POSIBLE = /ACREENCIA|APORTES? (PARA|FUTUR)/;
const RE_COBRAR = /POR COBRAR|\bCOBRAR\b|PRESTAMOS? OTORGAD|ANTICIPOS? OTORGAD/;
const RE_PAGAR = /POR PAGAR|\bPAGAR\b|ACREENCIA|PRESTAMOS? RECIBID|OBLIGACION/;

// Equivalencias aproximadas PDT <-> EEFF para el comparativo entre periodos de distinto formato.
const ALIAS = [
  ['Efectivo y equivalentes', [359], /^EFECTIVO/],
  ['Cuentas por cobrar comerciales', [361, 362], /^CUENTAS POR COBRAR COMERCIALES/],
  ['Cuentas por cobrar a accionistas / personal', [363], /^CUENTAS POR COBRAR (AL PERSONAL|ACCIONISTAS)/],
  ['Cuentas por cobrar diversas', [364, 365], /^CUENTAS POR COBRAR DIVERSAS/],
  ['Mercaderías', [368], /^MERCADERIAS/],
  ['Propiedades, planta y equipo', [382], /^(INM\.? ?MAQ|PROPIEDAD(ES)?,? PLANTA)/],
  ['Depreciación acumulada', [383], /^DEPRECIACION ACUMULADA/],
  ['Intangibles', [384], /^INTANGIBLES/],
  ['Activo diferido', [388], /^ACTIVO DIFERIDO/],
  ['TOTAL ACTIVO', [390], /^TOTAL ACTIVO$/],
  ['Tributos y aportes por pagar', [402], /^TRIBUTOS/],
  ['Remuneraciones por pagar', [403], /^REMUNERACIONES/],
  ['Cuentas por pagar comerciales', [404, 405], /^CUENTAS POR PAGAR COM/],
  ['Cuentas por pagar diversas', [406, 407, 408], /^CUENTAS POR PAGAR DIVERSAS/],
  ['Obligaciones financieras', [409], /^OBLIGACIONES FINANCIERAS/],
  ['TOTAL PASIVO', [412], /^TOTAL PASIVO$/],
  ['Capital', [414], /^CAPITAL$/],
  ['Capital adicional', [416, 417], /^CAPITAL ADICIONAL/],
  ['TOTAL PATRIMONIO', [425], /^TOTAL PATRIMONIO$/],
  ['TOTAL PASIVO Y PATRIMONIO', [426], /^TOTAL PASIVO Y PATRIMONIO$/],
];

const round2 = (x) => Math.round(x * 100) / 100;

function sectionOverlap(a, b) {
  const ta = new Set(norm(a).split(' ').filter((t) => t.length > 3));
  return norm(b).split(' ').filter((t) => t.length > 3 && ta.has(t)).length;
}

export function buildModel(docsIn, opts = {}) {
  const tol = opts.tol ?? 1;
  const incidencias = [];
  const inc = (severidad, tipo, detalle, extra = {}) => incidencias.push({ severidad, tipo, pdf: '', pagina: '', empresa: '', detalle, ...extra });

  // ---------- documentos y empresas ----------
  const docs = docsIn.map((d, idx) => ({
    ...d, idx,
    empresaNombre: opts.empresaOverrides?.[idx] || d.empresa.nombre,
    fechaCorte: opts.fechaOverrides?.[idx] || d.fechaCorte,
  }));
  const empresas = [];
  const byRuc = new Map(), byName = new Map();
  for (const d of docs) {
    const ruc = d.empresa.ruc || '';
    const nameKey = normEmpresa(d.empresaNombre);
    let emp = (ruc && byRuc.get(ruc)) || byName.get(nameKey);
    if (emp && ruc && emp.ruc && emp.ruc !== ruc) {
      inc('alta', 'Empresa', `Mismo nombre "${d.empresaNombre}" con RUC distinto (${emp.ruc} y ${ruc}). Se tratan como empresas separadas.`, { pdf: d.pdf });
      emp = null;
    }
    if (!emp) { emp = { id: empresas.length + 1, nombre: d.empresaNombre, ruc, docs: [] }; empresas.push(emp); }
    if (ruc && !emp.ruc) { emp.ruc = ruc; if (d.tipo === 'PDT710') emp.nombre = d.empresaNombre; }
    if (ruc) byRuc.set(ruc, emp);
    byName.set(nameKey, emp);
    d.empresa_ = emp;
    emp.docs.push(d);
  }
  // Nombres que difieren solo por palabras extra ("CINE NORTE" vs "CINE NORTE PREMIUM"): se unen si hay un único candidato.
  const toks = (e) => normEmpresa(e.nombre).split(' ').filter((t) => t && !['SAC', 'SA', 'SRL', 'SAA', 'EIRL', 'S', 'A', 'C'].includes(t));
  for (const a of [...empresas]) {
    const ta = toks(a);
    if (ta.length < 2) continue;
    const cands = empresas.filter((b) => b !== a && toks(b).length > ta.length && ta.every((t) => toks(b).includes(t)) && (!a.ruc || !b.ruc || a.ruc === b.ruc));
    if (cands.length !== 1) continue;
    const b = cands[0];
    for (const d of a.docs) { d.empresa_ = b; b.docs.push(d); }
    if (a.ruc && !b.ruc) b.ruc = a.ruc;
    empresas.splice(empresas.indexOf(a), 1);
    inc('media', 'Empresa', `Se unieron "${a.nombre}" y "${b.nombre}" por nombre similar (los documentos de "${a.nombre}" no traen el nombre completo). Confirmar que son la misma empresa; si no lo son, corrija el nombre.`);
  }
  empresas.forEach((e, i) => { e.id = i + 1; });
  for (let i = 0; i < empresas.length; i++) for (let j = i + 1; j < empresas.length; j++) {
    const a = firstToken(empresas[i].nombre), b = firstToken(empresas[j].nombre);
    if (a && a === b) inc('media', 'Empresa', `"${empresas[i].nombre}" y "${empresas[j].nombre}" comparten el nombre "${a}": confirmar si son la misma empresa (se tratan como distintas).`);
  }

  // duplicados: misma empresa + tipo + fecha de corte
  const seen = new Map();
  for (const d of docs) {
    const k = `${d.empresa_.id}|${d.tipo}|${d.fechaCorte}`;
    if (seen.has(k)) {
      d.duplicado = true;
      inc('alta', 'Duplicidad', `"${d.pdf}" repite empresa, tipo y fecha de corte de "${seen.get(k)}". Se excluye de los consolidados para no duplicar importes.`, { pdf: d.pdf });
    } else seen.set(k, d.pdf);
  }
  for (const d of docs) for (const i of d.incidencias) incidencias.push({ ...i, empresa: d.empresa_.nombre });
  for (const d of docs) if (!d.fechaCorte) inc('alta', 'Periodo', `"${d.pdf}" no tiene fecha de corte: no entra en los consolidados hasta que se indique.`, { pdf: d.pdf });

  // ---------- líneas ----------
  const lineas = [];
  const byGid = new Map();
  for (const d of docs) {
    if (d.duplicado || !d.fechaCorte) continue;
    const local = new Map(d.lineas.map((l) => [l.n, l]));
    const keyCount = new Map();
    for (const l of d.lineas) {
      let cuenta = l.cuenta;
      if (cuenta === '(subtotal sin rótulo)' && l.calc?.refs?.length) {
        const f = local.get(l.calc.refs[0]), z = local.get(l.calc.refs.at(-1));
        cuenta = `Subtotal: ${f.cuenta}${z !== f ? ' … ' + z.cuenta : ''}`;
      }
      const L = {
        ...l, cuenta, gid: `${d.idx}:${l.n}`, doc: d.idx, empresaId: d.empresa_.id, empresa: d.empresa_.nombre, fechaCorte: d.fechaCorte,
        periodoEtiqueta: d.periodoEtiqueta, tipo: d.tipo === 'DESCONOCIDO' ? 'EEFF' : d.tipo, pdf: d.pdf, grupo: `${d.tipo}|${d.fechaCorte}`,
      };
      if (L.clase === 'Balance' || L.clase === 'Resultados') {
        let key = L.tipo === 'PDT710' ? `${L.codigo}` : `${L.clase}|${norm(L.cuenta)}`;
        const c = (keyCount.get(key) || 0) + 1;
        keyCount.set(key, c);
        if (c > 1) key += ` #${c}`;
        L.rowKey = key;
      }
      lineas.push(L);
      byGid.set(L.gid, L);
    }
  }

  // ---------- vínculo estado <-> nota (EEFF) ----------
  const nIncNota = new Set();
  for (const d of docs) {
    if (d.duplicado || d.tipo === 'PDT710') continue;
    const dl = lineas.filter((l) => l.doc === d.idx);
    const stmt = dl.filter((l) => l.clase === 'Balance' && !l.subtotal && l.importe != null);
    let block = [];
    for (const l of dl.filter((x) => x.clase === 'Detalle balance')) {
      if (!l.subtotal) { block.push(l); continue; }
      if (l.columna === '' && l.importe) {
        const cands = stmt.filter((s) => Math.abs(s.importe - l.importe) <= 1.5 && !s.notaLinked);
        const best = cands.length <= 1 ? cands[0] : cands.sort((a, b) => sectionOverlap(l.seccion, b.cuenta) - sectionOverlap(l.seccion, a.cuenta))[0];
        if (best) { best.notaLinked = true; best.notaTitulo = l.seccion; for (const b of block) { b.stmtGid = best.gid; b.notaTitulo = l.seccion; } }
      }
      block = [];
    }
    for (const s of stmt) {
      if (s.codigo && /^N/i.test(s.codigo) && !s.notaLinked && s.importe && !nIncNota.has(s.gid)) {
        inc('media', 'Validación', `La línea "${s.cuenta}" (${s.codigo}, ${s.importe}) no tiene una nota de detalle con el mismo saldo.`, { pdf: s.pdf, pagina: s.page, empresa: s.empresa });
      }
    }
  }

  // ---------- controles cruzados por documento ----------
  const sumBy = (arr, fn) => arr.reduce((s, l) => s + (l.importe ?? 0), 0);
  for (const d of docs) {
    if (d.duplicado || !d.fechaCorte) continue;
    const dl = lineas.filter((l) => l.doc === d.idx);
    const find = (clase, re) => dl.find((l) => l.clase === clase && l.subtotal && re.test(norm(l.cuenta)));
    if (d.tipo === 'PDT710') {
      const act = dl.find((l) => l.codigo === '390'), pp = dl.find((l) => l.codigo === '426');
      if (act && pp && Math.abs(act.importe - pp.importe) > 1.5) inc('alta', 'Control', `Balance PDT no cuadra: total activo ${act.importe} vs total pasivo y patrimonio ${pp.importe}.`, { pdf: d.pdf, empresa: d.empresa_.nombre });
      const u = dl.find((l) => l.codigo === '492'), p = dl.find((l) => l.codigo === '493'), bu = dl.find((l) => l.codigo === '423'), bp = dl.find((l) => l.codigo === '424');
      if (u && bu) {
        const er = (u.importe ?? 0) - (p?.importe ?? 0), bal = (bu.importe ?? 0) - (bp?.importe ?? 0);
        if (Math.abs(er - bal) > 1.5) inc('alta', 'Control', `El resultado del ejercicio del estado de resultados (${er}) no coincide con el del balance (${bal}).`, { pdf: d.pdf, empresa: d.empresa_.nombre });
      }
    } else {
      const act = find('Balance', /^TOTAL ACTIVO$/), pp = find('Balance', /^TOTAL PASIVO Y PATRIMONIO$/);
      if (act && pp && Math.abs(act.importe - pp.importe) > 1.5) inc('alta', 'Control', `Balance no cuadra: total activo ${act.importe} vs total pasivo y patrimonio ${pp.importe}.`, { pdf: d.pdf, empresa: d.empresa_.nombre });
      const er = find('Resultados', /^UTILIDAD DEL EJERCICIO|^RESULTADO DEL EJERCICIO|^PERDIDA DEL EJERCICIO/);
      const bal = dl.find((l) => l.clase === 'Balance' && /^RESULTADO DEL EJERCICIO/.test(norm(l.cuenta)));
      if (er && bal && Math.abs(er.importe - bal.importe) > 1.5) {
        inc('alta', 'Control', `El resultado del ejercicio difiere entre el estado de resultados (${er.importe}) y el balance (${bal.importe}); diferencia ${round2(er.importe - bal.importe)}. Revisar si falta algún gasto (p. ej. depreciación) en uno de los dos estados.`, { pdf: d.pdf, pagina: `${er.page} y ${bal.page}`, empresa: d.empresa_.nombre });
      }
      const me = dl.filter((l) => /\bME\b/.test(norm(l.cuenta)) && !l.subtotal).length;
      if (me) inc('info', 'Moneda', `Hay ${me} cuenta(s) en moneda extranjera (ME) ya expresadas en soles en el documento; el tipo de cambio no consta. No se reconvirtieron.`, { pdf: d.pdf, empresa: d.empresa_.nombre });
    }
  }

  // ---------- cuentas relacionadas ----------
  const tokens = empresas.map((e) => ({ id: e.id, tok: firstToken(e.nombre) }));
  const rel = [];
  for (const l of lineas) {
    if (l.subtotal || l.importe == null || l.importe === 0) continue;
    let r = null;
    if (l.tipo === 'PDT710') {
      if (l.clase === 'Balance' && PDT_REL[l.codigo]) r = { ...PDT_REL[l.codigo] };
    } else if (l.clase === 'Detalle balance' || (l.clase === 'Balance' && !l.notaLinked)) {
      const t = norm(`${l.seccion} ${l.cuenta}`);
      const own = norm(l.cuenta);
      let categoria = null, contraparte = null;
      for (const o of tokens) if (o.tok && o.id !== l.empresaId && new RegExp(`\\b${o.tok}\\b`).test(t)) contraparte = o.id;
      if (contraparte) categoria = 'relacionada';
      else if (RE_STRONG.test(t)) categoria = 'relacionada';
      else if (RE_SOCIO.test(t)) categoria = 'accionista';
      else if (RE_POSIBLE.test(own)) categoria = 'posible';
      if (categoria) {
        const c = RE_COBRAR.test(t), p = RE_PAGAR.test(t);
        const lado = l.clase === 'Balance' ? (/ACTIVO/.test(norm(l.seccion)) ? 'cobrar' : /PASIVO/.test(norm(l.seccion)) ? 'pagar' : 'indeterminado')
          : c && !p ? 'cobrar' : p && !c ? 'pagar' : 'indeterminado';
        r = { lado, categoria, contraparte, nombre: l.cuenta };
      }
    }
    if (r) rel.push({ ...r, gid: l.gid, empresaId: l.empresaId, empresa: l.empresa, fechaCorte: l.fechaCorte, tipo: l.tipo, cuenta: l.cuenta, seccion: l.notaTitulo || l.seccion,
      codigo: l.codigo, importe: l.importe, moneda: l.moneda || 'PEN', pdf: l.pdf, page: l.page, stmtGid: l.clase === 'Balance' ? l.gid : l.stmtGid || null });
  }
  // Préstamos otorgados a "terceros": no se eliminan, pero conviene confirmar que no sean relacionados.
  for (const l of lineas) {
    if (l.clase === 'Detalle balance' && !l.subtotal && /PRESTAMO/.test(norm(l.cuenta)) && /TERCEROS/.test(norm(l.seccion)) && Math.abs(l.importe ?? 0) >= 50000) {
      inc('info', 'Revisión', `"${l.cuenta}" (${l.importe}) está bajo "${l.seccion}". Se trata como tercero y no se elimina; confirmar que no sea una empresa relacionada.`, { pdf: l.pdf, pagina: l.page, empresa: l.empresa });
    }
  }
  for (const r of rel) {
    if (r.categoria === 'posible') inc('media', 'Revisión', `"${r.cuenta}" (${r.importe}) podría ser una operación con relacionadas (acreencias/aportes). No se elimina sin sustento.`, { pdf: r.pdf, pagina: r.page, empresa: r.empresa });
    if (r.lado === 'indeterminado') inc('media', 'Revisión', `No se pudo determinar si "${r.cuenta}" (${r.importe}) es por cobrar o por pagar.`, { pdf: r.pdf, pagina: r.page, empresa: r.empresa });
  }

  // ---------- cruce por fecha de corte ----------
  const cortes = [...new Set(lineas.map((l) => l.fechaCorte))].sort();
  const pares = [];
  const usados = new Set();
  const empresasPorCorte = new Map(cortes.map((c) => [c, new Set(docs.filter((d) => !d.duplicado && d.fechaCorte === c).map((d) => d.empresa_.id))]));
  for (const c of cortes) {
    const cobrar = rel.filter((r) => r.fechaCorte === c && r.lado === 'cobrar').sort((a, b) => b.importe - a.importe);
    const pagar = rel.filter((r) => r.fechaCorte === c && r.lado === 'pagar');
    for (const x of cobrar) {
      const cands = pagar.filter((y) => !usados.has(y.gid) && y.empresaId !== x.empresaId && y.moneda === x.moneda &&
        Math.abs(Math.abs(x.importe) - Math.abs(y.importe)) <= tol &&
        (x.contraparte == null || x.contraparte === y.empresaId) && (y.contraparte == null || y.contraparte === x.empresaId));
      if (!cands.length) continue;
      let y = null, ambiguo = false;
      if (cands.length === 1) y = cands[0];
      else {
        const expl = cands.filter((z) => x.contraparte === z.empresaId || z.contraparte === x.empresaId);
        if (expl.length === 1) y = expl[0]; else { ambiguo = true; y = cands[0]; }
      }
      const sustento = ['monto idéntico (dif. ≤ ' + tol + ')', 'naturaleza recíproca (por cobrar vs por pagar)', 'mismo corte ' + c, 'misma moneda'];
      if (x.contraparte === y.empresaId || y.contraparte === x.empresaId) sustento.push('la cuenta nombra a la contraparte');
      const nivel = !ambiguo && x.categoria === 'relacionada' && y.categoria === 'relacionada' ? 'Alto' : 'Medio';
      const id = `${x.gid}|${y.gid}`;
      const auto = nivel === 'Alto' && !ambiguo;
      const confirmada = opts.confirmed && id in opts.confirmed ? !!opts.confirmed[id] : auto;
      if (!ambiguo) usados.add(y.gid);
      if (!ambiguo) usados.add(x.gid);
      pares.push({ id, corte: c, cobrar: x, pagar: y, monto: Math.min(Math.abs(x.importe), Math.abs(y.importe)), diff: round2(Math.abs(x.importe) - Math.abs(y.importe)),
        nivel, ambiguo, sustento, confirmada, auto,
        estado: ambiguo ? 'Ambiguo: varias contrapartes posibles' : nivel === 'Alto' ? 'Conciliado' : 'Coincide en monto; falta confirmar naturaleza (accionista/posible)' });
    }
  }

  // No conciliados: todo lo relacionado que no quedó en un par confirmado
  const conf = new Set(pares.filter((p) => p.confirmada).flatMap((p) => [p.cobrar.gid, p.pagar.gid]));
  const noConciliados = [];
  for (const r of rel) {
    if (conf.has(r.gid)) continue;
    const par = pares.find((p) => p.cobrar.gid === r.gid || p.pagar.gid === r.gid);
    let motivo;
    if (par) motivo = par.ambiguo ? 'Hay varias contrapartes con el mismo monto: confirmar manualmente.' : 'Coincide en monto con otra cuenta, pero falta confirmar que la naturaleza sea intercompañía (marcar para eliminar).';
    else if (r.lado === 'indeterminado') motivo = 'No se pudo determinar si es por cobrar o por pagar.';
    else {
      const otras = empresasPorCorte.get(r.fechaCorte);
      const sinDoc = empresas.filter((e) => e.id !== r.empresaId && !otras.has(e.id));
      const opuesto = rel.filter((o) => o.fechaCorte === r.fechaCorte && o.empresaId !== r.empresaId && o.lado === (r.lado === 'cobrar' ? 'pagar' : 'cobrar') && o.moneda === r.moneda);
      const cerca = opuesto.map((o) => ({ o, rel: Math.abs(Math.abs(o.importe) - Math.abs(r.importe)) / Math.max(1, Math.abs(r.importe)) })).sort((a, b) => a.rel - b.rel)[0];
      if (cerca && cerca.rel <= 0.25) motivo = `Diferencia de ${round2(Math.abs(r.importe) - Math.abs(cerca.o.importe))} frente a "${cerca.o.cuenta}" de ${cerca.o.empresa} (${cerca.o.importe}). No se elimina.`;
      else if (opuesto.length) motivo = 'Existen saldos opuestos de otras empresas a este corte, pero ninguno coincide en monto.';
      else if (sinDoc.length === empresas.length - 1 && empresas.length > 1) motivo = `Ninguna otra empresa tiene documento al corte ${r.fechaCorte}: no hay contraparte contra la cual cruzar.`;
      else motivo = 'Sin contraparte registrada en las demás empresas a este corte.';
    }
    noConciliados.push({ ...r, motivo });
  }
  for (const c of cortes) {
    const pres = empresasPorCorte.get(c);
    const tiene = rel.some((r) => r.fechaCorte === c);
    const faltan = empresas.filter((e) => !pres.has(e.id));
    if (faltan.length && empresas.length > 1 && tiene) inc('media', 'Cruce', `Al corte ${c} no hay documento de: ${faltan.map((e) => e.nombre).join(', ')}. Los saldos con relacionadas de ese corte no pueden conciliarse con esas empresas.`);
  }
  if (cortes.length > 1) inc('info', 'Periodo', `Se detectaron ${cortes.length} fechas de corte (${cortes.join(', ')}). Cada una se consolida por separado: no se mezclan periodos.`);

  // ---------- eliminaciones y su efecto en filas consolidadas ----------
  const parentsOf = (docIdx) => {
    const m = new Map();
    for (const l of lineas) if (l.doc === docIdx && l.calc?.refs && l.clase === 'Balance') for (const r of l.calc.refs) { (m.get(r) ?? m.set(r, []).get(r)).push(l.n); }
    return m;
  };
  const eliminaciones = [];
  const efectos = [];
  const sinUbicar = [];
  for (const p of pares.filter((x) => x.confirmada)) {
    for (const side of ['cobrar', 'pagar']) {
      const r = p[side];
      const st = r.stmtGid ? byGid.get(r.stmtGid) : null;
      eliminaciones.push({ parId: p.id, corte: p.corte, lado: side, empresa: r.empresa, empresaId: r.empresaId, cuenta: r.cuenta, codigo: r.codigo, importe: -p.monto,
        contraparte: (side === 'cobrar' ? p.pagar : p.cobrar).empresa, pdf: r.pdf, page: r.page, filaEstado: st ? st.cuenta : '(no ubicada)', rowKey: st?.rowKey || '' });
      if (!st) { sinUbicar.push(r); continue; }
      const parents = parentsOf(st.doc);
      const seen = new Set([st.n]);
      const queue = [st.n];
      while (queue.length) { const n = queue.shift(); for (const pn of parents.get(n) || []) if (!seen.has(pn)) { seen.add(pn); queue.push(pn); } }
      for (const n of seen) {
        const L = byGid.get(`${st.doc}:${n}`);
        efectos.push({ grupo: L.grupo, rowKey: L.rowKey, cuenta: L.cuenta, empresa: r.empresa, importe: -p.monto, parId: p.id });
      }
    }
  }
  for (const r of sinUbicar) inc('alta', 'Eliminación', `La eliminación de "${r.cuenta}" (${r.empresa}) no pudo ubicarse en una fila del estado financiero; el consolidado "después" no la refleja.`, { pdf: r.pdf, pagina: r.page, empresa: r.empresa });

  // ---------- consolidados por (tipo, corte) ----------
  // Las filas siguen el orden y las secciones de los estados originales; las cuentas que solo existen
  // en una empresa se insertan junto a la cuenta que las precede en su propio documento.
  const grupos = [];
  const gmap = new Map();
  for (const d of docs) {
    if (d.duplicado || !d.fechaCorte) continue;
    const tipo = d.tipo === 'DESCONOCIDO' ? 'EEFF' : d.tipo;
    const gkey = `${d.tipo}|${d.fechaCorte}`;
    let g = gmap.get(gkey);
    if (!g) { g = { key: gkey, tipo, corte: d.fechaCorte, etiqueta: d.periodoEtiqueta, empresas: [], rows: [] }; gmap.set(gkey, g); grupos.push(g); }
    if (!g.empresas.includes(d.empresa_.id)) g.empresas.push(d.empresa_.id);
    let ptr = -1;
    const seenH = new Map();
    for (const e of d.estructura) {
      if (e.tipo === 't' || e.consumido || !(e.clase === 'Balance' || e.clase === 'Resultados')) continue;
      let key, make, line = null;
      if (e.tipo === 'h') {
        const hk = `H|${e.clase}|${norm(e.texto)}`;
        const c = (seenH.get(hk) || 0) + 1;
        seenH.set(hk, c);
        key = c > 1 ? `${hk} #${c}` : hk;
        make = () => ({ key, heading: true, cuenta: e.texto, codigo: '', clase: e.clase, subtotal: false, lvl: 0, vals: new Map() });
      } else {
        line = byGid.get(`${d.idx}:${e.lineIds[0]}`);
        if (!line) continue;
        key = line.rowKey;
        make = () => ({ key, cuenta: line.cuenta, codigo: line.codigo, clase: line.clase, subtotal: line.subtotal, lvl: e.lvl || 0, vals: new Map() });
      }
      let idx = g.rows.findIndex((r) => r.key === key);
      if (idx < 0) { idx = ptr + 1; g.rows.splice(idx, 0, make()); }
      ptr = idx;
      if (line) g.rows[idx].vals.set(d.empresa_.id, { importe: line.importe ?? 0, enBlanco: line.enBlanco, pdf: line.pdf, page: line.page, gid: line.gid, validacion: line.validacion });
    }
  }
  for (const g of grupos) g.rmap = new Map(g.rows.map((r) => [r.key, r]));
  grupos.sort((a, b) => a.corte.localeCompare(b.corte) || a.tipo.localeCompare(b.tipo));
  for (const g of grupos) {
    const act = g.rows.find((r) => g.tipo === 'PDT710' ? r.key === '390' : /Balance\|TOTAL ACTIVO$/.test(r.key));
    const pp = g.rows.find((r) => g.tipo === 'PDT710' ? r.key === '426' : /Balance\|TOTAL PASIVO Y PATRIMONIO$/.test(r.key));
    g.control = act && pp ? { activo: act, pasivoPatrimonio: pp } : null;
    const sinNombre = g.rows.filter((r) => !r.heading && g.tipo !== 'PDT710' && r.vals.size === 1 && g.empresas.length > 1).length;
    if (g.empresas.length > 1 && sinNombre) inc('info', 'Consolidación', `${g.etiqueta}: ${sinNombre} cuenta(s) existen en una sola empresa (los nombres de cuenta de los EEFF no son estándar); se muestran en su propia fila sin agrupar.`);
  }

  // ---------- comparativo entre periodos ----------
  const comparativo = { empresas: [], grupo: null, notas: [] };
  for (const e of empresas) {
    const ds = e.docs.filter((d) => !d.duplicado && d.fechaCorte).sort((a, b) => a.fechaCorte.localeCompare(b.fechaCorte));
    const cs = [...new Set(ds.map((d) => d.fechaCorte))];
    if (cs.length < 2) continue;
    const d1 = ds.find((d) => d.fechaCorte === cs[0]), d2 = ds.find((d) => d.fechaCorte === cs.at(-1));
    const L = (d, clase) => lineas.filter((l) => l.doc === d.idx && l.clase === clase && l.importe != null);
    const filas = [];
    if (d1.tipo === d2.tipo) {
      const m2 = new Map(L(d2, 'Balance').map((l) => [l.rowKey, l]));
      for (const a of L(d1, 'Balance')) { const b = m2.get(a.rowKey); if (b) filas.push({ cuenta: a.cuenta, codigo: a.codigo, v1: a.importe, v2: b.importe, o1: `${a.pdf} p.${a.page}`, o2: `${b.pdf} p.${b.page}`, subtotal: a.subtotal }); }
    } else {
      const pdt = d1.tipo === 'PDT710' ? d1 : d2, ee = d1.tipo === 'PDT710' ? d2 : d1;
      for (const [label, cas, re] of ALIAS) {
        const p = L(pdt, 'Balance').filter((l) => cas.includes(+l.codigo));
        const q = L(ee, 'Balance').filter((l) => re.test(norm(l.cuenta)) && l.cuenta);
        if (!p.length && !q.length) continue;
        const sp = p.reduce((s, l) => s + l.importe, 0), sq = q.reduce((s, l) => s + l.importe, 0);
        const op = p.map((l) => `${l.pdf} p.${l.page}`)[0] || '', oq = q.map((l) => `${l.pdf} p.${l.page}`)[0] || '';
        const first = d1 === pdt;
        filas.push({ cuenta: label, codigo: '', v1: first ? sp : sq, v2: first ? sq : sp, o1: first ? op : oq, o2: first ? oq : op, subtotal: /^TOTAL/.test(label), aproximado: true });
      }
    }
    const mismoMes = d1.tipo === d2.tipo && d1.fechaCorte.slice(5) === d2.fechaCorte.slice(5);
    comparativo.empresas.push({ empresa: e, d1, d2, filas, aproximado: d1.tipo !== d2.tipo,
      nota: mismoMes ? 'Estado de resultados comparable (mismo mes de corte): ver detalle en los consolidados de cada periodo.'
        : 'El estado de resultados no es comparable: los periodos acumulados tienen distinta duración. Solo se comparan saldos del balance entre cortes distintos.' });
  }
  if (comparativo.empresas.length) {
    const rows = new Map();
    for (const c of comparativo.empresas) for (const f of c.filas) {
      const k = f.cuenta;
      const r = rows.get(k) || { cuenta: f.cuenta, v1: 0, v2: 0, subtotal: f.subtotal, aproximado: f.aproximado };
      r.v1 += f.v1; r.v2 += f.v2; rows.set(k, r);
    }
    comparativo.grupo = { perimetro: comparativo.empresas.map((c) => c.empresa), filas: [...rows.values()] };
  } else comparativo.notas.push('No hay ninguna empresa con documentos en dos fechas de corte distintas, por lo que no existen cifras comparables. Cargue el mismo tipo de documento de la misma empresa en ambos periodos para obtener el comparativo.');

  return { docs, empresas, lineas, rel, pares, noConciliados, eliminaciones, efectos, grupos, cortes, comparativo, incidencias, tol };
}

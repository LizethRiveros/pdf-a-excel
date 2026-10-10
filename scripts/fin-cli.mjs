// Modo financiero por línea de comandos: procesa una carpeta (o lista) de PDF y genera
//   - un .xlsx por empresa
//   - un .xlsx maestro consolidado
// Uso: node scripts/fin-cli.mjs <carpeta-o-pdfs...> [--out salida] [--tol 1]
// Requiere: npm install   (pdfjs-dist y xlsx-js-style solo se usan aquí)
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { pdfTextToItems } from '../js/extract.js';
import { splitHidden, formsPorItem } from '../js/visibility.js';
import { parseDocument } from '../js/fin-parse.js';
import { buildModel } from '../js/fin-cruce.js';
import { createKit } from '../js/fin-excel.js';

const XLSX = createRequire(import.meta.url)('xlsx-js-style');
const args = process.argv.slice(2);
const opt = (name, def) => { const i = args.indexOf(name); if (i < 0) return def; const v = args[i + 1]; args.splice(i, 2); return v; };
const outDir = opt('--out', 'salida-financiera');
const tol = Number(opt('--tol', 1));
const inputs = args;
if (!inputs.length) { console.error('Uso: node scripts/fin-cli.mjs <carpeta-o-pdfs...> [--out salida] [--tol 1]'); process.exit(1); }

const files = inputs.flatMap((p) => (fs.statSync(p).isDirectory()
  ? fs.readdirSync(p).filter((f) => /\.pdf$/i.test(f)).sort().map((f) => path.join(p, f)) : [p]));

// Renderiza las páginas en gris con pdftoppm (poppler) para verificar qué texto se ve realmente.
const hayPdftoppm = spawnSync('pdftoppm', ['-v']).error == null;
const VIS_DPI = 108, VIS_SCALE = VIS_DPI / 72;
function renderGris(file) {
  if (!hayPdftoppm) return null;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'finvis-'));
  const r = spawnSync('pdftoppm', ['-r', String(VIS_DPI), '-gray', file, path.join(dir, 'p')]);
  if (r.status !== 0) return null;
  const out = new Map();
  for (const f of fs.readdirSync(dir)) {
    const m = f.match(/-(\d+)\.pgm$/);
    if (!m) continue;
    const buf = fs.readFileSync(path.join(dir, f));
    let pos = 0;
    const tok = () => { while (buf[pos] <= 32) pos++; let s = ''; while (buf[pos] > 32) s += String.fromCharCode(buf[pos++]); return s; };
    tok(); const w = +tok(), h = +tok(); tok(); pos++;
    out.set(+m[1], { w, h, px: buf.subarray(pos) });
  }
  fs.rmSync(dir, { recursive: true, force: true });
  return out;
}

async function leer(file) {
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(file)), verbosity: 0 }).promise;
  const pages = [];
  const gris = renderGris(file);
  if (!gris) console.warn('  (aviso) sin pdftoppm: no se verificó el texto oculto de ' + path.basename(file));
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    let items = pdfTextToItems(tc, vp.height);
    const forms = items.length ? formsPorItem(tc.items, await page.getOperatorList(), pdfjs.OPS) : null;
    if (forms) for (const it of items) it.form = forms[it.idx];
    let ocultos = [], dudosos = [];
    const g = gris?.get(p);
    if (g && items.length) { const r = splitHidden(items, g.px, g.w, g.h, VIS_SCALE, { channels: 1 }); items = r.visible; ocultos = r.hidden; dudosos = r.doubtful; }
    const chars = items.reduce((s, i) => s + i.str.length, 0);
    pages.push({ page: p, pageW: vp.width, pageH: vp.height, items, source: 'pdf', sinTexto: chars < 20, ocultos, dudosos, sinVerificar: !g });
  }
  return { pdf: path.basename(file), pages };
}

const docs = [];
const sinTexto = [];
for (const f of files) {
  process.stdout.write(`Leyendo ${path.basename(f)} ... `);
  const raw = await leer(f);
  const vacias = raw.pages.filter((p) => p.sinTexto).map((p) => p.page);
  if (vacias.length) sinTexto.push({ pdf: raw.pdf, paginas: vacias });
  const d = parseDocument(raw);
  docs.push(d);
  console.log(`${d.tipo} · ${d.empresa.nombre} · ${d.periodoEtiqueta} · ${d.lineas.length} líneas · ${d.incidencias.length} incidencias`);
}
const model = buildModel(docs, { tol });
for (const s of sinTexto) model.incidencias.push({ severidad: 'alta', tipo: 'OCR', pdf: s.pdf, pagina: s.paginas.join(', '), empresa: '', detalle: 'Páginas sin texto (escaneo). La línea de comandos no hace OCR: use la versión web (con OCR) para estas páginas.' });

const kit = createKit(XLSX);
fs.mkdirSync(outDir, { recursive: true });
const write = (name, wb) => { const f = path.join(outDir, name); fs.writeFileSync(f, Buffer.from(kit.toBytes(wb))); return f; };
const out = [];
for (const e of model.empresas) out.push(write(`${kit.fileSafe(e.nombre)}.xlsx`, kit.buildEmpresaWorkbook(model, e)));
out.push(write('CONSOLIDADO MAESTRO.xlsx', kit.buildMasterWorkbook(model)));

const cnt = (s) => model.incidencias.filter((i) => i.severidad === s).length;
console.log('\n===== INFORME =====');
console.log(`PDF procesados: ${docs.length} (incorporados: ${docs.filter((d) => !d.duplicado && d.fechaCorte).length})`);
console.log(`Empresas identificadas: ${model.empresas.length}\n  - ${model.empresas.map((e) => `${e.nombre} (${e.ruc || 'sin RUC'}) · ${e.docs.length} doc.`).join('\n  - ')}`);
console.log(`Cortes: ${model.cortes.join(', ')}`);
console.log(`Relacionadas/accionistas: ${model.rel.length} · pares recíprocos: ${model.pares.length} · eliminados: ${model.pares.filter((p) => p.confirmada).length} · NO conciliados: ${model.noConciliados.length}`);
console.log(`Incidencias: ${cnt('alta')} altas, ${cnt('media')} medias, ${cnt('info')} informativas`);
for (const i of model.incidencias.filter((x) => x.severidad === 'alta')) console.log(`  [ALTA] ${i.tipo} ${i.empresa ? '(' + i.empresa + ') ' : ''}${i.detalle}`);
console.log('\nArchivos generados:'); out.forEach((f) => console.log('  ' + f));

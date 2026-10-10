import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { pdfTextToItems } from '../js/extract.js';
import { parseDocument } from '../js/fin-parse.js';
import { buildModel } from '../js/fin-cruce.js';
import { createKit } from '../js/fin-excel.js';

const dir = new URL('../samples/financiero/', import.meta.url).pathname;
async function leer(file) {
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(path.join(dir, file))), verbosity: 0 }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    pages.push({ page: p, pageW: vp.width, pageH: vp.height, items: pdfTextToItems(await page.getTextContent(), vp.height), source: 'pdf' });
  }
  return { pdf: file, pages };
}
const docs = [];
for (const f of ['EEFF_ALFA_07-2026.pdf', 'EEFF_BETA_07-2026.pdf', 'EEFF_GAMMA_07-2026.pdf']) docs.push(parseDocument(await leer(f)));
const model = buildModel(docs);

test('identifica 3 empresas y un solo corte', () => {
  assert.deepEqual(model.empresas.map((e) => e.nombre), ['ALFA PERU SAC', 'BETA SERVICIOS SAC', 'GAMMA SAC']);
  assert.deepEqual(model.cortes, ['2026-07-31']);
});

test('todos los subtotales cuadran y los balances cuadran', () => {
  for (const d of docs) assert.equal(d.lineas.filter((l) => l.validacion === 'diferencia').length, 0, d.pdf);
  assert.equal(model.incidencias.filter((i) => i.severidad === 'alta').length, 0);
});

test('elimina solo el par recíproco exacto y deja visibles las diferencias', () => {
  assert.equal(model.pares.filter((p) => p.confirmada).length, 1);
  assert.equal(model.eliminaciones.length, 2);
  assert.ok(model.eliminaciones.every((e) => e.importe === -100000));
  assert.equal(model.noConciliados.length, 2);
  assert.ok(model.noConciliados.every((r) => [48000, 50000].includes(r.importe)));
});

test('la eliminación reduce activo y pasivo+patrimonio por igual', () => {
  const g = model.grupos[0];
  const sum = (key) => [...g.rows.find((r) => r.key === key).vals.values()].reduce((s, v) => s + v.importe, 0);
  const el = (key) => model.efectos.filter((e) => e.rowKey === key).reduce((s, e) => s + e.importe, 0);
  const act = 'Balance|TOTAL ACTIVO', pp = 'Balance|TOTAL PASIVO Y PATRIMONIO';
  assert.equal(sum(act) + el(act), sum(pp) + el(pp));
  assert.equal(sum(act) + el(act), 380000);
});

test('no cruza documentos de periodos distintos', () => {
  const otro = structuredClone(docs[1]);
  otro.fechaCorte = '2025-12-31'; otro.pdf = 'EEFF_BETA_12-2025.pdf';
  const m = buildModel([docs[0], otro]);
  assert.equal(m.pares.length, 0);
  assert.ok(m.incidencias.some((i) => i.tipo === 'Cruce'));
});

test('marca como duplicado el mismo documento cargado dos veces', () => {
  const copia = structuredClone(docs[0]); copia.pdf = 'copia.pdf';
  const m = buildModel([docs[0], copia]);
  assert.equal(m.docs.filter((d) => d.duplicado).length, 1);
  assert.equal(m.lineas.filter((l) => l.pdf === 'copia.pdf').length, 0);
});

test('genera los libros sin errores', () => {
  const XLSX = createRequire(import.meta.url)('xlsx-js-style');
  const kit = createKit(XLSX);
  const master = kit.buildMasterWorkbook(model);
  for (const nombre of ['Resumen', 'Consolidado EEFF Jul-2026', 'Cruce CxC-CxP', 'Eliminaciones', 'Antes vs Después', 'Comparativo', 'Incidencias', 'Inventario PDF']) assert.ok(master.SheetNames.includes(nombre), nombre);
  for (const e of model.empresas) assert.ok(kit.toBytes(kit.buildEmpresaWorkbook(model, e)).byteLength > 1000);
});

test('las hojas por empresa tienen formato de estado financiero y totales con fórmula', () => {
  const XLSX = createRequire(import.meta.url)('xlsx-js-style');
  const kit = createKit(XLSX);
  const wb = kit.buildEmpresaWorkbook(model, model.empresas[0]);
  const ws = wb.Sheets['Balance Jul-2026'];
  assert.ok(ws, 'existe la hoja Balance');
  const celdas = Object.entries(ws).filter(([k]) => /^[A-Z]+\d+$/.test(k));
  assert.ok(celdas.some(([, c]) => c.v === 'ALFA PERU SAC'), 'título con el nombre de la empresa');
  assert.ok(celdas.some(([, c]) => c.v === 'ESTADO DE SITUACION FINANCIERA'), 'título del estado');
  const totales = celdas.filter(([, c]) => c.f && c.f.startsWith('SUM('));
  assert.ok(totales.length >= 3, 'los totales son fórmulas');
  const totalActivo = celdas.find(([, c]) => c.v === 'TOTAL ACTIVO');
  assert.equal(totalActivo[1].s.font.bold, true);
  assert.equal(totalActivo[1].s.border.top.style, 'thin');
  const importes = celdas.filter(([k, c]) => k.startsWith('C') && c.t === 'n');
  assert.ok(importes.length >= 5, 'los importes son números, no texto');
});

test('el consolidado conserva el orden y las secciones del estado', () => {
  const g = model.grupos[0];
  const filas = g.rows.map((r) => (r.heading ? `H:${r.cuenta}` : r.cuenta));
  assert.ok(filas.indexOf('H:ACTIVO CORRIENTE') < filas.indexOf('TOTAL ACTIVO CORRIENTE'));
  assert.ok(filas.indexOf('TOTAL ACTIVO') < filas.indexOf('H:PASIVO CORRIENTE'));
});

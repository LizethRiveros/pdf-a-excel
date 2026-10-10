import { test } from 'node:test';
import assert from 'node:assert/strict';
import { extractTables, mergeAcrossPages } from '../js/extract.js';
import { parseNumber } from '../js/numbers.js';

// Genera fragmentos como los entrega pdf.js: una palabra por fragmento.
function grid(rowsText, colX, { y0 = 100, pitch = 16, h = 10, charW = 5.5 } = {}) {
  const items = [];
  rowsText.forEach((row, r) => {
    row.forEach((text, c) => {
      let x = colX[c];
      for (const word of text.split(' ')) {
        const w = word.length * charW;
        items.push({ str: word, x, y: y0 + r * pitch, w, h });
        x += w + charW;
      }
    });
  });
  return items;
}

test('detecta una tabla simple de 3 columnas', () => {
  const data = [
    ['Producto', 'Cantidad', 'Precio'],
    ['Cuaderno grande', '10', '25.50'],
    ['Lapiz', '200', '1.20'],
    ['Regla de metal', '15', '8.00'],
  ];
  const { tables } = extractTables(grid(data, [50, 200, 300]));
  assert.equal(tables.length, 1);
  assert.deepEqual(tables[0].rows, data);
});

test('el texto fuera de la tabla va a textLines', () => {
  const items = [
    ...grid([['Reporte de ventas del mes de marzo']], [50], { y0: 40 }),
    ...grid([['Item', 'Monto'], ['A', '10'], ['B', '20']], [50, 250], { y0: 120 }),
  ];
  const { tables, textLines } = extractTables(items);
  assert.equal(tables.length, 1);
  assert.equal(tables[0].rows.length, 3);
  assert.deepEqual(textLines, ['Reporte de ventas del mes de marzo']);
});

test('tolera celdas vacías', () => {
  const data = [
    ['Nombre', 'Notas', 'Total'],
    ['Ana', '', '100'],
    ['Luis', 'pendiente', '200'],
    ['Eva', '', '300'],
  ];
  const { tables } = extractTables(grid(data, [50, 180, 320]));
  assert.deepEqual(tables[0].rows, data);
});

test('dos tablas separadas por un salto grande quedan separadas', () => {
  const items = [
    ...grid([['A', 'B'], ['1', '2'], ['3', '4']], [50, 200], { y0: 50 }),
    ...grid([['Texto suelto que separa las tablas']], [50], { y0: 200 }),
    ...grid([['C', 'D'], ['5', '6'], ['7', '8']], [50, 200], { y0: 400 }),
  ];
  const { tables } = extractTables(items);
  assert.equal(tables.length, 2);
});

test('un párrafo normal no se toma como tabla', () => {
  const items = grid(
    [['Este es un parrafo de texto corrido que no es una tabla'],
     ['y continua en una segunda linea de texto normal tambien']],
    [50]
  );
  const { tables } = extractTables(items);
  assert.equal(tables.length, 0);
});

test('une tablas de páginas consecutivas y quita el encabezado repetido', () => {
  const t = (page, rows) => ({ page, rows });
  const merged = mergeAcrossPages([
    t(1, [['A', 'B'], ['1', '2']]),
    t(2, [['A', 'B'], ['3', '4']]),
    t(4, [['X', 'Y'], ['5', '6']]),
  ]);
  assert.equal(merged.length, 2);
  assert.deepEqual(merged[0].rows, [['A', 'B'], ['1', '2'], ['3', '4']]);
});

test('parseNumber', () => {
  assert.deepEqual(parseNumber('1,234.50'), { value: 1234.5, format: null });
  assert.deepEqual(parseNumber('S/ 1,200'), { value: 1200, format: null });
  assert.deepEqual(parseNumber('(300.00)'), { value: -300, format: null });
  assert.deepEqual(parseNumber('12%'), { value: 0.12, format: '0.00%' });
  assert.deepEqual(parseNumber('1.234,56', ','), { value: 1234.56, format: null });
  assert.equal(parseNumber('00123'), null);
  assert.equal(parseNumber('abc'), null);
  assert.equal(parseNumber('12-03-2024'), null);
});

test('une un rótulo con su monto cuando quedaron en líneas desfasadas (estados financieros)', () => {
  const items = [
    ...grid([['VENTAS', '01', '1,000'], ['COSTOS', '02', '-400']], [50, 250, 330], { y0: 100 }),
    // El rótulo queda medio renglón más abajo que su monto.
    ...grid([['', '', '600']], [50, 250, 330], { y0: 132 }),
    ...grid([['UTILIDAD BRUTA']], [50], { y0: 140 }),
    ...grid([['GASTOS', '03', '-100'], ['OTROS', '04', '-50']], [50, 250, 330], { y0: 164 }),
  ];
  const { tables } = extractTables(items);
  assert.equal(tables.length, 1);
  assert.deepEqual(tables[0].rows[2], ['UTILIDAD BRUTA', '', '600']);
  assert.equal(tables[0].rows.length, 5);
});

test('separa una tabla ancha que sigue a una de dos columnas y asigna encabezados centrados', () => {
  const items = [
    ...grid([['TOTAL A', '10'], ['TOTAL B', '20'], ['TOTAL C', '30'], ['TOTAL D', '40']], [50, 400], { y0: 100 }),
    ...grid([['DETALLE', 'SALDO', 'INICIAL', 'FINAL'],
             ['Barranca', '143,965', '85,280', '229,245'],
             ['Huaral', '173,089', '104,282', '277,371'],
             ['Otros', '317,054', '189,562', '506,616']], [50, 200, 300, 400], { y0: 190 }),
  ];
  const { tables } = extractTables(items);
  assert.equal(tables.length, 2);
  assert.equal(tables[1].rows[0].length, 4);
  assert.deepEqual(tables[1].rows[1], ['Barranca', '143,965', '85,280', '229,245']);
});

test('una tabla de dos lados con la mitad derecha vacía en algunas filas no se parte', () => {
  const full = (i) => [`Activo ${i}`, `${300 + i}`, '100', `Pasivo ${i}`, `${400 + i}`, '200'];
  const left = (i) => [`Activo ${i}`, `${300 + i}`, '100', '', '', ''];
  const rows = [full(1), full(2), full(3), left(4), left(5), left(6), left(7)];
  const { tables } = extractTables(grid(rows, [40, 190, 250, 320, 480, 540], { y0: 100 }));
  assert.equal(tables.length, 1);
  assert.equal(tables[0].rows.length, 7);
});

test('la tabla guarda el título que tiene encima', () => {
  const items = [
    ...grid([['NOTA 10 EFECTIVO']], [50], { y0: 60 }),
    ...grid([['CAJA', '1,141'], ['BANCOS', '11,061']], [50, 250], { y0: 90 }),
  ];
  const { tables } = extractTables(items);
  assert.equal(tables[0].title, 'NOTA 10 EFECTIVO');
});

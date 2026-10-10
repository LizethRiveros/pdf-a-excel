// Reconstruye tablas a partir de fragmentos de texto con posición.
// Es un módulo puro (sin DOM), se usa igual con texto de PDF y con palabras de OCR.
//
// Entrada: items = [{ str, x, y, w, h }]  (y = borde superior, el origen es arriba-izquierda)
// Salida:  { tables: [{ rows: string[][], bbox }], textLines: string[] }

const median = (arr) => {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};

// 1) Agrupa los fragmentos en líneas horizontales.
function groupLines(items, glyphH) {
  const sorted = items
    .filter((i) => i.str && i.str.trim() !== '')
    .map((i) => ({ ...i, yc: i.y + i.h / 2 }))
    .sort((a, b) => a.yc - b.yc || a.x - b.x);
  const lines = [];
  for (const it of sorted) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(it.yc - last.yc) <= glyphH * 0.5) {
      last.items.push(it);
      last.yc = last.items.reduce((s, i) => s + i.yc, 0) / last.items.length;
    } else {
      lines.push({ items: [it], yc: it.yc });
    }
  }
  for (const l of lines) l.items.sort((a, b) => a.x - b.x);
  return lines;
}

// 2) Junta palabras cercanas de una misma línea en celdas.
function buildCells(line, glyphH, wordGap) {
  const cells = [];
  for (const it of line.items) {
    const last = cells[cells.length - 1];
    const gap = last ? it.x - (last.x + last.w) : Infinity;
    if (last && gap < glyphH * wordGap) {
      last.str += (gap > glyphH * 0.1 ? ' ' : '') + it.str.trim();
      last.w = it.x + it.w - last.x;
      last.y = Math.min(last.y, it.y);
      last.h = Math.max(last.h, it.y + it.h - last.y);
      if (it.conf != null) last.conf = Math.min(last.conf ?? 100, it.conf);
    } else {
      cells.push({ str: it.str.trim(), x: it.x, y: it.y, w: it.w, h: it.h, ...(it.conf != null ? { conf: it.conf } : {}) });
    }
  }
  return cells;
}

// 3) Encuentra las separaciones entre columnas: franjas verticales que casi ninguna fila cubre.
function findGutters(rows, glyphH, minX, maxX) {
  const bin = glyphH / 8;
  const n = Math.max(1, Math.ceil((maxX - minX) / bin));
  const cover = new Array(n).fill(0);
  // Las filas de una sola celda (títulos, párrafos largos) no dicen nada sobre las columnas.
  const multi = rows.filter((cells) => cells.length >= 2);
  for (const cells of multi) {
    const seen = new Array(n).fill(false);
    for (const c of cells) {
      const a = Math.max(0, Math.floor((c.x - minX) / bin));
      const b = Math.min(n - 1, Math.floor((c.x + c.w - minX) / bin));
      for (let i = a; i <= b; i++) seen[i] = true;
    }
    for (let i = 0; i < n; i++) if (seen[i]) cover[i]++;
  }
  const tolerance = Math.floor(multi.length * (multi.length >= 7 ? 0.15 : 0));
  const minBins = Math.max(2, Math.round(glyphH * 0.4 / bin));
  const gutters = [];
  let start = -1;
  for (let i = 0; i <= n; i++) {
    const empty = i < n && cover[i] <= tolerance;
    if (empty && start < 0) start = i;
    if (!empty && start >= 0) {
      if (i - start >= minBins && start > 0 && i < n) {
        gutters.push(minX + ((start + i) / 2) * bin);
      }
      start = -1;
    }
  }
  return gutters;
}

// Si la cobertura encuentra menos columnas que celdas distintas hay en las filas más completas,
// esas filas sirven de plantilla: los límites son los puntos medios entre sus celdas.
function templateGutters(rows) {
  const multi = rows.filter((cells) => cells.length >= 2);
  const max = Math.max(...multi.map((cells) => cells.length));
  const full = multi.filter((cells) => cells.length === max);
  if (full.length < 2 || full.length < multi.length * 0.3) return null;
  const out = [];
  for (let k = 0; k < max - 1; k++) {
    out.push(median(full.map((cells) => (cells[k].x + cells[k].w + cells[k + 1].x) / 2)));
  }
  return out;
}

function bboxOf(cellsRows) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const cells of cellsRows) {
    for (const c of cells) {
      x0 = Math.min(x0, c.x);
      y0 = Math.min(y0, c.y);
      x1 = Math.max(x1, c.x + c.w);
      y1 = Math.max(y1, c.y + c.h);
    }
  }
  return { x0, y0, x1, y1 };
}

// Convierte el texto de una página de pdf.js en fragmentos con posición (origen arriba-izquierda).
export function pdfTextToItems(textContent, pageH) {
  const items = [];
  for (let idx = 0; idx < textContent.items.length; idx++) {
    const it = textContent.items[idx];
    if (!it.str || !it.str.trim()) continue;
    const [a, b, , d, e, f] = it.transform;
    if (Math.abs(b) > Math.abs(a) * 0.2) continue; // texto girado: se ignora
    const h = it.height || Math.abs(d);
    items.push({ str: it.str, x: e, y: pageH - f - h * 0.8, w: it.width, h, sz: Math.abs(a), idx });
  }
  return items;
}

// Agrupa los fragmentos en líneas con celdas, sin decidir todavía qué es tabla.
export function analyzeLines(items, opts = {}) {
  const wordGap = opts.wordGap ?? 0.7;
  const clean = items.filter((i) => i.str && i.str.trim() !== '' && i.h > 0);
  if (clean.length === 0) return { glyphH: 0, pitch: 0, lines: [] };
  const glyphH = median(clean.map((i) => i.h));
  const lines = groupLines(clean, glyphH);
  for (const l of lines) l.cells = buildCells(l, glyphH, wordGap);
  const pitch = median(lines.slice(1).map((l, i) => l.yc - lines[i].yc)) || glyphH * 1.4;
  return { glyphH, pitch, lines };
}

export function extractTables(items, opts = {}) {
  const { glyphH, pitch, lines } = analyzeLines(items, opts);
  if (lines.length === 0) return { tables: [], textLines: [] };
  const tabular = lines.map((l) => l.cells.length >= 2);

  // Tramos de líneas consecutivas que parecen tabla (se tolera 1 línea suelta en medio: texto envuelto).
  const runs = [];
  let i = 0;
  while (i < lines.length) {
    if (!tabular[i]) { i++; continue; }
    let j = i;
    while (j + 1 < lines.length) {
      const gap = lines[j + 1].yc - lines[j].yc;
      if (gap > pitch * 3) break;
      if (tabular[j + 1]) { j++; continue; }
      if (j + 2 < lines.length && tabular[j + 2] && lines[j + 2].yc - lines[j + 1].yc <= pitch * 2) { j += 2; continue; }
      if (j + 3 < lines.length && !tabular[j + 2] && tabular[j + 3] && lines[j + 3].yc - lines[j + 2].yc <= pitch * 2 &&
          lines[j + 2].yc - lines[j + 1].yc <= pitch * 2) { j += 3; continue; }
      break;
    }
    runs.push([i, j]);
    i = j + 1;
  }

  // Si a mitad de un tramo aparece una tabla más ancha (p. ej. de 2 a 5 columnas) que ya no vuelve
  // a la estructura anterior, son tablas distintas: se parten ahí, subiendo los encabezados de arriba.
  // Las filas con pocas celdas que alternan con filas llenas son celdas vacías, no tablas nuevas.
  const splitRun = (a, b) => {
    const idx = [];
    for (let k = a; k <= b; k++) if (tabular[k]) idx.push(k);
    const n = idx.map((k) => lines[k].cells.length);
    for (let k = 3; k + 3 <= idx.length; k++) {
      const before = n.slice(k - 3, k);
      const stable = Math.max(...before) - Math.min(...before) <= 1;
      if (stable && Math.min(...n.slice(k)) >= Math.max(...before) + 2) {
        let cut = idx[k], pulled = 0;
        while (pulled < 3 && cut - 1 > a && lines[cut].yc - lines[cut - 1].yc <= pitch * 1.6) { cut--; pulled++; }
        if (cut - 1 > a) return [[a, cut - 1], [cut, b]];
      }
    }
    return [[a, b]];
  };
  const pieces = opts.split === false ? runs : runs.flatMap(([a, b]) => splitRun(a, b));

  const tables = [];
  const used = new Set();
  for (const [a, b] of pieces) {
    const runLines = lines.slice(a, b + 1);
    const rowsCells = runLines.map((l) => l.cells);
    const { x0, x1 } = bboxOf(rowsCells);
    let gutters = findGutters(rowsCells, glyphH, x0, x1);
    const tpl = templateGutters(rowsCells);
    if (tpl && tpl.length > gutters.length) gutters = tpl;
    const nCols = gutters.length + 1;
    const colOf = (x) => { let k = 0; while (k < gutters.length && x >= gutters[k]) k++; return k; };
    const colWidth = (k) => (k < gutters.length ? gutters[k] : x1) - (k > 0 ? gutters[k - 1] : x0);
    if (nCols < 2 || runLines.length < 2) continue;

    const grid = rowsCells.map((cells) => {
      const row = new Array(nCols).fill('');
      for (const c of cells) {
        // Por el centro (los encabezados suelen estar centrados sobre números alineados a la derecha);
        // si la celda es más ancha que su columna (texto largo que desborda), por su inicio.
        const startCol = colOf(c.x);
        const centerCol = colOf(c.x + c.w / 2);
        const col = startCol === centerCol || c.w > colWidth(startCol) ? startCol : centerCol;
        row[col] = row[col] ? row[col] + ' ' + c.str : c.str;
      }
      return row;
    });

    // Une filas "partidas": un rótulo y su monto que quedaron en líneas casi iguales pero desfasadas.
    const ys = runLines.map((l) => l.yc);
    for (let r = 0; r + 1 < grid.length; ) {
      const A = grid[r], B = grid[r + 1];
      const complementary = A.some(Boolean) && B.some(Boolean) && A.every((v, c) => !(v && B[c]));
      if (complementary && ys[r + 1] - ys[r] < pitch * 0.75) {
        grid[r] = A.map((v, c) => v || B[c]);
        grid.splice(r + 1, 1);
        ys.splice(r + 1, 1);
      } else r++;
    }

    // Quita columnas completamente vacías.
    const keep = [];
    for (let c = 0; c < nCols; c++) if (grid.some((r) => r[c] !== '')) keep.push(c);
    if (keep.length < 2) continue;
    const rows = grid.map((r) => keep.map((c) => r[c]));
    let title = '';
    if (a > 0 && !tabular[a - 1] && lines[a].yc - lines[a - 1].yc <= pitch * 3.5) {
      title = lines[a - 1].cells.map((c) => c.str).join(' ').slice(0, 80);
    }
    tables.push({ rows, bbox: bboxOf(rowsCells), title });
    for (let k = a; k <= b; k++) used.add(k);
  }

  const textLines = lines
    .map((l, idx) => (used.has(idx) ? null : l.cells.map((c) => c.str).join(' ')))
    .filter((t) => t);

  return { tables, textLines };
}

// Une tablas consecutivas (p. ej. de páginas distintas) con el mismo número de columnas.
export function mergeAcrossPages(tables) {
  const out = [];
  for (const t of tables) {
    const prev = out[out.length - 1];
    if (prev && prev.page !== t.page && prev.rows[0].length === t.rows[0].length && t.page === prev.lastPage + 1) {
      const dupHeader = t.rows[0].join('|') === prev.rows[0].join('|');
      prev.rows = prev.rows.concat(dupHeader ? t.rows.slice(1) : t.rows);
      prev.lastPage = t.page;
    } else {
      out.push({ ...t, rows: t.rows.map((r) => [...r]), lastPage: t.page });
    }
  }
  return out;
}

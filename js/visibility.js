// Verifica que cada fragmento de texto del PDF se vea realmente en la página renderizada.
// Algunos PDF traen texto oculto o encimado (capas ocultas, texto blanco, contenido recortado) que
// extractores de texto leen como si fuera visible y mezclan cifras que no se ven en el documento.
// Un fragmento "sin tinta" donde dice estar se considera oculto y se descarta (con alerta).

// Asigna a cada fragmento de texto el "objeto de formulario" del PDF del que proviene. Una capa oculta suele ser
// otro objeto (un estado pegado detrás del visible), así que se puede descartar completa.
// opList = page.getOperatorList(); OPS = pdfjs.OPS. Devuelve un arreglo paralelo a textContent.items
// (id de formulario o null) o null si no se pudo alinear el texto con las operaciones.
export function formsPorItem(textItems, opList, OPS) {
  const stack = [0];
  let next = 1;
  const shows = [];
  for (let i = 0; i < opList.fnArray.length; i++) {
    const fn = opList.fnArray[i];
    if (fn === OPS.paintFormXObjectBegin) stack.push(next++);
    else if (fn === OPS.paintFormXObjectEnd) { if (stack.length > 1) stack.pop(); }
    else if (fn === OPS.showText || fn === OPS.showSpacedText || fn === OPS.nextLineShowText || fn === OPS.nextLineSetSpacingShowText) {
      const arr = opList.argsArray[i]?.[0] ?? opList.argsArray[i]?.[2] ?? [];
      const text = (Array.isArray(arr) ? arr : []).map((g) => (g && typeof g === 'object' ? g.unicode ?? '' : '')).join('');
      shows.push({ form: stack[stack.length - 1], text: text.replace(/\s+/g, '') });
    }
  }
  // Se alinea sobre el texto continuo: una operación puede traer varias palabras que pdf.js reparte en varios fragmentos.
  let flat = '';
  const formAt = [];
  for (const sh of shows) { flat += sh.text; for (let k = 0; k < sh.text.length; k++) formAt.push(sh.form); }
  const out = new Array(textItems.length).fill(null);
  let p = 0;
  for (let i = 0; i < textItems.length; i++) {
    const want = (textItems[i].str || '').replace(/\s+/g, '');
    if (!want) continue;
    let at = flat.startsWith(want, p) ? p : flat.indexOf(want, p);
    if (at < 0 || at - p > 400) return null; // no alineó: se usan las otras pruebas
    const forms = new Set(formAt.slice(at, at + want.length));
    if (forms.size !== 1) return null;
    out[i] = [...forms][0];
    p = at + want.length;
  }
  return out;
}

// pixels: Uint8Array/Uint8ClampedArray; channels: 1 (gris) o 4 (RGBA); scale: píxeles por unidad de página.
export function splitHidden(items, pixels, width, height, scale, { channels = 4, minInk = 0.05, doubtInk = 0.10, minLen = 3 } = {}) {
  const dark = (i) => (channels === 1 ? pixels[i] : (pixels[i * 4] + pixels[i * 4 + 1] + pixels[i * 4 + 2]) / 3) < 180;
  const measured = [];
  for (const it of items) {
    const s = it.str.trim();
    if (s.length < minLen) { measured.push({ it, ratio: 1, short: true }); continue; } // fragmentos cortos ("-", "01") tienen poca tinta y no son fiables para esta prueba
    const x0 = Math.max(0, Math.floor(it.x * scale)), x1 = Math.min(width - 1, Math.ceil((it.x + it.w) * scale));
    const y0 = Math.max(0, Math.floor(it.y * scale)), y1 = Math.min(height - 1, Math.ceil((it.y + it.h) * scale));
    let n = 0, d = 0;
    for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) { n++; if (dark(y * width + x)) d++; }
    measured.push({ it, ratio: n ? d / n : 1 });
  }
  // Una capa oculta suele tener otro tamaño de letra: si una parte importante de un grupo del mismo tamaño
  // no deja tinta, todo el grupo es de la capa oculta (aunque algunos elementos se solapen con texto visible).
  const groups = new Map();
  for (const m of measured) {
    if (m.short || m.it.sz == null) continue;
    const k = Math.round(m.it.sz * 2) / 2;
    (groups.get(k) ?? groups.set(k, []).get(k)).push(m);
  }
  // Capas por objeto de formulario del PDF (más fiable que el tamaño de letra).
  const byForm = new Map();
  for (const m of measured) {
    if (m.short || m.it.form == null) continue;
    (byForm.get(m.it.form) ?? byForm.set(m.it.form, []).get(m.it.form)).push(m);
  }
  const formStats = [...byForm.entries()].filter(([, g]) => g.length >= 5).map(([f, g]) => ({ f, frac: g.filter((m) => m.ratio < minInk).length / g.length }));
  const hiddenForms = new Set();
  if (formStats.some((x) => x.frac < 0.1)) for (const x of formStats) if (x.frac >= 0.2) hiddenForms.add(x.f);
  const hiddenGroups = new Set();
  for (const [k, g] of groups) if (g.length >= 5 && g.filter((m) => m.ratio < minInk).length / g.length >= 0.15) hiddenGroups.add(k);
  const visible = [], hidden = [], doubtful = [];
  for (const m of measured) {
    const k = m.it.sz != null ? Math.round(m.it.sz * 2) / 2 : null;
    if ((!m.short && (m.ratio < minInk || (k != null && hiddenGroups.has(k)))) || (m.it.form != null && hiddenForms.has(m.it.form))) hidden.push({ ...m.it, ink: m.ratio });
    else { visible.push(m.it); if (!m.short && m.ratio < doubtInk) doubtful.push({ ...m.it, ink: m.ratio }); }
  }
  return { visible, hidden, doubtful };
}

// Red de seguridad: pares de fragmentos visibles que ocupan el mismo lugar (texto encimado).
export function findOverlaps(items, { minShare = 0.4, maxPairs = 20 } = {}) {
  const sorted = items.filter((i) => i.str.trim().length >= 2).map((i) => ({ i, x0: i.x, x1: i.x + i.w, y0: i.y, y1: i.y + i.h })).sort((a, b) => a.x0 - b.x0);
  const out = [];
  for (let a = 0; a < sorted.length; a++) {
    for (let b = a + 1; b < sorted.length && sorted[b].x0 < sorted[a].x1; b++) {
      const A = sorted[a], B = sorted[b];
      const w = Math.min(A.x1, B.x1) - Math.max(A.x0, B.x0), h = Math.min(A.y1, B.y1) - Math.max(A.y0, B.y0);
      if (w <= 0 || h <= 0) continue;
      const small = Math.min((A.x1 - A.x0) * (A.y1 - A.y0), (B.x1 - B.x0) * (B.y1 - B.y0));
      if (small > 0 && (w * h) / small >= minShare) { out.push([A.i.str.trim(), B.i.str.trim()]); if (out.length >= maxPairs) return out; }
    }
  }
  return out;
}

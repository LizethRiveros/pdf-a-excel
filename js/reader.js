// Lectura de documentos (PDF con texto, PDF escaneado e imágenes) a fragmentos de texto con posición.
// Lo usan los dos modos de la herramienta: conversor de tablas y modo financiero.
import * as pdfjsLib from '../vendor/pdfjs/pdf.min.mjs';
import { pdfTextToItems } from './extract.js';
import { splitHidden, formsPorItem } from './visibility.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdfjs/pdf.worker.min.mjs', import.meta.url).href;

const OCR_SCALE = 2.5; // más resolución = mejor OCR, pero más lento
let ocrWorker = null;
let ocrLang = null;

async function getOcrWorker(lang, onProgress) {
  if (ocrWorker && ocrLang === lang) return ocrWorker;
  if (ocrWorker) await ocrWorker.terminate();
  const base = new URL('../vendor/', import.meta.url).href;
  ocrWorker = await Tesseract.createWorker(lang.split('+'), 1, {
    workerPath: base + 'tesseract/worker.min.js',
    corePath: base + 'tesseract',
    langPath: base + 'tessdata',
    gzip: true,
    workerBlobURL: false,
    logger: (m) => { if (m.status === 'recognizing text') onProgress?.(m.progress); },
  });
  ocrLang = lang;
  return ocrWorker;
}

export async function ocrCanvas(canvas, lang, onProgress) {
  const worker = await getOcrWorker(lang, onProgress);
  const { data } = await worker.recognize(canvas);
  return (data.words || [])
    .filter((w) => w.text && /[\p{L}\p{N}]/u.test(w.text) && w.confidence >= 30) // descarta rayas y ruido sin letras ni números
    .map((w) => ({ str: w.text, x: w.bbox.x0, y: w.bbox.y0, w: w.bbox.x1 - w.bbox.x0, h: w.bbox.y1 - w.bbox.y0, conf: w.confidence }));
}

export async function renderPdfPage(pdf, pageNo, scale) {
  const page = await pdf.getPage(pageNo);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  return canvas;
}

// Devuelve { pdf, pages: [{ page, pageW, pageH, items, source }] }
export async function readPdf(file, o, progress) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({
    data,
    cMapUrl: new URL('../vendor/pdfjs/cmaps/', import.meta.url).href,
    cMapPacked: true,
  }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    progress(`${file.name}: página ${p} de ${pdf.numPages}`, (p - 1) / pdf.numPages);
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    let items = pdfTextToItems(tc, vp.height);
    if (items.length && o.verificarVisibilidad !== false) {
      const forms = formsPorItem(tc.items, await page.getOperatorList(), pdfjsLib.OPS);
      if (forms) for (const it of items) it.form = forms[it.idx];
    }
    let pageW = vp.width, pageH = vp.height, source = 'pdf';
    let ocultos = [], dudosos = [];
    if (items.length && o.verificarVisibilidad !== false) {
      // Se descarta el texto que existe en el PDF pero no se ve en la página.
      const VIS_SCALE = 1.5;
      const cv = await renderPdfPage(pdf, p, VIS_SCALE);
      const img = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height);
      const r = splitHidden(items, img.data, cv.width, cv.height, VIS_SCALE, { channels: 4 });
      items = r.visible; ocultos = r.hidden; dudosos = r.doubtful;
    }
    const chars = items.reduce((s, i) => s + i.str.length, 0);
    const needOcr = o.ocr === 'force' || (o.ocr === 'auto' && chars < 20);
    if (needOcr) {
      progress(`${file.name}: página ${p} (leyendo con OCR, puede tardar)`, (p - 1) / pdf.numPages);
      const canvas = await renderPdfPage(pdf, p, OCR_SCALE);
      items = await ocrCanvas(canvas, o.lang, (f) => progress(`${file.name}: página ${p} (OCR ${Math.round(f * 100)}%)`, (p - 1 + f) / pdf.numPages));
      pageW = canvas.width; pageH = canvas.height; source = 'ocr';
    }
    pages.push({ page: p, pageW, pageH, items, source, ocultos, dudosos });
  }
  return { pdf, pages };
}

// Devuelve { bitmap, pages: [{ page: 1, ... }] }
export async function readImage(file, o, progress) {
  const bitmap = await createImageBitmap(file);
  const canvas = document.createElement('canvas');
  // Las imágenes pequeñas se amplían para que el OCR lea mejor.
  const k = Math.max(1, Math.min(3, 2000 / Math.max(bitmap.width, 1)));
  canvas.width = Math.round(bitmap.width * k);
  canvas.height = Math.round(bitmap.height * k);
  canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  progress(`${file.name}: leyendo con OCR`, 0);
  const items = await ocrCanvas(canvas, o.lang, (f) => progress(`${file.name}: OCR ${Math.round(f * 100)}%`, f));
  return { bitmap, pages: [{ page: 1, pageW: canvas.width, pageH: canvas.height, items, source: 'ocr' }] };
}

export const isPdf = (f) => f.type === 'application/pdf' || /\.pdf$/i.test(f.name);

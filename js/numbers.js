// Convierte textos como "S/ 1,234.50", "12%" o "(300.00)" en números de Excel.
// `decimal` es el separador decimal del documento: '.' (Perú, México, EE. UU.) o ','.

const CURRENCY = /^[\s]*(S\/\.?|US\$|USD|PEN|MXN|EUR|COP|CLP|ARS|[$€£¥])[\s]*/i;

export function parseNumber(text, decimal = '.') {
  if (typeof text !== 'string') return null;
  let s = text.trim();
  if (!s) return null;

  let negative = false;
  if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1).trim(); }
  s = s.replace(CURRENCY, '');
  if (s.startsWith('-') || s.startsWith('−')) { negative = !negative; s = s.slice(1).trim(); }
  if (s.endsWith('-')) { negative = !negative; s = s.slice(0, -1).trim(); }

  let percent = false;
  if (s.endsWith('%')) { percent = true; s = s.slice(0, -1).trim(); }

  const thousands = decimal === '.' ? ',' : '.';
  const re = new RegExp(
    '^(\\d{1,3}(\\' + thousands + '\\d{3})+|\\d+)(\\' + decimal + '\\d+)?$'
  );
  if (!re.test(s)) return null;
  // Los códigos con ceros a la izquierda ("00123") se quedan como texto.
  if (/^0\d/.test(s)) return null;

  const value = parseFloat(s.split(thousands).join('').replace(decimal, '.'));
  if (Number.isNaN(value)) return null;
  const signed = negative ? -value : value;
  return percent ? { value: Math.round(signed * 1e10 / 100) / 1e10, format: "0.00%" } : { value: signed, format: null };
}

/*
 * Excel-compatible number format engine for DataGrid cells.
 *
 * A grid cell's style object (cell.s) may carry:
 *   - numFmt:        'general' | 'number' | 'currency' | 'accounting'
 *                    | 'shortDate' | 'longDate' | 'time' | 'percentage'
 *                    | 'fraction' | 'scientific' | 'text'. Absent == 'general'.
 *   - decimals:      0..30. For 'fraction' this repurposes decimals as
 *                    the denominator digit count (1..3), matching Excel's
 *                    three built-in fraction variants.
 *   - currency:      symbol string ('$', 'NT$', '\u20ac', ...). Applies
 *                    to 'currency' / 'accounting'.
 *   - useThousands:  boolean; only meaningful for 'number' (currency /
 *                    accounting always group thousands).
 *   - negativeStyle: 'minus' | 'parens' | 'red' | 'redParens'. Applies
 *                    to 'number' / 'currency' (Accounting always uses
 *                    its own parens-with-alignment treatment).
 *
 * formatCellValue(v, s) renders the display string for the grid cell.
 * parseCellInput(text, s) converts user input back into a stored value
 * + optional formula flag, honouring the format (e.g. stripping '%').
 * toExcelNumFmtCode / fromExcelNumFmtCode bridge to OOXML numFmtId +
 * formatCode strings for the XlsxStyleInjector/Extractor round-trip.
 */

export const FORMAT_GENERAL = 'general';
export const FORMAT_NUMBER = 'number';
export const FORMAT_CURRENCY = 'currency';
export const FORMAT_ACCOUNTING = 'accounting';
export const FORMAT_SHORT_DATE = 'shortDate';
export const FORMAT_LONG_DATE = 'longDate';
export const FORMAT_TIME = 'time';
export const FORMAT_PERCENTAGE = 'percentage';
export const FORMAT_FRACTION = 'fraction';
export const FORMAT_SCIENTIFIC = 'scientific';
export const FORMAT_TEXT = 'text';

export const FORMAT_ORDER = [
  FORMAT_GENERAL,
  FORMAT_NUMBER,
  FORMAT_CURRENCY,
  FORMAT_ACCOUNTING,
  FORMAT_SHORT_DATE,
  FORMAT_LONG_DATE,
  FORMAT_TIME,
  FORMAT_PERCENTAGE,
  FORMAT_FRACTION,
  FORMAT_SCIENTIFIC,
  FORMAT_TEXT,
];

export const FORMAT_LABELS = {
  general: 'General',
  number: 'Number',
  currency: 'Currency',
  accounting: 'Accounting',
  shortDate: 'Short Date',
  longDate: 'Long Date',
  time: 'Time',
  percentage: 'Percentage',
  fraction: 'Fraction',
  scientific: 'Scientific',
  text: 'Text',
};

// Shortlist for the Accounting currency submenu. Mirrors Excel's default
// "Accounting Number Format" dropdown on en-US installs, with a couple
// of regional additions that GanttGen users frequently request.
export const CURRENCY_SHORTLIST = [
  { symbol: '$',       label: '$ English (United States)' },
  { symbol: 'NT$',     label: 'NT$ Chinese (Taiwan)' },
  { symbol: 'HK$',     label: 'HK$ Chinese (Hong Kong)' },
  { symbol: '\u00a5',  label: '\u00a5 Japanese Yen' },
  { symbol: 'CN\u00a5',label: 'CN\u00a5 Chinese RMB' },
  { symbol: '\u20ac',  label: '\u20ac Euro' },
  { symbol: '\u00a3',  label: '\u00a3 English (United Kingdom)' },
  { symbol: 'CHF',     label: 'CHF Swiss Franc' },
  { symbol: 'kr',      label: 'kr Scandinavian Krone' },
];

// ---------- Common helpers ------------------------------------------

function isFiniteNumber(v) {
  return typeof v === 'number' && Number.isFinite(v);
}

function effectiveDecimals(s, defaultValue = 2) {
  const d = s?.decimals;
  if (!Number.isFinite(d)) return defaultValue;
  return Math.max(0, Math.min(30, Math.round(d)));
}

function effectiveCurrency(s) {
  return s?.currency || '$';
}

function effectiveNegativeStyle(s) {
  return s?.negativeStyle || 'minus';
}

function formatAbsFixed(n, decimals, useThousands) {
  const abs = Math.abs(n);
  const fixed = abs.toFixed(decimals);
  if (!useThousands) return fixed;
  const [intPart, fracPart] = fixed.split('.');
  const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return fracPart != null ? grouped + '.' + fracPart : grouped;
}

function wrapNegative(body, n, style) {
  if (n >= 0) return body;
  switch (style) {
    case 'parens':
    case 'redParens':
      return '(' + body + ')';
    case 'red':
    case 'minus':
    default:
      return '-' + body;
  }
}

// Excel's legacy epoch. 1899-12-30 is the canonical zero because of
// Excel's Feb 29, 1900 bug: day 60 is "1900-02-29" (a non-existent
// date), day 61 is 1900-03-01. Using 1899-12-30 as the epoch maps day 1
// correctly to 1900-01-01 and days >= 61 match Excel exactly. Serials
// < 60 are one day off vs Excel's display; acceptable for modern data.
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

// ---------- formatCellValue -----------------------------------------

export function formatCellValue(value, s) {
  const fmt = s?.numFmt || FORMAT_GENERAL;

  if (fmt === FORMAT_TEXT) {
    return value == null ? '' : String(value);
  }

  if (value == null || value === '') return '';

  if (fmt === FORMAT_GENERAL) {
    return formatGeneral(value);
  }

  if (
    fmt === FORMAT_NUMBER ||
    fmt === FORMAT_CURRENCY ||
    fmt === FORMAT_ACCOUNTING ||
    fmt === FORMAT_PERCENTAGE ||
    fmt === FORMAT_SCIENTIFIC ||
    fmt === FORMAT_FRACTION
  ) {
    const n = toFiniteNumber(value);
    if (n == null) return String(value);

    switch (fmt) {
      case FORMAT_NUMBER: {
        const decimals = effectiveDecimals(s);
        const useThousands = s?.useThousands !== false;
        const body = formatAbsFixed(n, decimals, useThousands);
        return wrapNegative(body, n, effectiveNegativeStyle(s));
      }
      case FORMAT_CURRENCY: {
        const decimals = effectiveDecimals(s);
        const sym = effectiveCurrency(s);
        const body = sym + formatAbsFixed(n, decimals, true);
        return wrapNegative(body, n, effectiveNegativeStyle(s));
      }
      case FORMAT_ACCOUNTING: {
        const decimals = effectiveDecimals(s);
        const sym = effectiveCurrency(s);
        // Accounting uses a fixed 4-section pattern on-screen:
        //   positive:  "SYM  1,234.56 "
        //   negative:  "SYM  (1,234.56)"
        //   zero:      "SYM   -     "   (dash right-padded by decimals)
        // The trailing spaces keep adjacent cells visually aligned when
        // the column is right-padded via hAlign:'right'.
        if (n === 0) {
          const dashPad = decimals > 0 ? ' '.repeat(decimals + 1) : ' ';
          return sym + '\u00a0 -' + dashPad;
        }
        const abs = formatAbsFixed(n, decimals, true);
        if (n < 0) return sym + '\u00a0(' + abs + ')';
        return sym + '\u00a0' + abs + '\u00a0';
      }
      case FORMAT_PERCENTAGE: {
        const decimals = effectiveDecimals(s, 0);
        const body = formatAbsFixed(n * 100, decimals, false);
        return wrapNegative(body + '%', n, 'minus');
      }
      case FORMAT_SCIENTIFIC: {
        const decimals = effectiveDecimals(s);
        return formatScientific(n, decimals);
      }
      case FORMAT_FRACTION: {
        const d = s?.decimals;
        const denomDigits = d === 2 ? 2 : d === 3 ? 3 : 1;
        return formatFraction(n, denomDigits);
      }
      default:
        return String(value);
    }
  }

  const d = coerceToDate(value);
  if (!d) return String(value);
  switch (fmt) {
    case FORMAT_SHORT_DATE: return formatShortDate(d);
    case FORMAT_LONG_DATE:  return formatLongDate(d);
    case FORMAT_TIME:       return formatTime(d);
    default:                return String(value);
  }
}

function formatGeneral(value) {
  if (isFiniteNumber(value)) {
    return String(value);
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatShortDate(toUtcWall(value));
  }
  return String(value);
}

function toFiniteNumber(v) {
  if (isFiniteNumber(v)) return v;
  if (typeof v === 'string') {
    const trimmed = v.trim();
    if (trimmed === '') return null;
    const n = Number(trimmed);
    if (Number.isFinite(n)) return n;
  }
  if (v instanceof Date && !isNaN(v.getTime())) {
    return (v.getTime() - EXCEL_EPOCH_MS) / 86400000;
  }
  return null;
}

function formatScientific(n, decimals) {
  if (!Number.isFinite(n)) return String(n);
  const str = n.toExponential(decimals);
  // toExponential uses "e+7" / "e-3" without zero-padding. Excel renders
  // the exponent with at least two digits ("E+07" / "E-03").
  return str.replace(/e([+-])(\d+)$/, (_m, sign, digits) =>
    'E' + sign + (digits.length < 2 ? '0' + digits : digits));
}

function formatFraction(n, denomDigits) {
  if (!Number.isFinite(n)) return String(n);
  if (n === 0) return '0';
  const maxDenom = Math.pow(10, denomDigits) - 1;
  const sign = n < 0 ? -1 : 1;
  const absN = Math.abs(n);
  const whole = Math.floor(absN);
  const frac = absN - whole;
  if (frac < 1e-12) return String(sign * whole);

  // Continued-fraction approximation. h_i/k_i tracks the best rational
  // approximation seen so far under the denominator bound.
  let h0 = 0, h1 = 1, k0 = 1, k1 = 0;
  let x = frac;
  for (let i = 0; i < 60; i++) {
    const a = Math.floor(x);
    const h2 = a * h1 + h0;
    const k2 = a * k1 + k0;
    if (k2 > maxDenom) break;
    h0 = h1; h1 = h2;
    k0 = k1; k1 = k2;
    const rem = x - a;
    if (rem < 1e-12) break;
    x = 1 / rem;
  }

  const num = h1;
  const den = k1;

  if (num === 0) return String(sign * whole);
  const signStr = sign < 0 ? '-' : '';
  if (whole === 0) return signStr + num + '/' + den;
  return signStr + whole + ' ' + num + '/' + den;
}

function coerceToDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return toUtcWall(v);
  if (isFiniteNumber(v)) {
    const ms = EXCEL_EPOCH_MS + v * 86400000;
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (iso) {
      return new Date(Date.UTC(
        Number(iso[1]),
        Number(iso[2]) - 1,
        Number(iso[3]),
        Number(iso[4] || 0),
        Number(iso[5] || 0),
        Number(iso[6] || 0),
      ));
    }
    const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (mdy) {
      return new Date(Date.UTC(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2])));
    }
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return toUtcWall(parsed);
    }
  }
  return null;
}

// Return a Date whose UTC fields equal the input's local wall-clock
// fields. Keeps all downstream formatting branches timezone-neutral.
function toUtcWall(d) {
  return new Date(Date.UTC(
    d.getFullYear(), d.getMonth(), d.getDate(),
    d.getHours(), d.getMinutes(), d.getSeconds(),
  ));
}

function pad2(n) { return n < 10 ? '0' + n : String(n); }

function formatShortDate(d) {
  return (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + '/' + d.getUTCFullYear();
}

const DAY_NAMES = [
  'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
];
const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

function formatLongDate(d) {
  return DAY_NAMES[d.getUTCDay()] + ', ' + MONTH_NAMES[d.getUTCMonth()]
    + ' ' + d.getUTCDate() + ', ' + d.getUTCFullYear();
}

function formatTime(d) {
  const h = d.getUTCHours();
  const m = d.getUTCMinutes();
  const s = d.getUTCSeconds();
  const ampm = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return h12 + ':' + pad2(m) + ':' + pad2(s) + ' ' + ampm;
}

// ---------- parseCellInput ------------------------------------------

/**
 * Parse user-entered text into { value, formula }. `formula` is non-null
 * when the raw text starts with '=' and the format permits formulas
 * (i.e. not Text). For Text format both the '=' prefix and any numeric
 * coercion are suppressed so the cell always stores the literal string.
 */
export function parseCellInput(rawText, s) {
  const fmt = s?.numFmt || FORMAT_GENERAL;
  const text = rawText == null ? '' : String(rawText);

  if (text === '') return { value: '', formula: null };

  if (fmt === FORMAT_TEXT) {
    return { value: text, formula: null };
  }

  if (text.startsWith('=')) {
    return { value: null, formula: text };
  }

  if (fmt === FORMAT_PERCENTAGE) {
    const pct = text.trim().match(/^(.*?)\s*%\s*$/);
    if (pct) {
      const n = parseFinancialNumber(pct[1]);
      if (n != null) return { value: n / 100, formula: null };
    }
    const n = parseFinancialNumber(text);
    if (n != null) return { value: n, formula: null };
  }

  if (
    fmt === FORMAT_CURRENCY ||
    fmt === FORMAT_ACCOUNTING ||
    fmt === FORMAT_NUMBER ||
    fmt === FORMAT_SCIENTIFIC ||
    fmt === FORMAT_FRACTION
  ) {
    const n = parseFinancialNumber(text);
    if (n != null) return { value: n, formula: null };
  }

  if (fmt === FORMAT_SHORT_DATE || fmt === FORMAT_LONG_DATE) {
    const d = coerceToDate(text);
    if (d) {
      const iso = d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1)
        + '-' + pad2(d.getUTCDate());
      return { value: iso, formula: null };
    }
  }
  if (fmt === FORMAT_TIME) {
    // Store the user's raw time text; the formatter will re-render it.
    const trimmed = text.trim();
    const d = coerceToDate(trimmed);
    if (d) return { value: trimmed, formula: null };
  }

  // General fallback: match the legacy commitEdit coercion.
  const asNum = Number(text);
  if (text.trim() !== '' && Number.isFinite(asNum)) {
    return { value: asNum, formula: null };
  }
  return { value: text, formula: null };
}

function parseFinancialNumber(raw) {
  if (raw == null) return null;
  let s = String(raw).trim();
  if (s === '') return null;

  let negative = false;
  const paren = s.match(/^\((.*)\)$/);
  if (paren) {
    negative = true;
    s = paren[1].trim();
  }

  if (s.startsWith('-')) { negative = !negative; s = s.slice(1).trim(); }
  else if (s.startsWith('+')) { s = s.slice(1).trim(); }

  // Drop any leading sigil run (currency symbols, spaces) before the
  // first digit / dot / minus. This covers '$', 'NT$', 'CHF', 'kr ',
  // '\u00a5', etc. Keeps the regex permissive; anything without digits
  // ends up rejected below.
  s = s.replace(/^[^\d.-]+/, '');
  // Strip trailing non-digit/non-dot sigils (e.g. ' kr').
  s = s.replace(/[^\d.]+$/, '');

  s = s.replace(/,/g, '').replace(/\s+/g, '');

  if (s === '' || s === '.') return null;
  const n = Number(s);
  if (!Number.isFinite(n)) return null;
  return negative ? -n : n;
}

// ---------- bumpDecimals --------------------------------------------

/**
 * Compute the partial style update emitted by the Increase / Decrease
 * Decimal quick-access buttons. Delta is +1 or -1.
 *
 * If the cell has no numFmt (or is Text/General), the button flips the
 * cell into Number format with a sensible base decimal count, then
 * bumps, matching Excel's behaviour.
 */
export function bumpDecimals(s, delta) {
  const fmt = s?.numFmt || FORMAT_GENERAL;
  let workingFmt = fmt;
  if (fmt === FORMAT_GENERAL || fmt === FORMAT_TEXT) workingFmt = FORMAT_NUMBER;

  const baseDefault = workingFmt === FORMAT_PERCENTAGE ? 0 : 2;
  const base = effectiveDecimals(s, baseDefault);
  const next = Math.max(0, Math.min(30, base + delta));

  const out = { numFmt: workingFmt, decimals: next };
  if (workingFmt === FORMAT_NUMBER) {
    if (s?.useThousands === undefined) out.useThousands = true;
    if (s?.negativeStyle === undefined) out.negativeStyle = 'minus';
  }
  return out;
}

// ---------- Excel numFmt bridge -------------------------------------

/**
 * Produce { id?, code } describing the OOXML format for this style.
 *
 * `id` is only set when the format is one of Excel's built-in numFmtIds,
 * which the injector can reference by id without registering a <numFmt>
 * entry. `code` is always set so the injector can emit a <numFmt> entry
 * when a custom id is needed.
 */
export function toExcelNumFmtCode(s) {
  const fmt = s?.numFmt || FORMAT_GENERAL;
  if (fmt === FORMAT_GENERAL) return { id: 0, code: 'General' };
  if (fmt === FORMAT_TEXT)    return { id: 49, code: '@' };

  const decimals = effectiveDecimals(s, fmt === FORMAT_PERCENTAGE ? 0 : 2);
  const decFrag = decimals > 0 ? '.' + '0'.repeat(decimals) : '';

  switch (fmt) {
    case FORMAT_NUMBER: {
      const thou = s?.useThousands !== false;
      const body = (thou ? '#,##0' : '0') + decFrag;
      const neg = effectiveNegativeStyle(s);
      const code = applyNegativeCodeWrap(body, neg);
      return { code };
    }
    case FORMAT_CURRENCY: {
      const sym = quoteExcelSymbol(effectiveCurrency(s));
      const body = sym + '#,##0' + decFrag;
      const neg = effectiveNegativeStyle(s);
      const code = applyNegativeCodeWrap(body, neg);
      return { code };
    }
    case FORMAT_ACCOUNTING: {
      const sym = quoteExcelSymbol(effectiveCurrency(s));
      const digits = decimals > 0 ? '?'.repeat(decimals) : '';
      const pos = '_(' + sym + '* #,##0' + decFrag + '_)';
      const neg = '_(' + sym + '* (#,##0' + decFrag + ')';
      const zero = '_(' + sym + '* "-"' + digits + '_)';
      const textPart = '_(@_)';
      return { code: [pos, neg, zero, textPart].join(';') };
    }
    case FORMAT_SHORT_DATE:
      return { id: 14, code: 'm/d/yyyy' };
    case FORMAT_LONG_DATE:
      return { code: '[$-F800]dddd, mmmm dd, yyyy' };
    case FORMAT_TIME:
      return { id: 19, code: 'h:mm:ss AM/PM' };
    case FORMAT_PERCENTAGE: {
      if (decimals === 0) return { id: 9,  code: '0%' };
      if (decimals === 2) return { id: 10, code: '0.00%' };
      return { code: '0' + decFrag + '%' };
    }
    case FORMAT_FRACTION: {
      const d = s?.decimals;
      if (d === 2) return { id: 13, code: '# ??/??' };
      if (d === 3) return { code: '# ???/???' };
      return { id: 12, code: '# ?/?' };
    }
    case FORMAT_SCIENTIFIC: {
      if (decimals === 2) return { id: 11, code: '0.00E+00' };
      return { code: '0' + decFrag + 'E+00' };
    }
    default:
      return { id: 0, code: 'General' };
  }
}

function applyNegativeCodeWrap(body, neg) {
  switch (neg) {
    case 'parens':    return body + ';(' + body + ')';
    case 'red':       return body + ';[Red]-' + body;
    case 'redParens': return body + ';[Red](' + body + ')';
    case 'minus':
    default:          return body;
  }
}

// Wrap the currency symbol for Excel's format code language. Single-
// character well-known sigils ($, \u00a3, \u00a5, \u20ac) are safe in
// bare form but wrapping them in quotes is also valid and sidesteps
// edge cases like the '$' metacharacter in OOXML locale tags. Multi-
// character symbols ('NT$', 'CHF', ...) must be quoted.
function quoteExcelSymbol(sym) {
  if (!sym) return '';
  return '"' + String(sym).replace(/"/g, '""') + '"';
}

/**
 * Invert toExcelNumFmtCode. Returns the app-side style subset or null
 * when the code is not recognised (callers should fall back to 'general').
 */
export function fromExcelNumFmtCode(code, builtinId) {
  if (builtinId != null) {
    switch (builtinId) {
      case 0:  return { numFmt: FORMAT_GENERAL };
      case 1:  return { numFmt: FORMAT_NUMBER, decimals: 0, useThousands: false, negativeStyle: 'minus' };
      case 2:  return { numFmt: FORMAT_NUMBER, decimals: 2, useThousands: false, negativeStyle: 'minus' };
      case 3:  return { numFmt: FORMAT_NUMBER, decimals: 0, useThousands: true,  negativeStyle: 'minus' };
      case 4:  return { numFmt: FORMAT_NUMBER, decimals: 2, useThousands: true,  negativeStyle: 'minus' };
      case 9:  return { numFmt: FORMAT_PERCENTAGE, decimals: 0 };
      case 10: return { numFmt: FORMAT_PERCENTAGE, decimals: 2 };
      case 11: return { numFmt: FORMAT_SCIENTIFIC, decimals: 2 };
      case 12: return { numFmt: FORMAT_FRACTION, decimals: 1 };
      case 13: return { numFmt: FORMAT_FRACTION, decimals: 2 };
      case 14:
      case 15:
      case 16:
      case 17:
      case 22:
        return { numFmt: FORMAT_SHORT_DATE };
      case 18:
      case 19:
      case 20:
      case 21:
      case 45:
      case 46:
      case 47:
        return { numFmt: FORMAT_TIME };
      case 37: return { numFmt: FORMAT_NUMBER, decimals: 0, useThousands: true, negativeStyle: 'parens' };
      case 38: return { numFmt: FORMAT_NUMBER, decimals: 0, useThousands: true, negativeStyle: 'redParens' };
      case 39: return { numFmt: FORMAT_NUMBER, decimals: 2, useThousands: true, negativeStyle: 'parens' };
      case 40: return { numFmt: FORMAT_NUMBER, decimals: 2, useThousands: true, negativeStyle: 'redParens' };
      case 48: return { numFmt: FORMAT_SCIENTIFIC, decimals: 1 };
      case 49: return { numFmt: FORMAT_TEXT };
      default: break;
    }
  }

  if (!code) return null;
  const c = String(code).trim();
  if (c === '' || c.toLowerCase() === 'general') return { numFmt: FORMAT_GENERAL };
  if (c === '@') return { numFmt: FORMAT_TEXT };

  // Percentage.
  const pct = c.match(/^(#,##0|0)(\.0+)?%$/);
  if (pct) {
    const decimals = pct[2] ? pct[2].length - 1 : 0;
    return { numFmt: FORMAT_PERCENTAGE, decimals };
  }

  // Scientific.
  const sci = c.match(/^0(\.0+)?E[+-]0+$/i);
  if (sci) {
    const decimals = sci[1] ? sci[1].length - 1 : 0;
    return { numFmt: FORMAT_SCIENTIFIC, decimals };
  }

  // Fraction.
  const fracMatch = c.match(/^#\s+(\?+)\/(\?+)$/);
  if (fracMatch) {
    return { numFmt: FORMAT_FRACTION, decimals: Math.max(fracMatch[1].length, fracMatch[2].length) };
  }

  // Accounting: first section of the 4-section pattern starts with `_(`.
  const accSym = c.match(/^_\(\s*("(?:[^"]|"")*"|\[\$[^\]]*\]|[^*\s]+)\s*\*\s*#,##0(\.0+)?_\)/);
  if (accSym) {
    const decimals = accSym[2] ? accSym[2].length - 1 : 0;
    const symbol = parseExcelSymbol(accSym[1]);
    return { numFmt: FORMAT_ACCOUNTING, decimals, currency: symbol };
  }

  // Currency: leading symbol, then number body, optional negative section.
  const curMatch = c.match(/^("(?:[^"]|"")*"|\[\$[^\]]*\]|\$|\u00a3|\u00a5|\u20ac)(#,##0)(\.0+)?(?:;(.+))?$/);
  if (curMatch) {
    const symbol = parseExcelSymbol(curMatch[1]);
    const decimals = curMatch[3] ? curMatch[3].length - 1 : 0;
    const neg = curMatch[4] ? detectNegativeStyle(curMatch[4]) : 'minus';
    return { numFmt: FORMAT_CURRENCY, decimals, currency: symbol, negativeStyle: neg };
  }

  // Number with or without thousands separators.
  const numMatch = c.match(/^(#,##0|0)(\.0+)?(?:;(.+))?$/);
  if (numMatch) {
    const useThousands = numMatch[1] === '#,##0';
    const decimals = numMatch[2] ? numMatch[2].length - 1 : 0;
    const neg = numMatch[3] ? detectNegativeStyle(numMatch[3]) : 'minus';
    return { numFmt: FORMAT_NUMBER, decimals, useThousands, negativeStyle: neg };
  }

  // Dates: look for date-ish tokens.
  if (/dddd|mmmm/i.test(c)) return { numFmt: FORMAT_LONG_DATE };
  if (/d{1,4}|m{1,4}|y{2,4}/i.test(c)) return { numFmt: FORMAT_SHORT_DATE };
  if (/h{1,2}|s{1,2}|AM\/PM|A\/P/i.test(c)) return { numFmt: FORMAT_TIME };

  return null;
}

function parseExcelSymbol(raw) {
  if (!raw) return '$';
  if (raw.startsWith('"')) {
    return raw.slice(1, -1).replace(/""/g, '"');
  }
  if (raw.startsWith('[$')) {
    // Format: [$SYM-LOCALE] or [$SYM]. Strip the LOCALE suffix.
    const inner = raw.slice(2, -1);
    const dash = inner.lastIndexOf('-');
    return dash >= 0 ? inner.slice(0, dash) : inner;
  }
  return raw;
}

function detectNegativeStyle(negSection) {
  const hasRed = /\[Red\]/i.test(negSection);
  const hasParens = /\(/.test(negSection);
  if (hasRed && hasParens) return 'redParens';
  if (hasRed) return 'red';
  if (hasParens) return 'parens';
  return 'minus';
}

// Utility used by the toolbar menu to describe the currently-selected
// cell's format in a compact single-word label.
export function describeFormat(s) {
  const fmt = s?.numFmt || FORMAT_GENERAL;
  return FORMAT_LABELS[fmt] || 'General';
}

// Utility used by the toolbar to render a live preview of what the
// given raw value would look like under the given format spec.
export function previewFormat(value, partial) {
  try {
    const rendered = formatCellValue(value, partial);
    return rendered === '' ? '\u2014' : rendered;
  } catch {
    return '\u2014';
  }
}

// ─── utils.js ───
// 공통 헬퍼 함수 모음 (텍스트 정리 / 날짜 / 금액 / 전화번호 / DOM 렌더링)

// ── Ghost Text ──────────────────────────────────────────────────────────────
export const GHOST_PATTERN =
  /[\u0000-\u001F\u007F-\u009F\u00A0\u034F\u061C\u115F\u1160\u17B4\u17B5\u180E\u2000-\u200F\u2028\u2029\u202F\u205F\u2060-\u206F\u2800\u3000\u3164\uFE00-\uFE0F\uFEFF\uFFA0]/g;

export function cleanGhostText(value) {
  if (value === null || value === undefined) return '';
  return String(value).normalize('NFKC').replace(GHOST_PATTERN, '').replace(/\r?\n|\t/g, ' ');
}

export function isBlankLike(value) {
  return cleanGhostText(value).trim() === '';
}

// ── Date ────────────────────────────────────────────────────────────────────
export function parseFlexibleDate(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  let s = cleanGhostText(value).trim();
  if (!s) return null;

  // 엑셀 일련번호 (45200 같은 5자리 숫자)
  if (/^\d{5}$/.test(s)) {
    const serial = parseInt(s, 10);
    if (serial > 20000 && serial < 80000) {
      const d = new Date((serial - 25569) * 86400 * 1000);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // 2024년 6월 1일
  let m = s.match(/^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일?$/);
  if (m) return _makeDate(m[1], m[2], m[3]);

  // YYYYMMDD
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return _makeDate(m[1], m[2], m[3]);

  // YYYY-M-D / YYYY/M/D / YYYY.M.D
  m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (m) return _makeDate(m[1], m[2], m[3]);

  // M/D/YYYY or D/M/YYYY — 첫 숫자 > 12 이면 DD/MM/YYYY
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2], 10), y = parseInt(m[3], 10);
    return a > 12 ? _makeDate(y, b, a) : _makeDate(y, a, b);
  }

  // 특수문자 제거 후 8자리 숫자
  const digits = s.replace(/\D/g, '');
  if (/^\d{8}$/.test(digits))
    return _makeDate(digits.slice(0, 4), digits.slice(4, 6), digits.slice(6, 8));

  return null;
}

function _makeDate(y, m, d) {
  const yy = parseInt(y, 10), mm = parseInt(m, 10), dd = parseInt(d, 10);
  if (!yy || !mm || !dd) return null;
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
  const date = new Date(yy, mm - 1, dd);
  if (date.getFullYear() !== yy || date.getMonth() !== mm - 1 || date.getDate() !== dd) return null;
  return date;
}

export function formatDateValue(date, fmt) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  if (fmt === 'YYYY/MM/DD') return `${y}/${m}/${d}`;
  if (fmt === 'DD/MM/YYYY') return `${d}/${m}/${y}`;
  return `${y}-${m}-${d}`;
}

// ── Amount ───────────────────────────────────────────────────────────────────
export const AMOUNT_UNIT_LABELS = { 1: '원', 1000: '천원', 10000: '만원', 1000000: '백만원', 100000000: '억원' };

export function parseFlexibleNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && isFinite(value)) return value;

  let s = cleanGhostText(value).trim();
  if (!s) return null;

  let negative = false;
  if (/^\(.+\)$/.test(s)) { negative = true; s = s.slice(1, -1); }

  s = s.replace(/,/g, '').replace(/\s/g, '').replace(/[₩￦원]/g, '')
       .replace(/▲|△/g, '-').replace(/[^0-9.+\-]/g, '');

  if (!s || s === '-' || s === '.' || s === '+') return null;
  const n = Number(s);
  if (!isFinite(n)) return null;
  return negative ? -Math.abs(n) : n;
}

export function convertAmountUnit(value, fromUnit, toUnit, decimalsMode) {
  const parsed = parseFlexibleNumber(value);
  if (parsed === null) return { ok: false, value };

  const from = Number(fromUnit || 1), to = Number(toUnit || 1);
  let converted = (parsed * from) / to;

  if (decimalsMode !== 'auto') {
    const digits = parseInt(decimalsMode, 10);
    converted = Number(converted.toFixed(isNaN(digits) ? 0 : digits));
  } else if (Number.isInteger(converted)) {
    converted = Number(converted.toFixed(0));
  } else {
    converted = Number(converted.toFixed(6));
  }
  return { ok: true, value: converted };
}

// ── Phone ────────────────────────────────────────────────────────────────────
export function normalizePhoneNumber(value) {
  if (!value && value !== 0) return { formatted: '', type: '', valid: false };

  let original = cleanGhostText(value).trim();
  if (!original) return { formatted: '', type: '', valid: false };

  let v = original
    .replace(/\b(ext|extension)\b\.?\s*\d+.*$/i, '')
    .replace(/내선\s*\d+.*$/g, '')
    .replace(/[#].*$/g, '')
    .trim();

  v = v.replace(/^\+\s*82\s*[-.)]?\s*/, '0')
       .replace(/^0082\s*[-.)]?\s*/, '0')
       .replace(/^82\s*[-.)]?\s*/, '0');

  let digits = v.replace(/\D/g, '');
  if (/^10\d{8}$/.test(digits)) digits = '0' + digits;
  if (!digits) return { formatted: '', type: '', valid: false };

  let formatted = digits, type = '기타', valid = true;

  if (/^01[016789]/.test(digits)) {
    type = '휴대폰';
    if (digits.length === 11)      formatted = digits.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
    else if (digits.length === 10) formatted = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    else valid = false;
  } else if (/^02/.test(digits)) {
    type = '지역번호';
    if (digits.length === 9)       formatted = digits.replace(/(02)(\d{3})(\d{4})/, '$1-$2-$3');
    else if (digits.length === 10) formatted = digits.replace(/(02)(\d{4})(\d{4})/, '$1-$2-$3');
    else valid = false;
  } else if (/^1[568]\d{6}$/.test(digits)) {
    type = '대표번호';
    formatted = digits.replace(/(\d{4})(\d{4})/, '$1-$2');
  } else if (/^0(50|70)\d+/.test(digits)) {
    type = '특수번호';
    if (digits.length === 11)      formatted = digits.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
    else if (digits.length === 10) formatted = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    else valid = false;
  } else if (/^0[3-6]\d/.test(digits)) {
    type = '지역번호';
    if (digits.length === 10)      formatted = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    else if (digits.length === 11) formatted = digits.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
    else valid = false;
  } else {
    valid = false;
  }

  return { formatted, type, valid };
}

// ── DOM / HTML ───────────────────────────────────────────────────────────────
export function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

export function escapeJs(str) {
  return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

export function renderTable(id, cols, rows) {
  const $wrap = $(`#${id}`);
  if (!rows.length) {
    $wrap.html('<p style="padding:16px;color:var(--muted)">데이터가 없습니다.</p>');
    return;
  }
  const thead = '<thead><tr>' + cols.map(h => `<th>${escapeHtml(h)}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.slice(0, 50).map(row =>
    '<tr>' + cols.map(h => `<td>${escapeHtml(row[h] ?? '')}</td>`).join('') + '</tr>'
  ).join('') +
  (rows.length > 50 ? `<tr><td colspan="${cols.length}" style="text-align:center;color:var(--muted);padding:10px">... 외 ${rows.length - 50}개 행 (다운로드 시 전체 포함)</td></tr>` : '') +
  '</tbody>';
  $wrap.html(`<table>${thead}${tbody}</table>`);
}

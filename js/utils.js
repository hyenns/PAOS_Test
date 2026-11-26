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
export const AMOUNT_UNIT_LABELS = { 1: '원', 1000: '천원', 10000: '만원', 1000000: '백만원', 100000000: '억원', auto: '혼합 단위 자동(단위 없는 숫자는 원)' };

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

// 한국어 금액 표기는 조/억/만 단위가 섞여 있어도 원 단위로 해석합니다.
// 단위가 없는 숫자만 '현재 데이터 단위' 설정을 사용합니다.
function parseKoreanDigitGroup(group) {
  if (!group) return null;
  if (/^\d+(?:\.\d+)?$/.test(group)) return Number(group);

  let total = 0;
  let cursor = 0;
  let previousScale = 10000;
  const pattern = /(\d+(?:\.\d+)?)(천|백|십)/g;
  let match;

  while ((match = pattern.exec(group)) !== null) {
    if (match.index !== cursor) return null;
    const scale = { 천: 1000, 백: 100, 십: 10 }[match[2]];
    if (scale >= previousScale) return null;
    total += Number(match[1]) * scale;
    previousScale = scale;
    cursor = pattern.lastIndex;
  }

  if (cursor === 0) return null;
  if (cursor < group.length) {
    const remainder = group.slice(cursor);
    if (!/^\d+(?:\.\d+)?$/.test(remainder)) return null;
    total += Number(remainder);
  }
  return total;
}

function parseAmountToWon(value, fromUnit) {
  if (value === null || value === undefined || value === '') return null;

  // Excel 숫자 셀은 단위 표시가 없으므로 선택한 기준 단위를 적용합니다.
  if (typeof value === 'number') {
    const result = value * (fromUnit === 'auto' ? 1 : Number(fromUnit));
    return Number.isSafeInteger(result) || (Number.isFinite(result) && Math.abs(result) <= Number.MAX_SAFE_INTEGER)
      ? result : null;
  }

  let source = cleanGhostText(value).trim();
  if (!source) return null;

  let sign = 1;
  if (/^\(.+\)$/.test(source)) {
    sign = -1;
    source = source.slice(1, -1);
  } else if (/^[-▲△]/.test(source)) {
    sign = -1;
    source = source.replace(/^[-▲△]/, '');
  } else {
    source = source.replace(/^\+/, '');
  }

  const hasWonMark = /[원₩￦]/.test(source);
  // 천·백·십 표현은 만 단위 안쪽의 숫자에도 사용될 수 있습니다.
  const compact = source.replace(/[\s,]/g, '').replace(/[원₩￦]/g, '');
  if (!compact) return null;
  if (!/^[\d.조억만천백십]+$/.test(compact)) return null;

  const hasKoreanUnit = /[조억만천백십]/.test(compact);
  if (!hasKoreanUnit && !hasWonMark) {
    if (!/^\d+(?:\.\d+)?$/.test(compact)) return null;
    const fallback = fromUnit === 'auto' ? 1 : Number(fromUnit);
    const won = sign * Number(compact) * fallback;
    return Number.isFinite(won) && Math.abs(won) <= Number.MAX_SAFE_INTEGER ? won : null;
  }

  let total = 0;
  let cursor = 0;
  let previousMajorScale = Infinity;
  const majorPattern = /([^조억만]+)(조|억|만)/g;
  let match;

  while ((match = majorPattern.exec(compact)) !== null) {
    if (match.index !== cursor) return null;
    const scale = { 조: 1e12, 억: 1e8, 만: 1e4 }[match[2]];
    if (scale >= previousMajorScale) return null;
    const groupValue = parseKoreanDigitGroup(match[1]);
    if (groupValue === null) return null;
    total += groupValue * scale;
    previousMajorScale = scale;
    cursor = majorPattern.lastIndex;
  }

  const remainder = compact.slice(cursor);
  if (remainder) {
    const remainderValue = parseKoreanDigitGroup(remainder);
    if (remainderValue === null) return null;
    total += remainderValue;
  } else if (cursor === 0) {
    return null;
  }

  const won = total * sign;
  return Number.isFinite(won) && Math.abs(won) <= Number.MAX_SAFE_INTEGER ? won : null;
}

export function convertAmountUnit(value, fromUnit, toUnit, decimalsMode) {
  const won = parseAmountToWon(value, fromUnit);
  if (won === null) return { ok: false, value };

  const to = Number(toUnit || 1);
  if (!Number.isFinite(to) || to <= 0) return { ok: false, value };

  let converted = won / to;
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


// ── Korean Address ──────────────────────────────────────────────────────────
const KOREAN_ADDRESS_PREFIX_MAP = [
  ['서울특별시', '서울'], ['서울시', '서울'],
  ['부산광역시', '부산'],
  ['대구광역시', '대구'],
  ['인천광역시', '인천'],
  ['광주광역시', '광주'],
  ['대전광역시', '대전'],
  ['울산광역시', '울산'],
  ['세종특별자치시', '세종'], ['세종시', '세종'],
  ['경기도', '경기'],
  ['강원특별자치도', '강원'], ['강원도', '강원'],
  ['충청북도', '충북'], ['충청남도', '충남'],
  ['전북특별자치도', '전북'], ['전라북도', '전북'],
  ['전라남도', '전남'],
  ['경상북도', '경북'], ['경상남도', '경남'],
  ['제주특별자치도', '제주'], ['제주도', '제주'],
];

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function normalizeKoreanAddress(value) {
  if (value === null || value === undefined || value === '') {
    return { ok: false, changed: false, value };
  }

  const original = cleanGhostText(value).replace(/\s+/g, ' ').trim();
  if (!original) return { ok: false, changed: false, value: original };

  for (const [fullName, shortName] of KOREAN_ADDRESS_PREFIX_MAP) {
    const re = new RegExp(`^${escapeRegExp(fullName)}\\s*`);
    if (!re.test(original)) continue;

    const rest = original.replace(re, '').replace(/\s+/g, ' ').trim();
    const normalized = rest ? `${shortName} ${rest}` : shortName;
    return {
      ok: true,
      changed: normalized !== original,
      value: normalized,
      matched: fullName,
    };
  }

  return { ok: true, changed: false, value: original };
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

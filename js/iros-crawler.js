// ─── iros-crawler.js ───
// 인터넷등기소 크롤링 화면 제어 + Python 백엔드 호출

let workbook = null;
let selectedFile = null;
let resultExcelBase64 = null;
let resultFilename = '등기소_크롤링_결과.xlsx';

function getIrosSearchMode() {
  return $('#iros-search-mode').val() === 'registration' ? 'registration' : 'company';
}

function updateIrosSearchModeUI() {
  const mode = getIrosSearchMode();
  const isRegistration = mode === 'registration';

  $('#iros-panel-subtitle').text(isRegistration ? '인터넷등기소 등록번호검색' : '인터넷등기소 상호검색');
  $('#iros-panel-desc').text(
    isRegistration
      ? '법인등록번호를 기준으로 등기사항증명서 등록번호검색 결과의 법인정보를 수집합니다.'
      : '인터넷등기소 상호검색 결과 중 * 표시된 법인정보와 상태값을 수집합니다.'
  );
  $('#iros-search-mode-help').text(
    isRegistration
      ? '법인등록번호가 들어있는 열을 선택하면 열람·발급 > 법인 > 등록번호검색 경로로 조회합니다.'
      : '회사명이 들어있는 열을 선택하면 기존 상호검색 방식으로 조회합니다.'
  );
  $('#iros-upload-title').text(
    isRegistration ? '법인등록번호가 들어있는 엑셀 파일을 업로드하세요' : '회사명이 들어있는 엑셀 파일을 업로드하세요'
  );
  $('#iros-upload-desc').text(
    isRegistration
      ? '.xlsx · .xls 파일 가능 / 법인등록번호 13자리 또는 하이픈 포함 형식 지원'
      : '.xlsx · .xls 파일 가능 / 첫 행 제목 여부 선택 가능'
  );
  $('#iros-company-search-options').toggle(!isRegistration);
  $('#iros-clean-options').toggle(!isRegistration);

  if (workbook) {
    updatePreview();
  }
}

const IROS_RESULT_COLUMNS = [
  '검색값',
  '법인종류',
  '상호(명칭)',
  '법인등록번호',
  '본점소재지',
  '등기상태',
  '상호말소상태',
  '주말 여부',
];

function excelColumnLetter(index) {
  let num = index + 1;
  let letters = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    letters = String.fromCharCode(65 + mod) + letters;
    num = Math.floor((num - mod) / 26);
  }
  return letters;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function sheetRows(sheetName) {
  const ws = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

function sheetStartColumnIndex(sheetName) {
  const ws = workbook.Sheets[sheetName];
  const ref = ws && ws['!ref'];
  if (!ref) return 0;
  try {
    return XLSX.utils.decode_range(ref).s.c || 0;
  } catch (e) {
    return 0;
  }
}

function looksLikeRegistrationNumber(value) {
  const digits = String(value ?? '').replace(/\D/g, '');
  return digits.length === 13;
}

function renderTable(targetSelector, rowsOrObjects, limit = 10, preferredHeaders = null, startColumnIndex = 0) {
  const $target = $(targetSelector);
  if (!rowsOrObjects || rowsOrObjects.length === 0) {
    $target.html('<div style="color: var(--muted); font-size: 0.8rem; padding: 16px;">표시할 데이터가 없습니다.</div>');
    return;
  }

  let headers = [];
  let rows = [];

  if (Array.isArray(rowsOrObjects[0])) {
    rows = rowsOrObjects.slice(0, limit);
    const maxLen = rows.reduce((m, r) => Math.max(m, r.length), 0);
    headers = Array.from({ length: maxLen }, (_, i) => excelColumnLetter(startColumnIndex + i));
  } else {
    const objectKeys = Object.keys(rowsOrObjects[0]);
    const preferred = Array.isArray(preferredHeaders) ? preferredHeaders : [];
    const extraKeys = objectKeys.filter(key => !preferred.includes(key));
    headers = [...preferred.filter(key => objectKeys.includes(key)), ...extraKeys];
    rows = rowsOrObjects.slice(0, limit).map(obj => headers.map(h => obj[h] ?? ''));
  }

  const thead = `<thead><tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr></thead>`;
  const tbody = `<tbody>${rows.map(r => `<tr>${headers.map((_, i) => `<td>${escapeHtml(r[i] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody>`;
  $target.html(`<table>${thead}${tbody}</table>`);
}

function updateSheetOptions() {
  if (!workbook) return;
  const names = workbook.SheetNames || [];
  $('#iros-sheet-select').html(names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join(''));
  updateColumnOptions();
}

function updateColumnOptions() {
  if (!workbook) return;

  const sheetName = $('#iros-sheet-select').val();
  const headerMode = $('#iros-header-mode').val();
  const rows = sheetRows(sheetName);
  const firstRow = rows[0] || [];
  const maxLen = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const startCol = sheetStartColumnIndex(sheetName);
  const options = [];

  for (let i = 0; i < maxLen; i++) {
    const actualColumnIndex = startCol + i;
    const col = excelColumnLetter(actualColumnIndex);
    const header = String(firstRow[i] ?? '').trim();
    let label = `${col}열`;
    if (headerMode === 'header' && header) label += ` - ${header}`;
    options.push(`<option value="${actualColumnIndex}" data-row-index="${i}" data-column-letter="${escapeHtml(col)}">${escapeHtml(label)}</option>`);
  }

  $('#iros-column-select').html(options.join(''));
  updatePreview();
}

function updatePreview() {
  if (!workbook) return;

  const sheetName = $('#iros-sheet-select').val();
  const headerMode = $('#iros-header-mode').val();
  const $selectedOption = $('#iros-column-select option:selected');
  const rowIndex = Number($selectedOption.attr('data-row-index') || 0);
  const rows = sheetRows(sheetName);
  const startCol = sheetStartColumnIndex(sheetName);
  const isRegistration = getIrosSearchMode() === 'registration';
  const firstSelectedValue = rows.length ? String(rows[0][rowIndex] ?? '').trim() : '';
  const includeFirstRowAsRegistrationData =
    isRegistration && headerMode === 'header' && looksLikeRegistrationNumber(firstSelectedValue);
  const dataRows =
    headerMode === 'header' && !includeFirstRowAsRegistrationData ? rows.slice(1) : rows;

  const selectedValues = dataRows
    .map(r => String(r[rowIndex] ?? '').trim())
    .filter(v => v && v.toLowerCase() !== 'nan');

  const selectedLabel = $selectedOption.text();
  const modeLabel = isRegistration ? '법인등록번호' : '회사명';
  const autoHeaderNote = includeFirstRowAsRegistrationData
    ? ' <span style="color: var(--muted);">첫 셀이 13자리 법인등록번호라 첫 행도 데이터로 포함했습니다.</span>'
    : '';
  $('#iros-preview-info').html(`✅ <strong>${escapeHtml(selectedLabel)}</strong> 기준으로 총 <strong>${selectedValues.length}</strong>개 ${modeLabel} 항목을 불러왔습니다.${autoHeaderNote} 아래는 원본 미리보기입니다.`);
  renderTable('#iros-preview-table', rows, 10, null, startCol);
}

async function checkServer() {
  const $status = $('#iros-server-status');
  try {
    const res = await fetch('/api/health', { cache: 'no-store' });
    if (!res.ok) throw new Error('server not ok');
    const data = await res.json();
    $status.removeClass('warn').addClass('ok').text(`✅ 서버 연결됨: ${data.app || 'WorkLab'} / 팀 테스트 가능`);
  } catch (e) {
    $status.removeClass('ok').addClass('warn').html('⚠️ 현재 GitHub Pages 또는 정적 화면으로 접속한 상태일 수 있습니다. 등기소 크롤링은 <strong>run_worklab.bat</strong>으로 서버를 켠 뒤 <strong>http://실행PC_IP:8000</strong> 주소에서 실행해야 합니다.');
  }
}


function appendIrosLog(message) {
  const $list = $('#irosLogList');
  $list.append(`<li>${escapeHtml(message)}</li>`);

  const listEl = $list[0];
  if (listEl) {
    const scrollTarget = listEl.parentElement || listEl;
    scrollTarget.scrollTop = scrollTarget.scrollHeight;
  }
}

function renderIrosSummary({ title = '크롤링 결과', badge = '', badgeType = 'default', items = [] }) {
  const badgeHtml = badge
    ? `<span class="iros-summary-badge ${escapeHtml(badgeType)}">${escapeHtml(badge)}</span>`
    : '';

  const itemHtml = items.map(item => `
    <div class="iros-summary-item">
      <div class="iros-summary-label">${escapeHtml(item.label)}</div>
      <div class="iros-summary-value">${escapeHtml(item.value)}</div>
      ${item.desc ? `<div class="iros-summary-desc">${escapeHtml(item.desc)}</div>` : ''}
    </div>
  `).join('');

  return `
    <div class="iros-summary-card">
      <div class="iros-summary-head">
        <div>
          <div class="iros-summary-kicker">IROS RESULT</div>
          <div class="iros-summary-title">${escapeHtml(title)}</div>
        </div>
        ${badgeHtml}
      </div>
      <div class="iros-summary-grid">${itemHtml}</div>
    </div>
  `;
}

function updateIrosProgress(current, total, resultCount) {
  $('#irosStatsRow').html(renderIrosSummary({
    title: '검색 진행 중',
    badge: '진행 중',
    badgeType: 'progress',
    items: [
      { label: '검색 진행', value: `${current ?? 0} / ${total ?? 0}` },
      { label: '현재 수집 결과', value: `${resultCount ?? 0}건` },
      { label: '처리 상태', value: '실시간 검색 중' },
    ],
  }));
}

async function readIrosStream(response) {
  const reader = response.body.getReader();
  const decoder = new TextDecoder('utf-8');
  let buffer = '';
  let finalData = null;

  while (true) {
    const { value, done } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop() || '';

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;

      const event = JSON.parse(trimmed);

      if (event.type === 'log') {
        appendIrosLog(event.message || '');
      } else if (event.type === 'progress') {
        updateIrosProgress(event.current, event.total, event.result_count);
      } else if (event.type === 'complete') {
        finalData = event;
      } else if (event.type === 'error') {
        throw new Error(event.message || '크롤링 실행 중 오류가 발생했습니다.');
      }
    }
  }

  if (buffer.trim()) {
    const event = JSON.parse(buffer.trim());
    if (event.type === 'complete') finalData = event;
    if (event.type === 'error') throw new Error(event.message || '크롤링 실행 중 오류가 발생했습니다.');
  }

  return finalData;
}

function resetIros() {
  workbook = null;
  selectedFile = null;
  resultExcelBase64 = null;
  $('#irosFileInput').val('');
  $('#iros-settings').hide();
  $('#iros-result-box').hide();
  $('#iros-search-mode').val('company');
  $('#iros-include-closed-records').prop('checked', false);
  $('#iros-include-erased-names').prop('checked', false);
  $('#iros-clean-enabled').prop('checked', false);
  $('#iros-clean-split-name').prop('checked', true);
  $('#iros-clean-remove-reg-hyphen').prop('checked', true);
  $('#iros-clean-standardize-address').prop('checked', true);
  $('#iros-preview-table').empty();
  $('#irosResultTable').empty();
  $('#irosLogList').empty();
  $('#irosStatsRow').empty();
  updateIrosSearchModeUI();
}

async function handleFile(file) {
  selectedFile = file;
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data, { type: 'array' });
  $('#iros-settings').show();
  $('#iros-result-box').hide();
  updateSheetOptions();
}

async function runIrosCrawler() {
  if (!selectedFile || !workbook) {
    alert('먼저 엑셀 파일을 업로드해 주세요.');
    return;
  }

  const sheetName = $('#iros-sheet-select').val();
  const headerMode = $('#iros-header-mode').val();
  const columnIndex = $('#iros-column-select').val();
  const columnLetter = $('#iros-column-select option:selected').attr('data-column-letter') || '';
  const searchMode = getIrosSearchMode();
  const isRegistration = searchMode === 'registration';

  const formData = new FormData();
  formData.append('file', selectedFile);
  formData.append('sheet_name', sheetName);
  formData.append('header_mode', headerMode);
  formData.append('column_index', columnIndex);
  formData.append('column_letter', columnLetter);
  formData.append('search_mode', searchMode);
  formData.append('include_closed_records', searchMode === 'company' && $('#iros-include-closed-records').is(':checked') ? 'true' : 'false');
  formData.append('include_erased_names', searchMode === 'company' && $('#iros-include-erased-names').is(':checked') ? 'true' : 'false');
  formData.append('clean_iros_enabled', !isRegistration && $('#iros-clean-enabled').is(':checked') ? 'true' : 'false');
  formData.append('clean_iros_split_name', !isRegistration && $('#iros-clean-split-name').is(':checked') ? 'true' : 'false');
  formData.append('clean_iros_remove_reg_hyphen', !isRegistration && $('#iros-clean-remove-reg-hyphen').is(':checked') ? 'true' : 'false');
  formData.append('clean_iros_standardize_address', !isRegistration && $('#iros-clean-standardize-address').is(':checked') ? 'true' : 'false');
  formData.append('headless', 'true');

  resultExcelBase64 = null;
  $('#btn-run-iros').prop('disabled', true).text('⏳ 크롤링 중...');
  $('#iros-result-box').show();
  $('#irosResultTable').empty();
  $('#irosLogList').empty();
  $('#irosStatsRow').html(renderIrosSummary({
    title: '크롤링 준비 중',
    badge: '준비 중',
    badgeType: 'progress',
    items: [
      { label: '처리 상태', value: '등기소 접속 준비' },
      { label: '검색 방식', value: searchMode === 'registration' ? '등록번호검색' : '상호검색' },
    ],
  }));
  appendIrosLog(`인터넷등기소 ${searchMode === 'registration' ? '등록번호검색' : '상호검색'}을 시작합니다.`);

  try {
    const res = await fetch('/api/iros/run', { method: 'POST', body: formData });

    if (!res.ok) {
      let errorMessage = '크롤링 실행 중 오류가 발생했습니다.';
      try {
        const errorData = await res.json();
        errorMessage = errorData.message || errorMessage;
      } catch (_) {
        const text = await res.text();
        if (text) errorMessage = text.slice(0, 200);
      }
      throw new Error(errorMessage);
    }

    const data = await readIrosStream(res);
    if (!data || data.ok === false) {
      throw new Error(data?.message || '크롤링 결과를 받지 못했습니다.');
    }

    resultExcelBase64 = data.excel_base64;
    resultFilename = data.filename || resultFilename;

    const total = data.total_input ?? 0;
    const count = data.total_result ?? 0;
    const completedMode = data.search_mode === 'registration' ? '등록번호검색' : '상호검색';
    const statusText = data.status_has_value ? '상태값 수집됨' : '상태값 없음/미노출';
    const badgeText = data.search_mode === 'registration' ? completedMode : statusText;
    const summaryItems = [
      { label: '검색 대상', value: `${total}건` },
      { label: '수집 결과', value: `${count}건` },
      { label: '검색 방식', value: completedMode },
      { label: '정제 방식', value: data.clean_sheet_added ? '시트 추가' : (data.clean_applied ? '결과 시트에 적용' : '미적용') },
    ];
    if (data.search_mode !== 'registration') {
      summaryItems.splice(2, 0, { label: '상태값', value: statusText, desc: '등기상태 · 상호말소상태 · 주말 여부' });
    }
    $('#irosStatsRow').html(renderIrosSummary({
      title: '크롤링 완료',
      badge: badgeText,
      badgeType: data.search_mode === 'registration' ? 'success' : (data.status_has_value ? 'success' : 'warning'),
      items: summaryItems,
    }));
    renderTable('#irosResultTable', data.results || [], 200, data.columns || IROS_RESULT_COLUMNS);
  } catch (e) {
    $('#irosStatsRow').html(renderIrosSummary({
      title: '크롤링 실패',
      badge: '오류',
      badgeType: 'warning',
      items: [
        { label: '처리 상태', value: '실패' },
        { label: '확인 필요', value: '로그 확인' },
      ],
    }));
    appendIrosLog(e.message);
    alert(e.message);
  } finally {
    $('#btn-run-iros').prop('disabled', false).text('🚀 등기소 크롤링 시작');
  }
}

function downloadIrosResult() {
  if (!resultExcelBase64) {
    alert('다운로드할 결과가 없습니다. 먼저 크롤링을 실행해 주세요.');
    return;
  }

  const byteChars = atob(resultExcelBase64);
  const byteNumbers = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
  const blob = new Blob([new Uint8Array(byteNumbers)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = resultFilename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export function showIrosPanel() {
  $('#saramin-panel').hide();
  $('#iros-panel').show();
  document.getElementById('iros-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  checkServer();
}

export function initIrosCrawler() {
  $('#irosFileInput').on('change', async function () {
    const file = this.files?.[0];
    if (file) await handleFile(file);
  });

  $('#iros-search-mode').on('change', updateIrosSearchModeUI);
  $('#iros-sheet-select').on('change', updateColumnOptions);
  $('#iros-header-mode').on('change', updateColumnOptions);
  $('#iros-column-select').on('change', updatePreview);
  $('#btn-run-iros').on('click', runIrosCrawler);
  $('#btn-download-iros').on('click', downloadIrosResult);
  $('#btn-reset-iros, #btn-reset-iros2').on('click', resetIros);

  $('#irosUploadZone').on('dragover', function (e) {
    e.preventDefault();
    $(this).addClass('dragover');
  });
  $('#irosUploadZone').on('dragleave drop', function () {
    $(this).removeClass('dragover');
  });
  $('#irosUploadZone').on('drop', async function (e) {
    e.preventDefault();
    const file = e.originalEvent.dataTransfer.files?.[0];
    if (file) await handleFile(file);
  });

  updateIrosSearchModeUI();
}

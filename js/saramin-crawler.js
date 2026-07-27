// ─── saramin-crawler.js ───
// 사람인 기업정보 크롤링 화면 제어 + Python 백엔드 호출

let saraminWorkbook = null;
let saraminSelectedFile = null;
let saraminResultExcelBase64 = null;
let saraminResultFilename = '사람인_기업정보_크롤링_결과.xlsx';

const SARAMIN_INTRO_COLUMNS = [
  '검색값',
  '회사명',
  '설립일',
  '대표자명',
  '업종',
  '기업형태',
  '사원수',
  '매출액',
  '평균연봉',
  '홈페이지',
  '사업내용',
  '주소',
  '사람인URL',
];

const SARAMIN_FINANCE_COLUMNS = [
  '사람인_재무정보URL',
  '재무_기준연도',
  '재무_매출액',
  '재무_동종업계순위',
  '재무_영업이익',
  '재무_당기순이익',
  '재무_자본금',
  '신용등급',
  '매출액_2019',
  '매출액_2020',
  '매출액_2021',
  '매출액_2022',
  '매출액_2023',
  '매출액_2024',
  '매출액_2025',
  '영업이익_2019',
  '영업이익_2020',
  '영업이익_2021',
  '영업이익_2022',
  '영업이익_2023',
  '영업이익_2024',
  '영업이익_2025',
  '당기순이익_2019',
  '당기순이익_2020',
  '당기순이익_2021',
  '당기순이익_2022',
  '당기순이익_2023',
  '당기순이익_2024',
  '당기순이익_2025',
  '자본금_2019',
  '자본금_2020',
  '자본금_2021',
  '자본금_2022',
  '자본금_2023',
  '자본금_2024',
  '자본금_2025',
];

const SARAMIN_RESULT_COLUMNS = [...SARAMIN_INTRO_COLUMNS, ...SARAMIN_FINANCE_COLUMNS];

function getSaraminPreferredColumns(collectFinance) {
  return collectFinance ? SARAMIN_RESULT_COLUMNS : SARAMIN_INTRO_COLUMNS;
}


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
  const ws = saraminWorkbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

function updateSaraminCollectDesc() {
  const collectFinance = $('#saramin-collect-finance').val() === 'true';
  if (collectFinance) {
    $('#saramin-collect-desc').text('기업소개와 재무정보를 함께 수집합니다. 결과에는 재무정보 URL, 재무 기준연도, 연도별 매출액·영업이익·당기순이익·자본금, 신용등급 열이 포함됩니다.');
  } else {
    $('#saramin-collect-desc').text('기업소개만 빠르게 수집합니다. 결과에는 회사명, 설립일, 대표자명, 업종, 기업형태, 사원수, 매출액, 홈페이지, 사업내용, 주소, 사람인 URL만 저장하며 재무정보 관련 열은 생성하지 않습니다.');
  }
}

function renderTable(targetSelector, rowsOrObjects, limit = 10, preferredHeaders = null) {
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
    headers = Array.from({ length: maxLen }, (_, i) => excelColumnLetter(i));
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

function updateSaraminSheetOptions() {
  if (!saraminWorkbook) return;
  const names = saraminWorkbook.SheetNames || [];
  $('#saramin-sheet-select').html(names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join(''));
  updateSaraminColumnOptions();
}

function updateSaraminColumnOptions() {
  if (!saraminWorkbook) return;

  const sheetName = $('#saramin-sheet-select').val();
  const headerMode = $('#saramin-header-mode').val();
  const rows = sheetRows(sheetName);
  const firstRow = rows[0] || [];
  const maxLen = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const options = [];

  for (let i = 0; i < maxLen; i++) {
    const col = excelColumnLetter(i);
    const header = String(firstRow[i] ?? '').trim();
    let label = `${col}열`;
    if (headerMode === 'header' && header) label += ` - ${header}`;
    options.push(`<option value="${i}">${escapeHtml(label)}</option>`);
  }

  $('#saramin-column-select').html(options.join(''));
  updateSaraminPreview();
}

function updateSaraminPreview() {
  if (!saraminWorkbook) return;

  const sheetName = $('#saramin-sheet-select').val();
  const headerMode = $('#saramin-header-mode').val();
  const columnIndex = Number($('#saramin-column-select').val() || 0);
  const rows = sheetRows(sheetName);
  const dataRows = headerMode === 'header' ? rows.slice(1) : rows;
  const selectedValues = dataRows
    .map(r => String(r[columnIndex] ?? '').trim())
    .filter(v => v && v.toLowerCase() !== 'nan');

  const selectedLabel = $('#saramin-column-select option:selected').text();
  $('#saramin-preview-info').html(`✅ <strong>${escapeHtml(selectedLabel)}</strong> 기준으로 총 <strong>${selectedValues.length}</strong>개 항목을 불러왔습니다. 아래는 원본 미리보기입니다.`);
  renderTable('#saramin-preview-table', rows, 10);
}

async function checkSaraminServer() {
  const $status = $('#saramin-server-status');
  try {
    const res = await fetch('/api/health', { cache: 'no-store' });
    if (!res.ok) throw new Error('server not ok');
    const data = await res.json();
    $status.removeClass('warn').addClass('ok').text(`✅ 서버 연결됨: ${data.app || 'WorkLab'} / 사람인 크롤링 가능`);
  } catch (e) {
    $status.removeClass('ok').addClass('warn').html('⚠️ 현재 GitHub Pages 또는 정적 화면으로 접속한 상태일 수 있습니다. 사람인 크롤링은 <strong>WorkLab 실행.bat</strong>으로 서버를 켠 뒤 실행해야 합니다.');
  }
}

function appendSaraminLog(message) {
  const $list = $('#saraminLogList');
  $list.append(`<li>${escapeHtml(message)}</li>`);

  const listEl = $list[0];
  if (listEl) {
    const scrollTarget = listEl.parentElement || listEl;
    scrollTarget.scrollTop = scrollTarget.scrollHeight;
  }
}

function renderSaraminSummary({ title = '크롤링 결과', badge = '', badgeType = 'default', items = [] }) {
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
          <div class="iros-summary-kicker">SARAMIN RESULT</div>
          <div class="iros-summary-title">${escapeHtml(title)}</div>
        </div>
        ${badgeHtml}
      </div>
      <div class="iros-summary-grid">${itemHtml}</div>
    </div>
  `;
}

function updateSaraminProgress(current, total, resultCount) {
  $('#saraminStatsRow').html(renderSaraminSummary({
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

async function readSaraminStream(response) {
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
        appendSaraminLog(event.message || '');
      } else if (event.type === 'progress') {
        updateSaraminProgress(event.current, event.total, event.result_count);
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

function resetSaramin() {
  saraminWorkbook = null;
  saraminSelectedFile = null;
  saraminResultExcelBase64 = null;
  $('#saraminFileInput').val('');
  $('#saramin-settings').hide();
  $('#saramin-result-box').hide();
  $('#saramin-preview-table').empty();
  $('#saraminResultTable').empty();
  $('#saraminLogList').empty();
  $('#saraminStatsRow').empty();
}

async function handleSaraminFile(file) {
  saraminSelectedFile = file;
  const data = await file.arrayBuffer();
  saraminWorkbook = XLSX.read(data, { type: 'array' });
  $('#saramin-settings').show();
  $('#saramin-result-box').hide();
  updateSaraminSheetOptions();
}

async function runSaraminCrawler() {
  if (!saraminSelectedFile || !saraminWorkbook) {
    alert('먼저 엑셀 파일을 업로드해 주세요.');
    return;
  }

  const sheetName = $('#saramin-sheet-select').val();
  const headerMode = $('#saramin-header-mode').val();
  const columnIndex = $('#saramin-column-select').val();
  const maxResultsPerKeyword = $('#saramin-max-results').val() || '1';
  const collectFinance = $('#saramin-collect-finance').val() === 'true';

  const formData = new FormData();
  formData.append('file', saraminSelectedFile);
  formData.append('sheet_name', sheetName);
  formData.append('header_mode', headerMode);
  formData.append('column_index', columnIndex);
  formData.append('max_results_per_keyword', maxResultsPerKeyword);
  formData.append('collect_finance', collectFinance ? 'true' : 'false');
  // 기본 실행은 Chrome 창 없이 백그라운드에서 진행합니다.
  formData.append('headless', 'true');

  saraminResultExcelBase64 = null;
  $('#btn-run-saramin').prop('disabled', true).text('⏳ 크롤링 중...');
  $('#saramin-result-box').show();
  $('#saraminResultTable').empty();
  $('#saraminLogList').empty();
  $('#saraminStatsRow').html(renderSaraminSummary({
    title: '크롤링 준비 중',
    badge: '준비 중',
    badgeType: 'progress',
    items: [
      { label: '처리 상태', value: '사람인 접속 준비' },
      { label: '검색 방식', value: '엑셀 기준 열 검색' },
      { label: '수집 범위', value: collectFinance ? '기업소개 + 재무정보' : '기업소개만' },
    ],
  }));
  appendSaraminLog('사람인 기업정보 검색을 시작합니다.');

  try {
    const res = await fetch('/api/saramin/run', { method: 'POST', body: formData });

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

    const data = await readSaraminStream(res);
    if (!data || data.ok === false) {
      throw new Error(data?.message || '크롤링 결과를 받지 못했습니다.');
    }

    saraminResultExcelBase64 = data.excel_base64;
    saraminResultFilename = data.filename || saraminResultFilename;

    const total = data.total_input ?? 0;
    const count = data.total_result ?? 0;
    $('#saraminStatsRow').html(renderSaraminSummary({
      title: '크롤링 완료',
      badge: count > 0 ? '수집 완료' : '결과 없음',
      badgeType: count > 0 ? 'success' : 'warning',
      items: [
        { label: '검색 대상', value: `${total}건` },
        { label: '수집 결과', value: `${count}건` },
        { label: '수집 항목', value: data.collect_finance ? '기업소개 · 재무정보' : '기업소개만', desc: data.collect_finance ? '상위 기업 상세정보 및 연도별 재무값' : '재무정보 관련 열 없이 기업소개 기본 열만 저장' },
      ],
    }));
    renderTable('#saraminResultTable', data.results || [], 200, data.columns || getSaraminPreferredColumns(data.collect_finance));
  } catch (e) {
    $('#saraminStatsRow').html(renderSaraminSummary({
      title: '크롤링 실패',
      badge: '오류',
      badgeType: 'warning',
      items: [
        { label: '처리 상태', value: '실패' },
        { label: '확인 필요', value: '로그 확인' },
      ],
    }));
    appendSaraminLog(e.message);
    alert(e.message);
  } finally {
    $('#btn-run-saramin').prop('disabled', false).text('🚀 사람인 크롤링 시작');
  }
}

function downloadSaraminResult() {
  if (!saraminResultExcelBase64) {
    alert('다운로드할 결과가 없습니다. 먼저 크롤링을 실행해 주세요.');
    return;
  }

  const byteChars = atob(saraminResultExcelBase64);
  const byteNumbers = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
  const blob = new Blob([new Uint8Array(byteNumbers)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = saraminResultFilename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export function showSaraminPanel() {
  $('#iros-panel').hide();
  $('#saramin-panel').show();
  document.getElementById('saramin-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  checkSaraminServer();
}

export function initSaraminCrawler() {
  $('#saraminFileInput').on('change', async function () {
    const file = this.files?.[0];
    if (file) await handleSaraminFile(file);
  });

  $('#saramin-sheet-select').on('change', updateSaraminColumnOptions);
  $('#saramin-header-mode').on('change', updateSaraminColumnOptions);
  $('#saramin-column-select').on('change', updateSaraminPreview);
  $('#saramin-collect-finance').on('change', updateSaraminCollectDesc);
  updateSaraminCollectDesc();
  $('#btn-run-saramin').on('click', runSaraminCrawler);
  $('#btn-download-saramin').on('click', downloadSaraminResult);
  $('#btn-reset-saramin, #btn-reset-saramin2').on('click', resetSaramin);

  $('#saraminUploadZone').on('dragover', function (e) {
    e.preventDefault();
    $(this).addClass('dragover');
  });
  $('#saraminUploadZone').on('dragleave drop', function () {
    $(this).removeClass('dragover');
  });
  $('#saraminUploadZone').on('drop', async function (e) {
    e.preventDefault();
    const file = e.originalEvent.dataTransfer.files?.[0];
    if (file) await handleSaraminFile(file);
  });
}

// ─── merge.js ───
// 파일 병합 기능 (업로드 / 병합 실행 / 다운로드 / 초기화)

import { cleanGhostText, isBlankLike, escapeHtml, renderTable } from './utils.js';

// ── 상태 ─────────────────────────────────────────────────────────────────────
let mergeFilesData     = [];
let mergeResultData    = null;
let mergeResultHeaders = [];

// ── 초기화 ────────────────────────────────────────────────────────────────────
export function initMerge() {
  const $zone  = $('#mergeUploadZone');
  const $input = $('#mergeFileInput');

  $zone.on('click', () => $input.trigger('click'));

  $zone.on('dragover', e => {
    e.preventDefault();
    $zone.addClass('dragover');
  });

  $zone.on('dragleave', () => $zone.removeClass('dragover'));

  $zone.on('drop', e => {
    e.preventDefault();
    $zone.removeClass('dragover');
    const files = e.originalEvent.dataTransfer.files;
    if (files && files.length) loadMergeFiles([...files]);
  });

  $input.on('change', function () {
    if (this.files && this.files.length) {
      loadMergeFiles([...this.files]);
      // 같은 파일을 다시 선택할 수 있도록 input 값을 비웁니다.
      this.value = '';
    }
  });
}

// ── 파일 로드 ─────────────────────────────────────────────────────────────────
function loadMergeFiles(files) {
  const excelFiles = files.filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
  if (!excelFiles.length) {
    alert('엑셀/CSV 파일을 선택해주세요.');
    return;
  }

  // 이미 올린 파일은 유지하고 새로 선택한 파일만 뒤에 추가합니다.
  // 같은 이름/크기/수정시간의 파일을 다시 선택한 경우에는 중복 추가하지 않습니다.
  const existingKeys = new Set(mergeFilesData.map(f => f.fileKey).filter(Boolean));
  const newFiles = excelFiles.filter(file => {
    const key = getMergeFileKey(file);
    return !existingKeys.has(key);
  });

  if (!newFiles.length) {
    alert('이미 추가된 파일입니다. 다른 파일을 선택해주세요.');
    return;
  }

  // 파일 구성이 바뀌면 이전 병합 결과는 무효화합니다.
  mergeResultData    = null;
  mergeResultHeaders = [];
  $('#merge-result-box').hide();

  Promise.all(newFiles.map(readMergeFile))
    .then(results => {
      const validResults = results.filter(Boolean);
      if (!validResults.length) {
        alert('정상적으로 읽힌 파일이 없습니다.');
        return;
      }

      mergeFilesData.push(...validResults);

      const $zone = $('#mergeUploadZone');
      $zone.css({ borderColor: 'var(--accent)', background: 'rgba(0,229,160,0.04)' });
      $zone.find('h3').text(`✅ ${mergeFilesData.length}개 파일 선택됨`);

      renderMergeFileList();
      $('#merge-settings').show();

      // 첫 파일만 올린 상태에서는 병합 기준 열을 만들지 않고 추가 업로드를 안내합니다.
      if (mergeFilesData.length < 2) {
        $('#merge-key-select').empty();
        $zone.find('p').text('파일을 1개 이상 더 추가하려면 다시 클릭하거나 드래그해주세요.');
        return;
      }

      const common = getCommonHeaders(mergeFilesData.map(f => f.headers));
      if (!common.length) {
        $('#merge-key-select').empty();
        $zone.find('p').text('공통 열이 없습니다. 병합할 파일의 열 이름을 확인해주세요.');
        alert('현재 추가된 파일 전체에 공통으로 존재하는 열이 없습니다. 열 이름을 확인해주세요.');
        return;
      }

      $zone.find('p').text(`공통 열 ${common.length}개 확인됨 · 파일을 더 추가하려면 다시 클릭하세요.`);
      populateMergeKeySelect(common);
    })
    .catch(err => {
      console.error(err);
      alert('파일 읽기 오류: ' + err.message);
    });
}

function readMergeFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb   = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

        if (!json || json.length < 2) { resolve(null); return; }

        let fileHeaders = (json[0] || []).map((h, i) => {
          const cleaned = cleanGhostText(h).trim();
          return cleaned || `열${i + 1}`;
        });

        // 중복 헤더 처리
        const seen = {};
        fileHeaders = fileHeaders.map(h => {
          if (seen[h] === undefined) { seen[h] = 0; return h; }
          seen[h]++;
          return `${h}_${seen[h] + 1}`;
        });

        const rows = json.slice(1)
          .filter(row => row.some(cell => !isBlankLike(cell)))
          .map(row => {
            const obj = {};
            fileHeaders.forEach((h, i) => {
              const val = row[i];
              obj[h] = val !== undefined && val !== null ? cleanGhostText(val).trim() : '';
            });
            return obj;
          });

        resolve({
          name: file.name,
          baseName: getBaseFileName(file.name),
          fileKey: getMergeFileKey(file),
          headers: fileHeaders,
          rows,
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error(`${file.name} 파일 읽기 실패`));
    reader.readAsArrayBuffer(file);
  });
}

// ── UI 렌더링 ─────────────────────────────────────────────────────────────────
function renderMergeFileList() {
  $('#mergeFileList').html(
    mergeFilesData.map((f, idx) => `
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;
                  background:var(--surface);border:1px solid var(--border);border-radius:10px;
                  padding:10px 12px;font-size:0.8rem;">
        <span><strong>${idx + 1}. ${escapeHtml(f.name)}</strong></span>
        <span style="color:var(--muted);flex-shrink:0">${f.rows.length}행 · ${f.headers.length}열</span>
      </div>
    `).join('')
  );
}

function populateMergeKeySelect(commonHeaders) {
  const $sel = $('#merge-key-select').empty();
  commonHeaders.forEach(h => $sel.append($('<option>').val(h).text(h)));
}

// ── 병합 실행 ─────────────────────────────────────────────────────────────────
export function runMerge() {
  if (mergeFilesData.length < 2) {
    alert('파일을 2개 이상 업로드해주세요.');
    return;
  }

  const keyCol    = $('#merge-key-select').val();
  const mergeType = $('#merge-type-select').val() || 'outer';

  if (!keyCol) { alert('병합 기준 열을 선택해주세요.'); return; }

  const allKeySets = mergeFilesData.map(f =>
    new Set(f.rows.map(r => cleanGhostText(r[keyCol] ?? '').trim()).filter(Boolean))
  );

  let finalKeys = [];
  if (mergeType === 'inner') {
    finalKeys = [...allKeySets[0]].filter(k => allKeySets.every(set => set.has(k)));
  } else {
    const union = new Set();
    allKeySets.forEach(set => set.forEach(k => union.add(k)));
    finalKeys = [...union];
  }

  const outputMap = new Map();
  finalKeys.forEach(k => outputMap.set(k, { [keyCol]: k }));

  mergeResultHeaders = [keyCol];
  const logs = [];

  mergeFilesData.forEach(fileObj => {
    const suffix        = '__' + fileObj.baseName;
    const duplicateCount = {};
    const firstByKey    = new Map();

    fileObj.rows.forEach(row => {
      const k = cleanGhostText(row[keyCol] ?? '').trim();
      if (!k) return;
      duplicateCount[k] = (duplicateCount[k] || 0) + 1;
      if (!firstByKey.has(k)) firstByKey.set(k, row);
    });

    const dupCount = Object.values(duplicateCount).filter(v => v > 1).length;
    if (dupCount) logs.push(`${fileObj.name}: 기준값 중복 ${dupCount}개는 첫 번째 행만 병합`);

    fileObj.headers.forEach(h => {
      if (h === keyCol) return;
      mergeResultHeaders.push(h + suffix);
    });

    finalKeys.forEach(k => {
      const out       = outputMap.get(k);
      const sourceRow = firstByKey.get(k);
      fileObj.headers.forEach(h => {
        if (h === keyCol) return;
        out[h + suffix] = sourceRow ? (sourceRow[h] ?? '') : '';
      });
    });
  });

  mergeResultData = finalKeys.map(k => outputMap.get(k));

  const sourceTotalRows = mergeFilesData.reduce((sum, f) => sum + f.rows.length, 0);
  $('#mergeStatsRow').html(
    `<div class="stat-pill">📄 파일 ${mergeFilesData.length}개</div>` +
    `<div class="stat-pill">📌 기준 열 ${escapeHtml(keyCol)}</div>` +
    `<div class="stat-pill green">✅ 병합 결과 ${mergeResultData.length}행</div>` +
    `<div class="stat-pill">원본 총 ${sourceTotalRows}행</div>`
  );

  const logMsg = logs.length
    ? logs
    : [`병합 완료: ${mergeType === 'inner' ? '모든 파일에 공통으로 있는 기준값만 포함' : '전체 기준값 포함'}`];
  $('#mergeLogList').html(logMsg.map(l => `<li>${escapeHtml(l)}</li>`).join(''));

  renderTable('mergeResultTable', mergeResultHeaders, mergeResultData);
  $('#merge-result-box').show();
}

// ── 다운로드 ──────────────────────────────────────────────────────────────────
export function downloadMergeResult() {
  if (!mergeResultData || !mergeResultHeaders.length) return;
  const ws = XLSX.utils.json_to_sheet(mergeResultData, { header: mergeResultHeaders });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Merged');
  XLSX.writeFile(wb, `병합결과_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ── 초기화 ────────────────────────────────────────────────────────────────────
export function resetMerge() {
  mergeFilesData     = [];
  mergeResultData    = null;
  mergeResultHeaders = [];

  const $input = $('#mergeFileInput');
  const $zone  = $('#mergeUploadZone');

  $input.val('');
  $zone.css({ borderColor: '', background: '' });
  $zone.find('h3').text('병합할 파일들을 여기에 드래그하거나 클릭하여 업로드');
  $zone.find('p').text('.xlsx · .xls · .csv 여러 개 선택 가능');

  $('#merge-settings').hide();
  $('#merge-result-box').hide();
  $('#mergeFileList').empty();
}

// ── 내부 헬퍼 ─────────────────────────────────────────────────────────────────

function getMergeFileKey(file) {
  return `${file.name}::${file.size}::${file.lastModified}`;
}

function getBaseFileName(name) {
  return String(name).replace(/\.[^.]+$/, '').replace(/[^가-힣a-zA-Z0-9_\-]+/g, '_');
}

function getCommonHeaders(headerLists) {
  if (!headerLists.length) return [];
  return headerLists.reduce((acc, headers) => acc.filter(h => headers.includes(h)));
}

// ─── excel-cleaner.js ───
// 엑셀 정제 기능 전체 (Step 1~3, 파일 업로드, 정제 실행, 결과 미리보기, 다운로드, 초기화)

import { cleanGhostText, isBlankLike, parseFlexibleDate, formatDateValue, parseFlexibleNumber, convertAmountUnit, AMOUNT_UNIT_LABELS, normalizePhoneNumber, escapeHtml, escapeJs, renderTable } from "./utils.js";

// ── 상태 ─────────────────────────────────────────────────────────────────────
let rawData = [];
let headers = [];
let originalHeaders = [];
let fileName = "";
let originalFileArrayBuffer = null;
let originalExcelJsWorkbook = null;
let originalExcelJsWorksheet = null;
let resultPreviewMode = "result";
let latestChangeInfo = null;
let caseRules = [];
let currentStep = 1;

const RAW_ROW_KEY = "__paosOriginalRowNumber";

// ── 초기화 (DOMContentLoaded 이후 호출) ───────────────────────────────────────
export function initExcelCleaner() {
  // 업로드 존
  const $zone = $("#uploadZone");
  const $input = $("#fileInput");

  $zone.on("click", () => $input.trigger("click"));
  $zone.on("dragover", (e) => {
    e.preventDefault();
    $zone.addClass("dragover");
  });
  $zone.on("dragleave", () => $zone.removeClass("dragover"));
  $zone.on("drop", (e) => {
    e.preventDefault();
    $zone.removeClass("dragover");
    const f = e.originalEvent.dataTransfer.files[0];
    if (f) loadFile(f);
  });
  $input.on("change", function () {
    if (this.files[0]) loadFile(this.files[0]);
  });

  // opt-sub 클릭 버블 차단 (카드 토글 방지)
  $(".opt-sub").on("click mousedown mouseup touchstart", (e) => e.stopPropagation());

  // 중복 제거 모드 변경
  $(document).on("change", 'input[name="dup-mode"]', function () {
    $("#sel-dup-col").toggle(this.value === "col");
    if (rawData.length > 0) showColSettings();
  });

  // 대소문자 변환 체크
  $(document).on("change", "#txt-case", function () {
    if (rawData.length > 0) showColSettings();
  });
}

// ── STEP 네비게이션 ───────────────────────────────────────────────────────────
export function goToStep(n) {
  if (n === 2) {
    const anyActive = $(".opt-card.active").length > 0;
    if (!anyActive) {
      $("#step1-warn").show();
      return;
    }
    $("#step1-warn").hide();
    updateSelOptsPreview();
    if (rawData.length > 0) showColSettings();
  }
  currentStep = n;
  $(".step-panel").removeClass("active");
  $(`#step-${n}`).addClass("active");
  _updateStepIndicator(n);
  $(".main").scrollTop(0);
}

function _updateStepIndicator(step) {
  for (let i = 1; i <= 3; i++) {
    $(`#si-${i}`)
      .removeClass("active done")
      .addClass(i < step ? "done" : i === step ? "active" : "");
  }
  for (let i = 1; i <= 2; i++) {
    $(`#sl-${i}`).toggleClass("done", i < step);
  }
}

function updateSelOptsPreview() {
  const labels = {
    dup: "🔁 중복 제거",
    empty: "🕳️ 빈 셀 처리",
    text: "✏️ 텍스트 정리",
    date: "📅 날짜 형식 통일",
    amount: "💰 금액 단위 변환",
    phone: "📞 전화번호 정제",
    filter: "🔍 데이터 필터링",
  };
  const active = ["dup", "empty", "text", "date", "amount", "phone", "filter"].filter((k) => $(`#card-${k}`).hasClass("active"));
  $("#sel-opts-preview").html(active.map((k) => `<div class="sopt-chip">${labels[k]}</div>`).join(""));
}

// ── 옵션 토글 ─────────────────────────────────────────────────────────────────
export function toggleOpt(key) {
  const $card = $(`#card-${key}`);
  const $chk = $(`#chk-${key}`);
  const $sub = $(`#sub-${key}`);
  const active = $card.toggleClass("active").hasClass("active");
  $chk.css("color", active ? "#0e0f14" : "transparent");
  $sub.toggleClass("show", active);
}

// ── 파일 로드 ─────────────────────────────────────────────────────────────────
async function loadFile(file) {
  fileName = file.name;
  originalFileArrayBuffer = null;
  originalExcelJsWorkbook = null;
  originalExcelJsWorksheet = null;

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      originalFileArrayBuffer = e.target.result.slice(0);
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array", cellDates: true, raw: false, cellStyles: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });

      if (/\.xlsx$/i.test(file.name) && window.ExcelJS) {
        originalExcelJsWorkbook = new ExcelJS.Workbook();
        await originalExcelJsWorkbook.xlsx.load(originalFileArrayBuffer.slice(0));
        originalExcelJsWorksheet = originalExcelJsWorkbook.worksheets[0] || null;
      }

      if (!json || json.length < 2) {
        alert("데이터가 없거나 헤더만 있습니다.");
        return;
      }

      // 헤더 처리
      headers = (json[0] || []).map((h, i) => {
        const cleaned = cleanGhostText(h).trim();
        return cleaned || `열${i + 1}`;
      });
      const seen = {};
      headers = headers.map((h) => {
        if (seen[h] === undefined) {
          seen[h] = 0;
          return h;
        }
        seen[h]++;
        return `${h}_${seen[h] + 1}`;
      });
      originalHeaders = headers.slice();

      rawData = json
        .slice(1)
        .map((row, rowIdx) => {
          const obj = {};
          headers.forEach((h, i) => {
            const val = row[i];
            obj[h] = val !== undefined && val !== null ? cleanGhostText(val) : "";
          });
          obj[RAW_ROW_KEY] = rowIdx + 2;
          return obj;
        })
        .filter((row) => headers.some((h) => !isBlankLike(row[h])));

      if (rawData.length === 0) {
        alert("데이터 행이 없습니다.");
        return;
      }

      const $zone = $("#uploadZone");
      $zone.css({ borderColor: "var(--accent)", background: "rgba(0,229,160,0.04)" });
      $zone.find("h3").text(`✅ ${file.name}`);
      $zone.find("p").text(`${rawData.length}행 · ${headers.length}열 로드 완료`);

      populateCols();
      showColSettings();
      renderExcelPreview("previewTable", 10);
      $("#previewBox").show();
      $("#step2-btns").show();
    } catch (err) {
      alert("파일 읽기 오류: " + err.message);
      console.error(err);
    }
  };
  reader.onerror = () => alert("파일 읽기 실패. 다시 시도해주세요.");
  reader.readAsArrayBuffer(file);
}

// ── 열 선택 UI 채우기 ─────────────────────────────────────────────────────────
function populateCols() {
  // select 공통
  ["#filter-col-select", "#case-col-select"].forEach((sel) => {
    $(sel).empty();
    headers.forEach((h) => $(sel).append($("<option>").val(h).text(h)));
  });

  // 체크박스 공통 빌더
  const buildChecks = (wrapId, cls, autoCheck) => {
    $(`#${wrapId}`).empty();
    headers.forEach((h) => {
      const checked = autoCheck && autoCheck(h) ? "checked" : "";
      $(`#${wrapId}`).append(`
        <label style="display:flex;align-items:center;gap:8px;font-size:0.8rem;color:var(--text);cursor:pointer;">
          <input type="checkbox" class="${cls}" value="${escapeHtml(h)}" ${checked}
                 style="accent-color:var(--accent);width:14px;height:14px;">
          <span>${escapeHtml(h)}</span>
        </label>
      `);
    });
  };

  buildChecks("date-col-checks", "date-col-check", null);
  buildChecks("amount-col-checks", "amount-col-check", (h) => /금액|매출|매출액|비용|단가|가격|원|amount|sales|revenue|price|cost/i.test(h));
  buildChecks("dup-col-checks", "dup-col-check", null);
  buildChecks("phone-col-checks", "phone-col-check", null);
}

// ── 열 설정 표시/숨김 ─────────────────────────────────────────────────────────
function showColSettings() {
  const hasDup = $("#card-dup").hasClass("active");
  const dupCol = $('input[name="dup-mode"]:checked').val() === "col";
  const hasDate = $("#card-date").hasClass("active");
  const hasAmount = $("#card-amount").hasClass("active");
  const hasFilter = $("#card-filter").hasClass("active");
  const hasPhone = $("#card-phone").hasClass("active");
  const hasCase = $("#card-text").hasClass("active") && $("#txt-case").is(":checked");

  const needCols = (hasDup && dupCol) || hasDate || hasAmount || hasFilter || hasPhone || hasCase;
  $("#col-settings").css("display", needCols ? "grid" : "none");
  $("#col-set-dup").toggle(hasDup && dupCol);
  $("#col-set-date").toggle(hasDate);
  $("#col-set-amount").toggle(hasAmount);
  $("#col-set-phone").toggle(hasPhone);
  $("#col-set-filter").toggle(hasFilter);
  $("#col-set-case").toggle(hasCase);

  if (hasCase) renderCaseRules();
}

// ── 대소문자 변환 규칙 ────────────────────────────────────────────────────────
export function addCaseRule() {
  const col = $("#case-col-select").val();
  const mode = $("#case-mode-select").val();
  if (!col || !mode) return;
  const existing = caseRules.find((r) => r.col === col);
  if (existing) existing.mode = mode;
  else caseRules.push({ col, mode });
  renderCaseRules();
}

export function removeCaseRule(col) {
  caseRules = caseRules.filter((r) => r.col !== col);
  renderCaseRules();
}

function renderCaseRules() {
  const labelMap = { upper: "모두 대문자", lower: "모두 소문자", first: "첫 글자만 대문자" };
  if (!caseRules.length) {
    $("#case-rule-list").html('<div style="font-size:0.75rem;color:var(--muted);padding:6px 2px">아직 추가된 변환 규칙이 없습니다.</div>');
    return;
  }
  $("#case-rule-list").html(
    caseRules
      .map(
        (r) => `
    <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;
                background:var(--surface);border:1px solid var(--border);border-radius:8px;
                padding:9px 12px;font-size:0.8rem;">
      <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">
        <strong>${escapeHtml(r.col)}</strong> → ${labelMap[r.mode] || r.mode}
      </span>
      <button type="button" data-col="${escapeHtml(r.col)}" class="js-remove-case-rule"
              style="background:transparent;border:none;color:var(--danger);cursor:pointer;font-size:0.8rem;flex-shrink:0">삭제</button>
    </div>
  `,
      )
      .join(""),
  );
}

// 삭제 버튼은 동적 생성이므로 이벤트 위임
$(document).on("click", ".js-remove-case-rule", function () {
  removeCaseRule($(this).data("col"));
});

// ── 정제 실행 ─────────────────────────────────────────────────────────────────
export function runCleaning() {
  if (!rawData.length) {
    alert("파일을 먼저 업로드해주세요.");
    return;
  }

  let data = rawData.map((r) => ({ ...r }));
  const logs = [],
    orig = data.length;
  const removedReasonMap = new Map();

  // 빈 셀 처리
  if ($("#card-empty").hasClass("active")) {
    const mode = $('input[name="empty-mode"]:checked').val();
    const fill = $("#empty-fill-val").val();

    if (mode === "clear") {
      let n = 0;
      data.forEach((row) =>
        headers.forEach((h) => {
          if (isBlankLike(row[h])) {
            row[h] = null;
            n++;
          }
        }),
      );
      logs.push(`빈 문자열 ${n}개 셀 제거 (빈 칸으로 처리)`);
    } else if (mode === "remove") {
      const b = data.length;
      const kept = [];
      data.forEach((row) => {
        if (headers.every((h) => !isBlankLike(row[h]))) kept.push(row);
        else removedReasonMap.set(row[RAW_ROW_KEY], "빈 셀 포함 행 삭제");
      });
      data = kept;
      logs.push(`빈 셀 포함 행 ${b - data.length}개 삭제`);
    } else {
      let n = 0;
      data.forEach((row) =>
        headers.forEach((h) => {
          if (isBlankLike(row[h])) {
            row[h] = fill || "N/A";
            n++;
          }
        }),
      );
      logs.push(`빈 셀 ${n}개 → "${fill || "N/A"}"로 채움`);
    }
  }

  // 텍스트 정리
  if ($("#card-text").hasClass("active")) {
    const ghost = $("#txt-ghost").is(":checked");
    const trim = $("#txt-trim").is(":checked");
    const multi = $("#txt-multi").is(":checked");
    const special = $("#txt-special").is(":checked");
    const useCase = $("#txt-case").is(":checked");
    let ghostCount = 0;

    data.forEach((row) =>
      headers.forEach((h) => {
        if (row[h] !== null && row[h] !== undefined) {
          row[h] = String(row[h]);
          if (ghost) {
            const b = row[h];
            row[h] = cleanGhostText(row[h]);
            if (row[h] !== b) ghostCount++;
          }
          if (trim) row[h] = row[h].trim();
          if (multi) row[h] = row[h].replace(/\s+/g, " ");
          if (special) row[h] = row[h].replace(/[!@#$%^&*()\-_+=\[\]{};':"\\|,.<>\/?`~]/g, "");
          if (useCase) {
            const rule = caseRules.find((r) => r.col === h);
            if (rule?.mode === "upper") row[h] = row[h].toUpperCase();
            else if (rule?.mode === "lower") row[h] = row[h].toLowerCase();
            else if (rule?.mode === "first")
              row[h] = row[h]
                .split(" ")
                .map((w) => (w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : w))
                .join(" ");
          }
          if (isBlankLike(row[h])) row[h] = null;
        }
      }),
    );

    const txtOps = [ghost && `유령문자 ${ghostCount}개 제거`, trim && "공백 제거", multi && "연속공백 정리", special && "특수문자 제거", useCase && `대소문자 변환 ${caseRules.length}개 규칙 적용`].filter(Boolean);
    logs.push("텍스트 정리: " + txtOps.join(", "));
  }

  // 날짜 형식 통일
  if ($("#card-date").hasClass("active")) {
    const fmt = $('input[name="date-fmt"]:checked').val();
    const cols = $(".date-col-check:checked")
      .map((_, el) => el.value)
      .get();
    let n = 0,
      fail = 0;
    data.forEach((row) =>
      cols.forEach((h) => {
        const d = parseFlexibleDate(row[h]);
        if (d) {
          row[h] = formatDateValue(d, fmt);
          n++;
        } else if (!isBlankLike(row[h])) fail++;
      }),
    );
    logs.push(`날짜 ${n}개 → ${fmt} 변환${fail ? ` / 변환 실패 ${fail}개` : ""} (열: ${cols.join(", ") || "없음"})`);
  }

  // 금액 단위 변환
  if ($("#card-amount").hasClass("active")) {
    const cols = $(".amount-col-check:checked")
      .map((_, el) => el.value)
      .get();
    const fromUnit = $("#amount-from-unit").val() || "1";
    const toUnit = $("#amount-to-unit").val() || "1";
    const decimalsMode = $("#amount-decimals").val() || "auto";

    if (!cols.length) {
      alert("금액 단위 변환할 열을 하나 이상 체크해주세요.");
      return;
    }

    let n = 0,
      fail = 0;
    data.forEach((row) =>
      cols.forEach((h) => {
        if (isBlankLike(row[h])) return;
        const result = convertAmountUnit(row[h], fromUnit, toUnit, decimalsMode);
        if (result.ok) {
          row[h] = result.value;
          n++;
        } else fail++;
      }),
    );
    logs.push(`금액 ${n}개 단위 변환: ${AMOUNT_UNIT_LABELS[fromUnit] || fromUnit} → ${AMOUNT_UNIT_LABELS[toUnit] || toUnit}${fail ? ` / 변환 실패 ${fail}개` : ""} (열: ${cols.join(", ")})`);
  }

  // 전화번호 정제
  if ($("#card-phone").hasClass("active")) {
    const cols = $(".phone-col-check:checked")
      .map((_, el) => el.value)
      .get();
    const doFormat = $("#phone-format").is(":checked");
    const doSplit = $("#phone-split").is(":checked");

    if (!cols.length) {
      alert("전화번호 정제할 열을 하나 이상 체크해주세요.");
      return;
    }

    let n = 0,
      fail = 0;
    cols.forEach((h) => {
      if (doSplit) {
        const typeCol = `${h}_유형`;
        if (!headers.includes(typeCol)) headers.push(typeCol);
      }
    });

    data.forEach((row) =>
      cols.forEach((h) => {
        if (isBlankLike(row[h])) return;
        const result = normalizePhoneNumber(row[h]);
        if (result.valid && result.formatted) {
          if (doFormat) row[h] = result.formatted;
          if (doSplit) row[`${h}_유형`] = result.type;
          n++;
        } else {
          if (doSplit) row[`${h}_유형`] = "확인필요";
          fail++;
        }
      }),
    );
    logs.push(`전화번호 ${n}개 정제 완료${fail ? ` / 확인 필요 ${fail}개` : ""} (열: ${cols.join(", ")})`);
  }

  // 데이터 필터링
  if ($("#card-filter").hasClass("active")) {
    const col = $("#filter-col-select").val();
    const op = $("#filter-op").val();
    const val = $("#filter-val").val();
    const b = data.length;
    const kept = [];

    data.forEach((row) => {
      const cell = cleanGhostText(row[col] ?? "").trim();
      const num = parseFloat(cell);
      let keep = true;
      if (op === "contains") keep = cell.includes(val);
      else if (op === "not_contains") keep = !cell.includes(val);
      else if (op === "equals") keep = cell === val;
      else if (op === "starts") keep = cell.startsWith(val);
      else if (op === "ends") keep = cell.endsWith(val);
      else if (op === "gt") keep = !isNaN(num) && num > parseFloat(val);
      else if (op === "lt") keep = !isNaN(num) && num < parseFloat(val);

      if (keep) kept.push(row);
      else removedReasonMap.set(row[RAW_ROW_KEY], "필터링 조건 제외");
    });
    data = kept;
    logs.push(`[${col}] 필터 → ${data.length}행 추출, ${b - data.length}행 제외`);
  }

  // 중복 제거
  if ($("#card-dup").hasClass("active")) {
    const mode = $('input[name="dup-mode"]:checked').val();
    const cols =
      mode === "col"
        ? $(".dup-col-check:checked")
            .map((_, el) => el.value)
            .get()
        : headers;
    if (mode === "col" && !cols.length) {
      alert("중복 제거 기준 열을 하나 이상 체크해주세요.");
      return;
    }

    const seen = new Set(),
      b = data.length;
    const kept = [];
    data.forEach((row) => {
      const k = cols.map((h) => cleanGhostText(row[h] ?? "").trim()).join("||");
      if (seen.has(k)) removedReasonMap.set(row[RAW_ROW_KEY], "중복 제거");
      else {
        seen.add(k);
        kept.push(row);
      }
    });
    data = kept;
    logs.push(`중복 ${b - data.length}행 제거 (기준: ${mode === "all" ? "전체 열" : cols.join(", ")})`);
  }

  window._resultData = data;
  window._removedReasonMap = removedReasonMap;
  latestChangeInfo = buildChangeInfo(data);
  window._changeInfo = latestChangeInfo;
  resultPreviewMode = "result";

  const info = latestChangeInfo;
  $("#statsRow").html(`<div class="stat-pill">📄 원본 ${orig}행</div>` + `<div class="stat-pill green">✅ 결과 ${data.length}행</div>` + `<div class="stat-pill green">✏️ 변경 셀 ${info.changedCount}개</div>` + `<div class="stat-pill red">🗑 제거 ${orig - data.length}행</div>`);
  $("#logList").html(logs.map((l) => `<li>${escapeHtml(l)}</li>`).join("") || "<li>적용된 옵션 없음</li>");
  $("#resultViewToolbar").css("display", "flex");

  renderResultPreview();
  goToStep(3);
}

// ── 결과 미리보기 ─────────────────────────────────────────────────────────────
export function setResultView(mode) {
  resultPreviewMode = mode;
  renderResultPreview();
}

function renderResultPreview() {
  const rows = window._resultData || [];
  const info = latestChangeInfo || buildChangeInfo(rows);
  latestChangeInfo = info;
  _updateResultViewButtons(info);

  if (resultPreviewMode === "changed") {
    renderCleanedExcelPreview("resultTable", rows, 50, { changedSet: info.changedSet, detailMap: info.detailMap, onlyChanged: true, note: "변경된 셀이 있는 행만 표시됩니다. 변경되지 않은 셀은 비워두고, 변경 셀은 주황색 테두리로 표시됩니다." });
    return;
  }
  if (resultPreviewMode === "compare") {
    renderComparePreview("resultTable", rows, info, 30);
    return;
  }
  renderCleanedExcelPreview("resultTable", rows, 50, { changedSet: info.changedSet, detailMap: info.detailMap, note: "변경된 셀은 주황색 테두리와 모서리 표시로 확인할 수 있습니다." });
}

function _updateResultViewButtons(info) {
  ["result", "changed", "compare"].forEach((mode) => {
    $(`#btn-view-${mode}`).toggleClass("active", resultPreviewMode === mode);
  });
  $("#diffLegend").html(`<span><span class="legend-dot"></span>변경 셀 ${info?.changedCount || 0}개</span>` + `<span>· 변경 행 ${info?.changedRowCount || 0}개</span>` + `<span>· 삭제/제외 행 ${info?.removedCount || 0}개</span>`);
}

// ── diff 헬퍼 ─────────────────────────────────────────────────────────────────
function diffKey(rowOrNum, header) {
  const num = typeof rowOrNum === "object" ? rowOrNum?.[RAW_ROW_KEY] : rowOrNum;
  return String(num || "") + "||" + String(header || "");
}

function normalizeDiffValue(v) {
  return v === null || v === undefined ? "" : String(v);
}

function isSameForDiff(before, after) {
  const b = normalizeDiffValue(before),
    a = normalizeDiffValue(after);
  if (isBlankLike(b) && isBlankLike(a)) return true;
  return b === a;
}

function formatDiffValue(v) {
  if (v === null || v === undefined || isBlankLike(v)) return "(빈칸)";
  return String(v);
}

function buildChangeInfo(rows) {
  const rawByRowNum = new Map(rawData.map((r) => [r[RAW_ROW_KEY], r]));
  const resultRowNums = new Set(rows.map((r) => r[RAW_ROW_KEY]));
  const changedSet = new Set();
  const detailMap = new Map();
  const details = [];

  rows.forEach((row, resultIdx) => {
    const sourceRowNum = row[RAW_ROW_KEY];
    const beforeRow = rawByRowNum.get(sourceRowNum) || {};
    headers.forEach((h) => {
      if (!isSameForDiff(beforeRow[h], row[h])) {
        const key = diffKey(sourceRowNum, h);
        const detail = { key, sourceRowNum, resultRowNum: resultIdx + 2, column: h, before: beforeRow[h], after: row[h] };
        changedSet.add(key);
        detailMap.set(key, detail);
        details.push(detail);
      }
    });
  });

  const changedRowNums = new Set(details.map((d) => d.sourceRowNum));
  const removedReasonMap = window._removedReasonMap || new Map();
  const removedRows = rawData.filter((r) => !resultRowNums.has(r[RAW_ROW_KEY])).map((r) => ({ ...r, __paosRemoveReason: removedReasonMap.get(r[RAW_ROW_KEY]) || "중복 제거/필터링/빈셀 삭제 등으로 제외됨" }));

  return { changedSet, detailMap, details, changedCount: details.length, changedRowCount: changedRowNums.size, removedRows, removedCount: removedRows.length };
}

// ── 비교 미리보기 ─────────────────────────────────────────────────────────────
function renderComparePreview(id, rows, info, maxDataRows = 30) {
  const $wrap = $(`#${id}`);
  const originalId = `${id}-original`,
    resultId = `${id}-cleaned`,
    summaryId = `${id}-summary`;

  $wrap.html(`
    <div class="compare-grid">
      <div class="compare-panel">
        <div class="compare-panel-title">📄 원본</div>
        <div id="${originalId}"></div>
      </div>
      <div class="compare-panel">
        <div class="compare-panel-title">✅ 정제 결과</div>
        <div id="${resultId}"></div>
      </div>
    </div>
    <div id="${summaryId}"></div>
  `);

  const originalRowsForResult = rows.map((row) => rawData.find((r) => r[RAW_ROW_KEY] === row[RAW_ROW_KEY])).filter(Boolean);
  renderStyledRowsPreview(originalId, originalRowsForResult, maxDataRows, false, { valuesFromOriginal: true, changedSet: info.changedSet, detailMap: info.detailMap, highlightOriginal: true, note: "결과에 남아있는 행 기준으로 원본 값을 표시합니다." });
  renderStyledRowsPreview(resultId, rows, maxDataRows, false, { changedSet: info.changedSet, detailMap: info.detailMap, note: "변경된 셀은 주황색 테두리로 표시됩니다." });
  renderDiffSummary(summaryId, info);
}

function renderDiffSummary(id, info) {
  const $wrap = $(`#${id}`);
  const shown = info.details.slice(0, 120);
  const rowsHtml = shown
    .map(
      (d) => `
    <tr>
      <td>원본 ${escapeHtml(d.sourceRowNum)}행</td>
      <td>${escapeHtml(d.column)}</td>
      <td class="diff-before">${escapeHtml(formatDiffValue(d.before))}</td>
      <td class="diff-after">${escapeHtml(formatDiffValue(d.after))}</td>
    </tr>
  `,
    )
    .join("");

  const removedHtml = info.removedRows
    .slice(0, 60)
    .map((r) => {
      const preview = (originalHeaders.length ? originalHeaders : headers)
        .slice(0, 4)
        .map((h) => formatDiffValue(r[h]))
        .join(" / ");
      const reason = r.__paosRemoveReason || "삭제/제외됨";
      return `<tr><td>원본 ${escapeHtml(r[RAW_ROW_KEY])}행</td><td colspan="3">${escapeHtml(reason)} · ${escapeHtml(preview)}</td></tr>`;
    })
    .join("");

  const moreChanged = info.details.length > shown.length ? '<div class="excel-preview-note">변경 셀이 많아 상위 120개만 목록에 표시됩니다. 다운로드 결과에는 전체가 반영됩니다.</div>' : "";
  const moreRemoved = info.removedRows.length > 60 ? '<div class="excel-preview-note">삭제/제외 행이 많아 상위 60개만 목록에 표시됩니다.</div>' : "";

  $wrap.html(`
    <div class="diff-summary-card">
      <h4>변경 상세</h4>
      <div class="diff-table-wrap">
        <table class="diff-table">
          <thead><tr><th>행</th><th>열</th><th>원본</th><th>결과</th></tr></thead>
          <tbody>
            ${rowsHtml || '<tr><td colspan="4" style="color:var(--muted)">변경된 셀이 없습니다.</td></tr>'}
            ${removedHtml}
          </tbody>
        </table>
      </div>
      ${moreChanged}${moreRemoved}
    </div>
  `);
}

// ── 스타일드 미리보기 (ExcelJS 서식 반영) ────────────────────────────────────
function renderExcelPreview(id, maxDataRows = 10) {
  renderStyledRowsPreview(id, rawData, maxDataRows, true);
}
function renderCleanedExcelPreview(id, rows, maxDataRows = 50, options = {}) {
  renderStyledRowsPreview(id, rows, maxDataRows, false, options);
}

function renderStyledRowsPreview(id, rows, maxDataRows = 10, useOriginalValues = false, options = {}) {
  const $wrap = $(`#${id}`);
  if (!$wrap.length) return;

  if (!originalExcelJsWorksheet) {
    renderTable(id, headers, useOriginalValues ? rawData.slice(0, maxDataRows) : rows.slice(0, maxDataRows));
    return;
  }

  let sourceRows = useOriginalValues ? rawData : rows;
  if (options.onlyChanged && options.changedSet) {
    sourceRows = sourceRows.filter((row) => headers.some((h) => options.changedSet.has(diffKey(row, h))));
  }

  if (!sourceRows.length) {
    $wrap.html('<p style="padding:16px;color:var(--muted)">표시할 변경 셀이 없습니다.</p>');
    return;
  }

  const dataRowCount = Math.min(sourceRows.length, maxDataRows);
  const totalCols = headers.length;

  let html = '<div class="excel-preview-wrap"><table class="excel-preview-table"><colgroup>';
  html += '<col style="width:42px;min-width:42px">';
  for (let c = 0; c < totalCols; c++) {
    const srcCol = originalExcelJsWorksheet.getColumn(Math.min(c + 1, Math.max(1, originalExcelJsWorksheet.columnCount || totalCols)));
    const width = srcCol?.width ? Math.round(srcCol.width * 7 + 5) : 110;
    html += `<col style="width:${width}px;min-width:${width}px">`;
  }
  html += '</colgroup><thead><tr><th class="excel-corner"></th>';
  for (let c = 0; c < totalCols; c++) html += `<th class="excel-col-head">${excelColumnName(c)}</th>`;
  html += "</tr></thead><tbody>";

  for (let r = 0; r <= dataRowCount; r++) {
    const dataRow = r === 0 ? null : sourceRows[r - 1];
    const sourceRowNum = r === 0 ? 1 : dataRow?.[RAW_ROW_KEY] || r + 1;
    const srcRow = originalExcelJsWorksheet.getRow(sourceRowNum);
    const hpx = srcRow?.height ? Math.round(srcRow.height * 1.333) : "";

    html += `<tr${hpx ? ` style="height:${hpx}px"` : ""}><th class="excel-row-head">${r + 1}</th>`;

    for (let c = 0; c < totalCols; c++) {
      const header = headers[c];
      const sourceCell = _getPreviewSourceCell(sourceRowNum, c + 1);
      const style = excelJsCellStyleToCss(sourceCell);
      const value = r === 0 ? header : (dataRow?.[header] ?? "");
      const key = r === 0 ? "" : diffKey(dataRow, header);
      const isChanged = r > 0 && options.changedSet?.has(key);
      const classes = [];

      if (isChanged && (!useOriginalValues || options.highlightOriginal || options.valuesFromOriginal)) classes.push("cell-changed");
      let displayValue = value;
      if (options.onlyChanged && r > 0 && !isChanged) {
        displayValue = "";
        classes.push("cell-unchanged-hidden");
      }

      const detail = isChanged && options.detailMap ? options.detailMap.get(key) : null;
      const title = detail ? `원본: ${formatDiffValue(detail.before)}\n결과: ${formatDiffValue(detail.after)}` : displayValue;

      html += `<td${classes.length ? ` class="${classes.join(" ")}"` : ""}${style ? ` style="${style}"` : ""} title="${escapeHtml(title)}">${escapeHtml(displayValue)}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table></div>";

  if (sourceRows.length > maxDataRows) html += `<div class="excel-preview-note">상위 ${maxDataRows}개 데이터 행만 미리보기로 표시됩니다.</div>`;
  if (options.note) html += `<div class="excel-preview-note">${escapeHtml(options.note)}</div>`;

  $wrap.html(html);
}

// ── ExcelJS 유틸 ─────────────────────────────────────────────────────────────
function excelColumnName(index) {
  let name = "",
    n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function _getPreviewSourceCell(sourceRowNum, colNum) {
  if (!originalExcelJsWorksheet) return null;
  const maxCol = Math.max(1, originalExcelJsWorksheet.columnCount || headers.length || 1);
  return originalExcelJsWorksheet.getCell(sourceRowNum, Math.min(Math.max(1, colNum), maxCol));
}

function _excelJsColorToCss(color) {
  if (!color) return "";
  let hex = color.argb || color.rgb;
  if (!hex) return "";
  hex = String(hex).replace(/^#/, "");
  if (hex.length === 8) hex = hex.slice(2);
  return hex.length === 6 ? "#" + hex : "";
}

function _excelJsBorderToCss(side) {
  if (!side?.style) return "";
  const color = _excelJsColorToCss(side.color) || "#d9d9d9";
  const map = { thin: "1px solid", hair: "1px solid", medium: "2px solid", thick: "3px solid", dashed: "1px dashed", dashDot: "1px dashed", dashDotDot: "1px dashed", dotted: "1px dotted", double: "3px double", mediumDashed: "2px dashed", mediumDashDot: "2px dashed", mediumDashDotDot: "2px dashed", slantDashDot: "2px dashed" };
  return (map[side.style] || "1px solid") + " " + color;
}

function excelJsCellStyleToCss(cell) {
  if (!cell) return "";
  const css = [];
  const font = cell.font || cell.style?.font;
  const fill = cell.fill || cell.style?.fill;
  const border = cell.border || cell.style?.border;
  const alignment = cell.alignment || cell.style?.alignment;

  if (fill?.type === "pattern" && fill.pattern !== "none") {
    const bg = _excelJsColorToCss(fill.fgColor || fill.bgColor);
    if (bg) css.push("background-color:" + bg);
  }
  if (font) {
    const color = _excelJsColorToCss(font.color);
    if (color) css.push("color:" + color);
    if (font.bold) css.push("font-weight:700");
    if (font.italic) css.push("font-style:italic");
    if (font.size) css.push("font-size:" + font.size + "pt");
    if (font.name) css.push("font-family:" + String(font.name).replace(/;/g, "") + ", Calibri, Arial, sans-serif");
    if (font.underline) css.push("text-decoration:underline");
    if (font.strike) css.push("text-decoration:line-through");
  }
  if (alignment) {
    if (alignment.horizontal) css.push("text-align:" + alignment.horizontal);
    if (alignment.vertical) css.push("vertical-align:" + alignment.vertical);
    if (alignment.wrapText) css.push("white-space:normal");
  }
  if (border) {
    ["top", "right", "bottom", "left"].forEach((side) => {
      const b = _excelJsBorderToCss(border[side]);
      if (b) css.push(`border-${side}:${b}`);
    });
  }
  return css.join(";");
}

function _getExcelJsCellText(cell) {
  if (!cell) return "";
  if (cell.text !== undefined && cell.text !== null) return cell.text;
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text || "").join("");
    if (v.text !== undefined) return v.text;
    if (v.result !== undefined) return v.result;
    if (v instanceof Date) return v.toLocaleDateString("ko-KR");
  }
  return String(v);
}

// ── 다운로드 ──────────────────────────────────────────────────────────────────
export async function downloadResult() {
  if (!window._resultData) return;

  if (/\.xlsx$/i.test(fileName) && originalExcelJsWorksheet && window.ExcelJS) {
    const wb = await buildStyledWorkbookWithExcelJs(window._resultData);
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const $a = $("<a>")
      .attr({ href: url, download: `정제결과_${fileName.replace(/\.(xls|csv)$/i, ".xlsx")}` })
      .appendTo("body");
    $a[0].click();
    $a.remove();
    URL.revokeObjectURL(url);
    return;
  }

  const wb = buildSheetJsWorkbook(window._resultData);
  XLSX.writeFile(wb, `정제결과_${fileName.replace(/\.(xls|csv)$/i, ".xlsx")}`);
}

// ── 초기화 ────────────────────────────────────────────────────────────────────
export function resetExcel() {
  rawData = [];
  headers = [];
  originalHeaders = [];
  fileName = "";
  window._resultData = null;
  window._changeInfo = null;
  window._removedReasonMap = null;
  caseRules = [];
  originalFileArrayBuffer = null;
  originalExcelJsWorkbook = null;
  originalExcelJsWorksheet = null;
  resultPreviewMode = "result";
  latestChangeInfo = null;

  $("#fileInput").val("");
  const $zone = $("#uploadZone");
  $zone.css({ borderColor: "", background: "" });
  $zone.find("h3").text("파일을 여기에 드래그하거나 클릭하여 업로드");
  $zone.find("p").text(".xlsx · .xls · .csv 지원");

  $(".opt-card").removeClass("active");
  $(".opt-check").css("color", "transparent");
  $(".opt-sub").removeClass("show");

  $("#previewBox").hide();
  $("#step2-btns").hide();
  $("#col-settings").hide();
  $("#resultViewToolbar").hide();
  $("#resultTable").empty();
  $("#case-rule-list").empty();
  $("#step1-warn").hide();

  goToStep(1);
}

// ── 워크북 빌드 (ExcelJS) ─────────────────────────────────────────────────────
async function buildStyledWorkbookWithExcelJs(rows) {
  const outWb = new ExcelJS.Workbook();
  const srcWs = originalExcelJsWorksheet;
  const info = latestChangeInfo || buildChangeInfo(rows);
  const reviewHeaders = ["검수_변경여부", "검수_변경컬럼", "검수_변경내용"];
  const totalCols = headers.length + reviewHeaders.length;

  outWb.creator = "PAOS WorkLab";
  outWb.created = new Date();
  outWb.modified = new Date();

  addCleanDataSheet(outWb, rows);

  const outWs = outWb.addWorksheet("정제결과");
  headers.forEach((h, idx) => {
    const srcColIdx = Math.min(idx + 1, Math.max(1, srcWs.columnCount || headers.length));
    outWs.getColumn(idx + 1).width = srcWs.getColumn(srcColIdx)?.width || 15;
  });
  outWs.getColumn(headers.length + 1).width = 14;
  outWs.getColumn(headers.length + 2).width = 30;
  outWs.getColumn(headers.length + 3).width = 70;

  outWs.getRow(1).height = srcWs.getRow(1).height || 22;
  headers.forEach((h, idx) => copyExcelJsCellWithValue(_getPreviewSourceCell(1, idx + 1), outWs.getCell(1, idx + 1), h));
  reviewHeaders.forEach((h, idx) => {
    const cell = outWs.getCell(1, headers.length + idx + 1);
    cell.value = h;
    styleAuditHeaderCell(cell);
  });

  rows.forEach((row, rIdx) => {
    const outRowNum = rIdx + 2,
      srcRowNum = row[RAW_ROW_KEY] || outRowNum;
    const outRow = outWs.getRow(outRowNum);
    outRow.height = srcWs.getRow(srcRowNum).height;

    headers.forEach((h, cIdx) => {
      const srcCell = _getPreviewSourceCell(srcRowNum, cIdx + 1);
      const outCell = outWs.getCell(outRowNum, cIdx + 1);
      copyExcelJsCellWithValue(srcCell, outCell, row[h]);
      const detail = info.detailMap?.get(diffKey(srcRowNum, h));
      if (detail) applyChangedCellHighlight(outCell, detail);
    });

    const summary = getRowChangeSummary(srcRowNum, info);
    const reviewValues = [summary.changed ? "변경" : "변경없음", summary.columns, summary.detail];
    reviewValues.forEach((value, idx) => {
      const cell = outWs.getCell(outRowNum, headers.length + idx + 1);
      cell.value = value || null;
      styleAuditBodyCell(cell);
      if (idx === 0 && summary.changed) {
        cell.font = { bold: true, color: { argb: "FFB45309" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF7ED" } };
      }
    });
  });

  freezeTopRow(outWs);
  setSheetAutoFilter(outWs, rows.length + 1, totalCols);
  addChangeLogSheet(outWb, info);
  addRemovedRowsSheet(outWb, info);
  copyOriginalWorksheetTo(outWb);

  return outWb;
}

function buildSheetJsWorkbook(rows) {
  const info = latestChangeInfo || buildChangeInfo(rows);
  const wb = XLSX.utils.book_new();
  const cleanRows = rows.map((row) => {
    const obj = {};
    headers.forEach((h) => {
      obj[h] = row[h] ?? "";
    });
    return obj;
  });
  const resultRows = buildResultRowsForSheetJs(rows, info);
  const resultHdrs = [...headers, "검수_변경여부", "검수_변경컬럼", "검수_변경내용"];

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cleanRows, { header: headers }), "정제데이터");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resultRows, { header: resultHdrs }), "정제결과");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildChangeRowsForSheetJs(info), { header: ["원본행번호", "결과행번호", "컬럼명", "변경전", "변경후", "정제유형"] }), "변경내역");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildRemovedRowsForSheetJs(info), { header: ["원본행번호", "제외사유", ...(originalHeaders.length ? originalHeaders : headers)] }), "제외삭제행");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildOriginalRowsForSheetJs(), { header: originalHeaders.length ? originalHeaders : headers }), "원본데이터");
  return wb;
}

// ── SheetJS 데이터 빌더 ───────────────────────────────────────────────────────
function buildResultRowsForSheetJs(rows, info) {
  return rows.map((row) => {
    const obj = {};
    headers.forEach((h) => {
      obj[h] = row[h] ?? "";
    });
    const summary = getRowChangeSummary(row[RAW_ROW_KEY], info);
    obj["검수_변경여부"] = summary.changed ? "변경" : "변경없음";
    obj["검수_변경컬럼"] = summary.columns;
    obj["검수_변경내용"] = summary.detail;
    return obj;
  });
}

function buildChangeRowsForSheetJs(info) {
  return (info?.details || []).map((d) => ({
    원본행번호: d.sourceRowNum,
    결과행번호: d.resultRowNum,
    컬럼명: d.column,
    변경전: formatDiffValue(d.before),
    변경후: formatDiffValue(d.after),
    정제유형: inferCleanType(d),
  }));
}

function buildRemovedRowsForSheetJs(info) {
  const cols = originalHeaders.length ? originalHeaders : headers;
  return (info?.removedRows || []).map((r) => {
    const obj = { 원본행번호: r[RAW_ROW_KEY], 제외사유: r.__paosRemoveReason || "삭제/제외됨" };
    cols.forEach((h) => {
      obj[h] = r[h] ?? "";
    });
    return obj;
  });
}

function buildOriginalRowsForSheetJs() {
  const cols = originalHeaders.length ? originalHeaders : headers;
  return rawData.map((r) => {
    const obj = {};
    cols.forEach((h) => {
      obj[h] = r[h] ?? "";
    });
    return obj;
  });
}

function getRowChangeSummary(sourceRowNum, info) {
  const rowDetails = (info?.details || []).filter((d) => d.sourceRowNum === sourceRowNum);
  if (!rowDetails.length) return { changed: false, columns: "", detail: "" };
  const columns = [...new Set(rowDetails.map((d) => d.column))].join(", ");
  const detail = rowDetails.map((d) => `${d.column}: ${formatDiffValue(d.before)} → ${formatDiffValue(d.after)}`).join(" / ");
  return { changed: true, columns, detail: _truncateExcelText(detail) };
}

function inferCleanType(detail) {
  const col = String(detail.column || "");
  const before = formatDiffValue(detail.before),
    after = formatDiffValue(detail.after);
  if (col.endsWith("_유형")) return "전화번호 유형 추가";
  if (/금액|매출|매출액|비용|단가|가격|원|amount|sales|revenue|price|cost/i.test(col)) return "금액 단위 변환";
  if (/전화|phone|mobile|tel/i.test(col)) return "전화번호 정제";
  if (/일자|날짜|date|가입일|등록일|설립일/i.test(col)) return "날짜 형식 통일";
  if (before === "(빈칸)" && after !== "(빈칸)") return "빈 셀 채우기";
  if (before !== "(빈칸)" && after === "(빈칸)") return "빈 셀 비우기";
  return "텍스트/값 정리";
}

// ── ExcelJS 시트 빌더 ─────────────────────────────────────────────────────────
function addCleanDataSheet(outWb, rows) {
  const ws = outWb.addWorksheet("정제데이터");
  const srcWs = originalExcelJsWorksheet;

  headers.forEach((h, idx) => {
    const srcColIdx = Math.min(idx + 1, Math.max(1, srcWs?.columnCount || headers.length));
    const srcCol = srcWs?.getColumn(srcColIdx);
    const outCol = ws.getColumn(idx + 1);
    outCol.width = srcCol?.width || 15;
    outCol.hidden = !!srcCol?.hidden;
    if (srcCol?.style) outCol.style = cloneExcelJsStyle(srcCol.style);
  });

  ws.getRow(1).height = srcWs?.getRow(1)?.height || 22;
  headers.forEach((h, idx) => copyExcelJsCellWithValue(_getPreviewSourceCell(1, idx + 1), ws.getCell(1, idx + 1), h));

  rows.forEach((row, rIdx) => {
    const outRowNum = rIdx + 2,
      srcRowNum = row[RAW_ROW_KEY] || outRowNum;
    const srcRow = srcWs?.getRow(srcRowNum);
    const outRow = ws.getRow(outRowNum);
    if (srcRow?.height) outRow.height = srcRow.height;
    headers.forEach((h, cIdx) => copyExcelJsCellWithValue(_getPreviewSourceCell(srcRowNum, cIdx + 1), ws.getCell(outRowNum, cIdx + 1), row[h]));
  });

  freezeTopRow(ws);
  setSheetAutoFilter(ws, rows.length + 1, headers.length);
  return ws;
}

function addChangeLogSheet(outWb, info) {
  const ws = outWb.addWorksheet("변경내역");
  const hdrs = ["원본행번호", "결과행번호", "컬럼명", "변경전", "변경후", "정제유형"];
  ws.addRow(hdrs);
  applyAuditHeaderRow(ws, hdrs.length);
  (info.details || []).forEach((d) => {
    const row = ws.addRow([d.sourceRowNum, d.resultRowNum, d.column, formatDiffValue(d.before), formatDiffValue(d.after), inferCleanType(d)]);
    row.eachCell((cell) => styleAuditBodyCell(cell));
  });
  ws.columns = [{ width: 12 }, { width: 12 }, { width: 22 }, { width: 32 }, { width: 32 }, { width: 18 }];
  freezeTopRow(ws);
  setSheetAutoFilter(ws, Math.max(1, (info.details || []).length + 1), hdrs.length);
}

function addRemovedRowsSheet(outWb, info) {
  const ws = outWb.addWorksheet("제외삭제행");
  const baseHeaders = originalHeaders.length ? originalHeaders : headers;
  const removedHdrs = ["원본행번호", "제외사유", ...baseHeaders];
  ws.addRow(removedHdrs);
  applyAuditHeaderRow(ws, removedHdrs.length);
  (info.removedRows || []).forEach((r) => {
    const row = ws.addRow([r[RAW_ROW_KEY], r.__paosRemoveReason || "삭제/제외됨", ...baseHeaders.map((h) => r[h] ?? "")]);
    row.eachCell((cell) => styleAuditBodyCell(cell));
  });
  ws.getColumn(1).width = 12;
  ws.getColumn(2).width = 24;
  baseHeaders.forEach((h, idx) => {
    const srcCol = originalExcelJsWorksheet?.getColumn(idx + 1);
    ws.getColumn(idx + 3).width = srcCol?.width || Math.min(Math.max(String(h).length + 4, 12), 28);
  });
  freezeTopRow(ws);
  setSheetAutoFilter(ws, Math.max(1, (info.removedRows || []).length + 1), removedHdrs.length);
}

function copyOriginalWorksheetTo(outWb) {
  if (!originalExcelJsWorksheet) return null;
  const srcWs = originalExcelJsWorksheet;
  const outWs = outWb.addWorksheet("원본데이터");
  const maxRow = Math.max(srcWs.rowCount || 0, 1);
  const maxCol = Math.max(srcWs.columnCount || originalHeaders.length || headers.length || 0, 1);

  for (let c = 1; c <= maxCol; c++) {
    const srcCol = srcWs.getColumn(c),
      outCol = outWs.getColumn(c);
    outCol.width = srcCol?.width || 15;
    outCol.hidden = !!srcCol?.hidden;
    if (srcCol?.style) outCol.style = cloneExcelJsStyle(srcCol.style);
  }
  for (let r = 1; r <= maxRow; r++) {
    const srcRow = srcWs.getRow(r),
      outRow = outWs.getRow(r);
    outRow.height = srcRow.height;
    outRow.hidden = !!srcRow.hidden;
    for (let c = 1; c <= maxCol; c++) {
      const srcCell = srcRow.getCell(c),
        outCell = outRow.getCell(c);
      outCell.value = cloneExcelJsValue(srcCell.value);
      if (srcCell.style) outCell.style = cloneExcelJsStyle(srcCell.style);
      if (srcCell.note) outCell.note = cloneExcelJsValue(srcCell.note);
    }
  }
  try {
    const merges = srcWs._merges ? Object.values(srcWs._merges) : [];
    merges.forEach((m) => {
      const range = m.range || m.model?.range;
      if (range) outWs.mergeCells(range);
    });
  } catch (e) {
    console.warn("원본 병합 셀 복사 생략:", e);
  }
  if (srcWs.autoFilter) outWs.autoFilter = cloneExcelJsValue(srcWs.autoFilter);
  return outWs;
}

// ── ExcelJS 스타일 헬퍼 ───────────────────────────────────────────────────────
function cloneExcelJsStyle(s) {
  return s ? JSON.parse(JSON.stringify(s)) : {};
}
function cloneExcelJsValue(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return new Date(v.getTime());
  if (typeof v === "object") return JSON.parse(JSON.stringify(v));
  return v;
}
function safeExcelValue(v) {
  return v === null || v === undefined || v === "" ? null : v;
}
function copyExcelJsCellWithValue(src, tgt, value) {
  if (src?.style) tgt.style = cloneExcelJsStyle(src.style);
  tgt.value = safeExcelValue(value);
}

function applyChangedCellHighlight(cell, detail) {
  const orange = { style: "medium", color: { argb: "FFF59E0B" } };
  cell.border = { top: orange, right: orange, bottom: orange, left: orange };
  if (detail) cell.note = `원본: ${formatDiffValue(detail.before)}\n결과: ${formatDiffValue(detail.after)}`;
}

function styleAuditHeaderCell(cell) {
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  const thin = { style: "thin", color: { argb: "FFD9D9D9" } };
  cell.border = { top: thin, right: thin, bottom: thin, left: thin };
}

function styleAuditBodyCell(cell) {
  cell.alignment = { vertical: "top", wrapText: true };
  const thin = { style: "thin", color: { argb: "FFE5E7EB" } };
  cell.border = { top: thin, right: thin, bottom: thin, left: thin };
}

function applyAuditHeaderRow(ws, count) {
  const row = ws.getRow(1);
  row.height = Math.max(row.height || 18, 22);
  for (let c = 1; c <= count; c++) styleAuditHeaderCell(row.getCell(c));
}

function setSheetAutoFilter(ws, rowCount, colCount) {
  if (!rowCount || !colCount) return;
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, rowCount), column: colCount } };
}

function freezeTopRow(ws) {
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

function _truncateExcelText(text, limit = 32700) {
  const s = String(text ?? "");
  return s.length > limit ? s.slice(0, limit) + "…" : s;
}

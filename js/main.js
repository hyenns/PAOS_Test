// ─── main.js ───
// 진입점: 모든 모듈 import + DOM 준비 후 이벤트 바인딩

import { showScreen, goBack }             from './app.js';
import { initExcelCleaner, goToStep, toggleOpt, addCaseRule, runCleaning, setResultView, downloadResult, resetExcel } from './excel-cleaner.js';
import { initMerge, runMerge, downloadMergeResult, resetMerge } from './merge.js';

$(function () {
  // 모듈 초기화
  initExcelCleaner();
  initMerge();

  // ── 사이드바 / 홈 ──────────────────────────────────────────────────────────
  $('#sidebar-logo').on('click', () => showScreen('home'));
  $('#nav-excel').on('click',    () => showScreen('excel'));
  $('#nav-merge').on('click',    () => showScreen('merge'));
  $('#nav-crawling').on('click', () => showScreen('crawling'));

  // 홈 카드
  $('#home-card-excel').on('click', () => showScreen('excel'));
  $('#home-card-merge').on('click', () => showScreen('merge'));
  $('#home-card-crawling').on('click', () => showScreen('crawling'));

  // ── 공통 뒤로 가기 ────────────────────────────────────────────────────────
  $('.back-btn[data-back]').on('click', goBack);

  // ── 크롤링 ───────────────────────────────────────────────────────────────
  $('#btn-back-crawling').on('click', goBack);
  $('#btn-home-crawling').on('click', () => showScreen('home'));

  // ── 엑셀 정제 ─────────────────────────────────────────────────────────────
  // Topbar
  $('#btn-back-excel').on('click',  goBack);
  $('#btn-reset-excel').on('click', resetExcel);
  $('#btn-home-excel').on('click',  () => showScreen('home'));

  // 옵션 카드 토글
  $('#card-dup').on('click',    () => toggleOpt('dup'));
  $('#card-empty').on('click',  () => toggleOpt('empty'));
  $('#card-text').on('click',   () => toggleOpt('text'));
  $('#card-date').on('click',   () => toggleOpt('date'));
  $('#card-phone').on('click',  () => toggleOpt('phone'));
  $('#card-filter').on('click', () => toggleOpt('filter'));
  $('#card-amount').on('click', () => toggleOpt('amount'));
  $('#card-address').on('click', () => toggleOpt('address'));

  // Step 네비게이션
  $('#btn-to-step2').on('click', () => goToStep(2));
  $('#btn-back-to-step1').on('click', () => goToStep(1));
  $('#btn-run').on('click', runCleaning);

  // 결과 뷰 전환
  $('#btn-view-result').on('click',  () => setResultView('result'));
  $('#btn-view-changed').on('click', () => setResultView('changed'));
  $('#btn-view-compare').on('click', () => setResultView('compare'));

  // 다운로드 / 재시도
  $('#btn-download-result').on('click',  downloadResult);
  $('#btn-back-to-step2').on('click',   () => goToStep(2));
  $('#btn-reset-excel2').on('click',    resetExcel);

  // 대소문자 변환 규칙 추가
  $('#btn-add-case-rule').on('click', addCaseRule);

  // ── 파일 병합 ──────────────────────────────────────────────────────────────
  $('#btn-back-merge').on('click',          goBack);
  $('#btn-reset-merge').on('click',         resetMerge);
  $('#btn-home-merge').on('click',          () => showScreen('home'));
  $('#btn-run-merge').on('click',           runMerge);
  $('#btn-download-merge').on('click',      downloadMergeResult);
  $('#btn-reset-merge2').on('click',        resetMerge);
});

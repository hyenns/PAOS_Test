// Shared checkpoint UI for the IROS and Saramin panels.
export function createCrawlCheckpointUI(kind) {
  const key = `worklab.checkpoints.v1.${kind}`;
  const label = kind === 'iros' ? '등기소' : '사람인';
  const modeLabels = { company: '상호검색', registration: '등록번호검색', intro: '기업소개만', finance: '기업소개+재무정보' };
  let activeId = null;
  let memoryHistory = [];

  function readHistory() {
    try {
      const stored = JSON.parse(localStorage.getItem(key) || '[]');
      if (Array.isArray(stored)) {
        return [...memoryHistory, ...stored].filter((job, i, all) => job && /^[0-9a-f]{32}$/.test(job.id) && all.findIndex(x => x?.id === job.id) === i).slice(0, 30);
      }
    } catch (_) { /* Private browsing / disabled storage. */ }
    return memoryHistory;
  }

  function refreshHistory(preferredId) {
    const select = document.getElementById(`${kind}-checkpoint-history`);
    const button = document.getElementById(`btn-${kind}-checkpoint-download`);
    if (!select || !button) return;
    const oldValue = preferredId || select.value;
    const jobs = readHistory();
    select.replaceChildren();
    for (const job of jobs) {
      const option = document.createElement('option');
      option.value = job.id;
      const date = new Date(job.createdAt);
      const time = Number.isNaN(date.getTime()) ? (job.createdAt || '') : date.toLocaleString('ko-KR');
      option.textContent = `${time} · ${modeLabels[job.mode] || job.mode || label} · ${job.total ?? '?'}개 검색`;
      select.appendChild(option);
    }
    if (!jobs.length) {
      const option = document.createElement('option');
      option.value = '';
      option.textContent = '아직 저장된 작업이 없습니다.';
      select.appendChild(option);
    } else if (jobs.some(job => job.id === oldValue)) {
      select.value = oldValue;
    }
    select.disabled = !jobs.length;
    button.disabled = !jobs.length;
  }

  function showStatus(text) {
    const el = document.getElementById(`${kind}-checkpoint-status`);
    if (el) el.textContent = text;
  }

  function prepare() {
    activeId = null;
    showStatus('자동 저장 준비 중…');
  }

  function handleEvent(event) {
    if (event.type === 'checkpoint_start') {
      activeId = event.job_id;
      const job = { id: event.job_id, createdAt: event.created_at, mode: event.mode, total: event.total };
      memoryHistory = [job, ...readHistory().filter(x => x.id !== job.id)].slice(0, 30);
      try {
        localStorage.setItem(key, JSON.stringify(memoryHistory));
        showStatus('자동 저장 활성 · 수집된 결과 0건 저장');
      } catch (_) {
        showStatus('자동 저장 활성 · 브라우저 기록 저장 불가: 현재 창에서 중간결과를 내려받아 주세요.');
      }
      refreshHistory(activeId);
    } else if (event.type === 'checkpoint_status') {
      if (event.job_id !== activeId) return;
      showStatus(event.ok
        ? `자동 저장 · ${event.saved_count ?? 0}건 저장됨`
        : `⚠ 중간저장 실패 · 마지막 저장 ${event.saved_count ?? 0}건 (로그 확인)`);
    }
  }

  async function download(jobId) {
    if (!jobId) {
      alert('아직 현재 작업의 저장 파일이 없습니다. 저장된 이전 작업은 위 중간결과 메뉴에서 선택해 주세요.');
      return;
    }
    try {
      const res = await fetch(`/api/crawl/checkpoint/${encodeURIComponent(jobId)}`, { cache: 'no-store' });
      if (!res.ok) {
        let message = `중간결과 다운로드 실패 (HTTP ${res.status})`;
        try { const data = await res.json(); message = data.message || message; } catch (_) {}
        throw new Error(message);
      }
      if (!(res.headers.get('Content-Type') || '').includes('spreadsheetml')) {
        throw new Error('중간저장 API를 찾을 수 없습니다. 최신 WorkLab 서버로 다시 실행해 주세요.');
      }
      let filename = `${label}_중간결과_${jobId.slice(0, 8)}.xlsx`;
      const disposition = res.headers.get('Content-Disposition') || '';
      const match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
      if (match) { try { filename = decodeURIComponent(match[1]); } catch (_) {} }
      const url = URL.createObjectURL(await res.blob());
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) {
      alert(e.message);
    }
  }

  function init() {
    const button = document.getElementById(`btn-${kind}-checkpoint-download`);
    if (button) button.addEventListener('click', async () => {
      const id = document.getElementById(`${kind}-checkpoint-history`)?.value;
      button.disabled = true;
      try { await download(id); } finally { refreshHistory(); }
    });
    refreshHistory();
  }

  return { init, prepare, handleEvent, refreshHistory, downloadCurrent: () => download(activeId) };
}

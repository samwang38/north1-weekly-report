(() => {
  const $ = id => document.getElementById(id);

  const configCard   = $('configCard');
  const progressCard = $('progressCard');
  const downloadCard = $('downloadCard');
  const errorCard    = $('errorCard');

  const weekEndInput  = $('weekEnd');
  const weekRangeHint = $('weekRangeHint');
  const generateBtn   = $('generateBtn');
  const progressBar   = $('progressBar');
  const progressLabel = $('progressLabel');
  const logBox        = $('logBox');
  const downloadBtn   = $('downloadBtn');
  const downloadInfo  = $('downloadInfo');
  const resetBtn      = $('resetBtn');
  const retryBtn      = $('retryBtn');
  const errorMsg      = $('errorMsg');

  let pollTimer    = null;
  let currentJobId = null;
  let seenCount    = 0;   // how many messages we've already rendered

  function fmtDate(d) {
    return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
  }

  function updateHint(val) {
    if (!val) { weekRangeHint.textContent = ''; return; }
    const end   = new Date(val + 'T00:00:00');
    const start = new Date(end); start.setDate(end.getDate() - 6);
    weekRangeHint.textContent = `本週範圍：${fmtDate(start)} ～ ${fmtDate(end)}`;
  }

  weekEndInput.addEventListener('change', () => updateHint(weekEndInput.value));

  async function loadDefaultDate() {
    try {
      const res = await fetch('/api/default-date');
      const { date } = await res.json();
      weekEndInput.value = date;
      updateHint(date);
    } catch (e) {
      console.warn('無法取得預設日期', e);
    }
  }

  function show(card) {
    [configCard, progressCard, downloadCard, errorCard].forEach(c => {
      c.hidden = (c !== card);
    });
  }

  function appendLog(text) {
    const cls = text.includes('[OK]') || text.includes('✓') ? 'ok'
              : text.includes('[ERR]') || text.includes('✗') ? 'err'
              : 'info';
    const line = document.createElement('span');
    line.className = cls;
    line.textContent = text + '\n';
    logBox.appendChild(line);
    logBox.scrollTop = logBox.scrollHeight;
  }

  function stopPoll() {
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  // Rough progress estimate based on log content keywords
  function estimateProgress(messages) {
    const all = messages.join(' ');
    if (all.includes('儲存 Excel'))     return 95;
    if (all.includes('配件'))           return 75;
    if (all.includes('BY店'))           return 55;
    if (all.includes('指標計算完成'))   return 40;
    if (all.includes('去年資料'))       return 25;
    if (all.includes('今年資料'))       return 10;
    return 3;
  }

  async function pollStatus() {
    if (!currentJobId) return;
    try {
      const res  = await fetch(`/api/status?jobId=${currentJobId}`);
      const data = await res.json();

      // render only new messages
      const msgs = data.messages || [];
      for (let i = seenCount; i < msgs.length; i++) appendLog(msgs[i]);
      seenCount = msgs.length;

      const pct = estimateProgress(msgs);
      progressBar.style.width = pct + '%';

      if (data.status === 'done') {
        stopPoll();
        progressBar.style.width = '100%';
        progressLabel.textContent = '完成！';
        downloadInfo.textContent = `週報已產生完成（${data.filename || '北一區週報.xlsx'}）`;
        downloadBtn.onclick = () => {
          window.location.href = `/api/download?jobId=${currentJobId}`;
        };
        show(downloadCard);
      } else if (data.status === 'error') {
        stopPoll();
        errorMsg.textContent = data.error || '未知錯誤，請確認 VPN 已連線並重試。';
        show(errorCard);
      } else {
        progressLabel.textContent = data.status === 'pending' ? '排隊中…' : '處理中…';
      }
    } catch (e) {
      appendLog('[ERR] 連線中斷：' + e.message);
    }
  }

  // ── 自訂日期：區間定義 + 表單 ──
  const CP_FIELDS = [
    ['wk',     '本週'],
    ['pw',     '上週'],
    ['mtd',    '本月'],
    ['pm',     '上月'],
    ['lymo',   '去年同月'],
    ['ytd_ly', '去年年累積'],
    ['ytd_cy', '今年年累積'],
  ];

  function buildCustomGrid() {
    const grid = $('customGrid');
    if (grid.dataset.built) return;
    grid.innerHTML = CP_FIELDS.map(([k, label]) => `
      <div class="cp-row">
        <label>${label}</label>
        <input type="date" id="cp_${k}_s">
        <span class="cp-sep">～</span>
        <input type="date" id="cp_${k}_e">
      </div>`).join('');
    grid.dataset.built = '1';
  }

  async function prefillCustom() {
    // 用目前自動模式的週結束日 + 延伸月底，向後端取得預設區間帶入
    const wkEnd = weekEndInput.value || '';
    const fm = document.getElementById('useFullMonth')?.checked ? '1' : '0';
    try {
      const res = await fetch(`/api/periods?week_end=${wkEnd}&full_month=${fm}`);
      const p = await res.json();
      if (p.error) return;
      CP_FIELDS.forEach(([k]) => {
        if (p[k]) { $(`cp_${k}_s`).value = p[k][0]; $(`cp_${k}_e`).value = p[k][1]; }
      });
    } catch (e) { /* 預設帶入失敗不影響手動輸入 */ }
  }

  function collectCustomPeriods() {
    const cp = {};
    for (const [k, label] of CP_FIELDS) {
      const s = $(`cp_${k}_s`).value, e = $(`cp_${k}_e`).value;
      if (!s || !e) { alert(`請完整填寫「${label}」的起迄日期`); return null; }
      if (s > e)    { alert(`「${label}」的起日不能晚於迄日`); return null; }
      cp[k] = [s, e];
    }
    return cp;
  }

  function applyMode() {
    const isCustom = document.querySelector('input[name="dateMode"]:checked')?.value === 'custom';
    $('autoPane').hidden = isCustom;
    $('customPane').hidden = !isCustom;
    if (isCustom) { buildCustomGrid(); prefillCustom(); }
  }
  document.querySelectorAll('input[name="dateMode"]').forEach(r =>
    r.addEventListener('change', applyMode));

  generateBtn.addEventListener('click', async () => {
    const isCustom = document.querySelector('input[name="dateMode"]:checked')?.value === 'custom';
    let body;
    if (isCustom) {
      const cp = collectCustomPeriods();
      if (!cp) return;   // 驗證失敗（collectCustomPeriods 內已提示）
      body = { customPeriods: cp };
    } else {
      const wkEnd = weekEndInput.value;
      if (!wkEnd) { alert('請選擇週結束日期'); return; }
      const useFullMonth = !!document.getElementById('useFullMonth')?.checked;
      body = { week_end: wkEnd, useFullMonth };
    }

    logBox.innerHTML = '';
    seenCount = 0;
    progressBar.style.width = '0%';
    progressLabel.textContent = '準備中…';
    show(progressCard);

    try {
      const res  = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!res.ok || data.error) {
        errorMsg.textContent = data.error || '啟動失敗';
        show(errorCard);
        return;
      }
      currentJobId = data.jobId;
      appendLog(`[INFO] 工作已啟動，請稍候（約 2 分鐘）…`);
      pollTimer = setInterval(pollStatus, 2000);
    } catch (e) {
      errorMsg.textContent = '無法連線至本機伺服器：' + e.message;
      show(errorCard);
    }
  });

  resetBtn.addEventListener('click', () => {
    stopPoll();
    currentJobId = null;
    seenCount = 0;
    show(configCard);
  });

  retryBtn.addEventListener('click', () => {
    stopPoll();
    currentJobId = null;
    seenCount = 0;
    show(configCard);
  });

  const trafficStatus = $('trafficStatus');
  const N1_STORES = 5;   // 北一區有計數器的門市數（羅東用公式不計）

  async function refreshTrafficStatus() {
    if (!trafficStatus) return;
    try {
      const res = await fetch('/api/traffic-status');
      const d = await res.json();
      if (d.storeCount > 0) {
        const ok = d.storeCount >= N1_STORES;
        trafficStatus.textContent =
          `🟢 人流已更新：${d.storeCount}/${N1_STORES} 店 · 最新 ${d.latest}` +
          (d.updated ? ` · 推送於 ${d.updated}` : '') +
          (ok ? '' : '（門市數不足，請確認 ShopperTrak 已登入並重整）');
      } else {
        trafficStatus.textContent =
          '⚪ 尚未收到人流資料 — 請開啟登入的 ShopperTrak 網頁，插件會自動背景推送（人流欄將留空）';
      }
    } catch (e) {
      trafficStatus.textContent = '人流狀態：無法查詢';
    }
  }

  loadDefaultDate();
  refreshTrafficStatus();
  setInterval(refreshTrafficStatus, 15000);
})();

// ===== 状態管理 =====
let venues = [];

const FORM_FIELDS = ['purpose', 'areas', 'budget', 'capacity', 'fullday-hours', 'extra-keywords', 'search-keywords', 'period-start', 'period-end', 'other-conditions', 'request-delay', 'max-fetch'];

const DEFAULT_SEARCH_KEYWORDS = [
  '貸会議室', '貸会議室 個室', '貸会議室 格安', '貸会議室 少人数',
  'レンタルスペース', 'レンタルスペース 個室',
  '会議室 時間貸し', '会議室 レンタル',
  'コワーキングスペース', 'コワーキング 個室',
  '公民館', '市民センター', '商工会議所', '図書館 会議室',
  'ホテル 会議室', 'TKP', 'リージャス',
  'レンタルオフィス', 'シェアオフィス',
  'インスタベース', 'スペースマーケット', 'スペイシー',
  '多目的室', '研修室', '面接 貸室',
];

function getSearchKeywords() {
  const el = document.getElementById('search-keywords');
  if (!el) return DEFAULT_SEARCH_KEYWORDS;
  const lines = el.value.trim().split('\n').map(l => l.trim()).filter(l => l);
  return lines.length > 0 ? lines : DEFAULT_SEARCH_KEYWORDS;
}

// ===== 初期化 =====
document.addEventListener('DOMContentLoaded', async () => {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  document.getElementById('btn-search').addEventListener('click', startPhase1);
  document.getElementById('btn-fetch-prices').addEventListener('click', startPhase2);
  document.getElementById('btn-stop').addEventListener('click', stopSearch);
  document.getElementById('btn-export').addEventListener('click', exportExcel);
  document.getElementById('btn-clear').addEventListener('click', clearAllResults);
  document.getElementById('btn-save-settings').addEventListener('click', saveFormData);
  document.getElementById('filter-area').addEventListener('change', renderResults);

  document.getElementById('btn-top').addEventListener('click', () => {
    const active = document.querySelector('.tab-content.active');
    if (active) active.scrollTo({ top: 0, behavior: 'smooth' });
  });
  document.getElementById('btn-bottom').addEventListener('click', () => {
    const active = document.querySelector('.tab-content.active');
    if (active) active.scrollTo({ top: active.scrollHeight, behavior: 'smooth' });
  });

  const stored = await chrome.storage.local.get(['venues', 'formData', 'activeTab', 'searchState']);
  if (stored.venues) venues = stored.venues;
  if (stored.formData) restoreFormData(stored.formData);
  if (stored.activeTab) switchTab(stored.activeTab);

  const kwEl = document.getElementById('search-keywords');
  if (kwEl && !kwEl.value.trim()) kwEl.value = DEFAULT_SEARCH_KEYWORDS.join('\n');

  updateResultCount();
  renderResults();

  // バックグラウンドで検索中ならUIに反映
  if (stored.searchState) {
    applySearchState(stored.searchState);
  }

  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', saveFormData);
      el.addEventListener('change', saveFormData);
    }
  });

  // ストレージ変更を監視（venues + searchState）
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'local') return;
    if (changes.venues) {
      venues = changes.venues.newValue || [];
      updateResultCount();
      renderResults();
    }
    if (changes.searchState) {
      applySearchState(changes.searchState.newValue);
    }
  });
});

// ===== バックグラウンドの検索状態をUIに反映 =====
function applySearchState(state) {
  if (!state) return;
  const btnSearch = document.getElementById('btn-search');
  const btnPrice = document.getElementById('btn-fetch-prices');
  const btnStop = document.getElementById('btn-stop');
  const statusArea = document.getElementById('search-status');

  if (state.running) {
    btnStop.classList.remove('hidden');
    if (state.phase === 'phase1') {
      btnSearch.disabled = true;
      btnSearch.innerHTML = `<span class="spinner"></span>${state.done}/${state.total}`;
      btnPrice.disabled = true;
    } else if (state.phase === 'phase2') {
      btnPrice.disabled = true;
      btnPrice.innerHTML = `<span class="spinner"></span>${state.done}/${state.total}`;
      btnSearch.disabled = true;
    }

    // 進捗表示：最新の状況をシンプルに
    const lastLog = state.log && state.log.length > 0 ? state.log[state.log.length - 1] : null;
    const phaseName = state.phase === 'phase1' ? '施設検索' : '料金取得';
    let html = `<div class="status-running">`;
    html += `<div class="status-running-title"><span class="spinner-dark"></span>${phaseName}中... ${state.done}/${state.total}</div>`;
    if (lastLog) {
      html += `<div class="status-running-detail ${lastLog.type}">${lastLog.text}</div>`;
    }
    html += `<div class="status-running-sub">ポップアップを閉じても検索は続きます</div>`;
    html += `</div>`;
    statusArea.innerHTML = html;
  } else {
    btnStop.classList.add('hidden');
    btnSearch.disabled = false;
    btnSearch.innerHTML = '施設を検索';
    btnPrice.disabled = false;
    btnPrice.innerHTML = '料金を取得';

    // 完了時：最後のログだけ表示
    const lastLog = state.log && state.log.length > 0 ? state.log[state.log.length - 1] : null;
    if (lastLog) {
      statusArea.innerHTML = `<div class="status-line ${lastLog.type}">${lastLog.text}</div>`;
    }
  }
}

// ===== フォーム =====
function saveFormData() {
  const formData = {};
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) formData[id] = el.value;
  });
  chrome.storage.local.set({ formData });
}
function restoreFormData(fd) {
  Object.keys(fd).forEach(id => {
    const el = document.getElementById(id);
    if (el && fd[id] !== undefined) el.value = fd[id];
  });
}

// ===== タブ切り替え =====
function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  document.querySelector(`[data-tab="${tabId}"]`).classList.add('active');
  document.getElementById(`tab-${tabId}`).classList.add('active');
  if (tabId === 'results') renderResults();
  chrome.storage.local.set({ activeTab: tabId });
}

// =========================================================
// 第1段階：バックグラウンドに検索指示を送る
// =========================================================
function startPhase1() {
  const areasText = document.getElementById('areas').value.trim();
  const statusArea = document.getElementById('search-status');

  if (!areasText) {
    statusArea.innerHTML = '<div class="status-line error">基準地点を入力してください</div>';
    return;
  }

  saveFormData();
  const areas = areasText.split('\n').map(a => a.trim()).filter(a => a);
  const extra = document.getElementById('extra-keywords').value.trim();
  const keywords = getSearchKeywords();
  const delay = (parseInt(document.getElementById('request-delay').value) || 3) * 1000;

  chrome.runtime.sendMessage({
    type: 'start-phase1',
    params: { areas, keywords, extra, delay }
  }, (res) => {
    if (res?.error) {
      statusArea.innerHTML = `<div class="status-line error">${res.error}</div>`;
    }
  });
}

// =========================================================
// 第2段階：バックグラウンドに料金取得指示を送る
// =========================================================
function startPhase2() {
  saveFormData();
  const delay = (parseInt(document.getElementById('request-delay').value) || 3) * 1000;
  const maxFetch = parseInt(document.getElementById('max-fetch').value) || 20;

  chrome.runtime.sendMessage({
    type: 'start-phase2',
    params: { delay, maxFetch }
  }, (res) => {
    if (res?.error) {
      const statusArea = document.getElementById('search-status');
      statusArea.innerHTML = `<div class="status-line error">${res.error}</div>`;
    }
  });
}

// ===== 中止 =====
function stopSearch() {
  chrome.runtime.sendMessage({ type: 'stop-search' });
}

// ===== データ管理 =====
function removeVenue(id) {
  venues = venues.filter(v => v.id !== id);
  saveVenues(); updateResultCount(); renderResults();
}
function clearAllResults() {
  if (!confirm('すべての結果を削除しますか？')) return;
  venues = []; saveVenues(); updateResultCount(); renderResults();
}
function saveVenues() { chrome.storage.local.set({ venues }); }
function updateResultCount() { document.getElementById('result-count').textContent = venues.length; }

// ===== 結果表示 =====
function renderResults() {
  const list = document.getElementById('results-list');
  const filterArea = document.getElementById('filter-area').value;

  const areas = [...new Set(venues.map(v => v.area).filter(a => a))];
  const sel = document.getElementById('filter-area');
  const cur = sel.value;
  sel.innerHTML = '<option value="">全エリア</option>';
  areas.forEach(a => {
    const o = document.createElement('option');
    o.value = a; o.textContent = a;
    if (a === cur) o.selected = true;
    sel.appendChild(o);
  });

  const filtered = filterArea ? venues.filter(v => v.area === filterArea) : [...venues];

  if (filtered.length === 0) {
    list.innerHTML = '<div class="empty-state"><p>まだ施設データがありません</p></div>';
    return;
  }

  filtered.sort((a, b) => {
    const da = a.distanceKm ?? 9999, db = b.distanceKm ?? 9999;
    if (da !== db) return da - db;
    return (a.hourlyPrice || 999999) - (b.hourlyPrice || 999999);
  });

  list.innerHTML = filtered.map(v => {
    const dist = v.distanceKm != null ? `${v.distanceKm}km` : '';
    const price = v.hourlyPrice ? `¥${v.hourlyPrice.toLocaleString()}/h` : '料金不明';
    const maps = v.address
      ? `https://www.google.com/maps/dir/${encodeURIComponent(v.area)}/${encodeURIComponent(v.address)}`
      : `https://www.google.com/maps/search/${encodeURIComponent(v.name)}`;
    return `
    <div class="venue-card ${v.hourlyPrice ? '' : 'no-price'}">
      <div class="venue-name">${esc(v.name)}</div>
      <div class="venue-meta">
        ${dist ? `<span class="venue-distance">${dist}</span>` : ''}
        <span class="venue-price">${price}</span>
        <span class="venue-platform">${esc(v.platform || '')}</span>
        ${v.transferPayment ? `<span>振込:${esc(v.transferPayment)}</span>` : ''}
      </div>
      <div class="venue-info">
        <span>${esc(v.address || '')}</span>
        <span>${esc(v.capacity || '')}</span>
      </div>
      ${v.priceDetail ? `<div class="venue-price-detail">${esc(v.priceDetail)}</div>` : ''}
      <div class="venue-actions">
        ${v.officialUrl ? `<button class="btn-link" data-url="${escA(v.officialUrl)}">サイト</button>` : ''}
        <button class="btn-link" data-url="${escA(maps)}">地図</button>
        <button class="btn-remove" data-id="${v.id}">削除</button>
      </div>
    </div>`;
  }).join('');

  list.querySelectorAll('.btn-link').forEach(b => {
    b.addEventListener('click', () => chrome.tabs.create({ url: b.dataset.url }));
  });
  list.querySelectorAll('.btn-remove').forEach(b => {
    b.addEventListener('click', () => removeVenue(parseFloat(b.dataset.id)));
  });
}

// ===== Excel出力 =====
function exportExcel() {
  if (venues.length === 0) { alert('出力するデータがありません'); return; }
  const fh = parseInt(document.getElementById('fullday-hours')?.value) || 8;
  const wb = XLSX.utils.book_new();

  const all = venues.map(v => ({
    '基準地点': v.area||'', '距離（km）': v.distanceKm??'', '施設名': v.name||'',
    '施設住所': v.address||'', '公式URL': v.officialUrl||'', '予約URL': v.bookingUrl||'',
    'Google Map経路': v.address&&v.area ? `https://www.google.com/maps/dir/${encodeURIComponent(v.area)}/${encodeURIComponent(v.address)}` : '',
    '担当者連絡先': v.contactInfo||'', '商用利用可否': v.commercialUse||'要確認',
    '1時間あたり料金（円）': v.hourlyPrice||'',
    '終日利用概算（税込）': v.hourlyPrice ? v.hourlyPrice*fh : '',
    '料金詳細': v.priceDetail||'', '振込対応': v.transferPayment||'要確認',
    '支払方法の詳細': v.paymentDetail||'', 'プラットフォーム': v.platform||'',
    '収容人数': v.capacity||'', '設備': v.equipment||'',
    '予約方法': v.bookingMethod||'', '備考': v.note||''
  }));
  const ws1 = XLSX.utils.json_to_sheet(all);
  ws1['!cols'] = [{wch:20},{wch:10},{wch:30},{wch:35},{wch:40},{wch:40},{wch:50},{wch:15},{wch:10},{wch:15},{wch:18},{wch:45},{wch:10},{wch:25},{wch:15},{wch:10},{wch:25},{wch:15},{wch:30}];
  ws1['!autofilter'] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(ws1['!ref'])) };
  XLSX.utils.book_append_sheet(wb, ws1, '全施設一覧');

  const areas = [...new Set(venues.map(v=>v.area).filter(a=>a))];
  for (const area of areas) {
    const av = venues.filter(v=>v.area===area).sort((a,b)=>(a.hourlyPrice||999999)-(b.hourlyPrice||999999));
    const ad = av.map(v=>({
      '施設名':v.name||'','距離（km）':v.distanceKm??'','1時間あたり料金（円）':v.hourlyPrice||'',
      '終日利用概算（税込）':v.hourlyPrice?v.hourlyPrice*fh:'','料金詳細':v.priceDetail||'',
      '振込対応':v.transferPayment||'要確認','商用利用可否':v.commercialUse||'要確認',
      '施設住所':v.address||'','公式URL':v.officialUrl||'',
      'Google Map経路':v.address?`https://www.google.com/maps/dir/${encodeURIComponent(area)}/${encodeURIComponent(v.address)}`:'',
      '収容人数':v.capacity||'','設備':v.equipment||'','備考':v.note||''
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ad), area.substring(0,28));
  }

  const pd = [
    {'プラットフォーム':'スペースマーケット','振込対応':'○','条件':'法人Paid登録（審査2-3営業日）','支払サイクル':'月末締め→翌月末払い'},
    {'プラットフォーム':'インスタベース','振込対応':'○','条件':'法人Paid登録（審査即時～3営業日）','支払サイクル':'月末締め→翌月末払い'},
    {'プラットフォーム':'upnow','振込対応':'○','条件':'請求書・領収書発行可','支払サイクル':'都度'},
    {'プラットフォーム':'日本会議室','振込対応':'○','条件':'法人利用対応','支払サイクル':'都度請求書'},
    {'プラットフォーム':'TKP','振込対応':'○','条件':'法人請求書払い標準対応','支払サイクル':'都度請求書'},
    {'プラットフォーム':'公共施設','振込対応':'△','条件':'窓口現金精算が基本','支払サイクル':'利用当日精算'},
    {'プラットフォーム':'個人運営スペース','振込対応':'要確認','条件':'施設による','支払サイクル':'施設による'}
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pd), '支払方法まとめ');

  const ai = [];
  venues.filter(v=>!v.hourlyPrice).forEach(v=>{ai.push({'優先度':'高','アクション':'料金を確認','対象施設':v.name,'連絡先':v.contactInfo||v.officialUrl||'','備考':''})});
  venues.filter(v=>v.commercialUse==='要確認').forEach(v=>{ai.push({'優先度':'高','アクション':'商用利用可否を確認','対象施設':v.name,'連絡先':v.contactInfo||v.officialUrl||'','備考':''})});
  venues.filter(v=>/公民館|市民|図書館/.test(v.name)).forEach(v=>{ai.push({'優先度':'高','アクション':'面接利用が商行為に該当するか電話確認','対象施設':v.name,'連絡先':v.contactInfo||'','備考':'商行為該当の場合は料金2～3倍'})});
  if(ai.length===0)ai.push({'優先度':'','アクション':'アクション項目なし','対象施設':'','連絡先':'','備考':''});
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ai), '次のアクション');

  const d = new Date();
  const dateStr = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
  const purpose = document.getElementById('purpose')?.value || '会場';
  // エリアから地名を取得（最初の1件の都道府県+市区町村）
  const areasText = document.getElementById('areas')?.value?.trim() || '';
  const firstArea = areasText.split('\n').map(a => a.trim()).filter(a => a)[0] || '';
  // 「東京都新宿区...」→「東京都新宿区」、「岡山県岡山市...」→「岡山県岡山市」
  const areaMatch = firstArea.match(/((?:東京都|北海道|(?:京都|大阪)府|.{2,3}県)(?:[^区市町村]{1,5}[区市町村])?)/)
    || firstArea.match(/(.{2,10}[区市町村駅])/);
  const areaName = areaMatch ? areaMatch[1] : firstArea.substring(0, 10);
  const fileName = `${dateStr}_${areaName}${purpose}`;
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

// ===== ユーティリティ =====
function esc(s){const d=document.createElement('div');d.textContent=s||'';return d.innerHTML}
function escA(s){return(s||'').replace(/'/g,"\\'").replace(/"/g,'&quot;')}
function addStatus(c,t,type){const l=document.createElement('div');l.className=`status-line ${type}`;l.textContent=t;c.appendChild(l);c.scrollTop=c.scrollHeight}

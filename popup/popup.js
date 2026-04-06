// ===== 状態管理 =====
let venues = [];

const FORM_FIELDS = ['purpose', 'areas', 'budget', 'capacity', 'fullday-hours', 'extra-keywords', 'search-keywords', 'period-start', 'period-end', 'other-conditions'];

const DEFAULT_SEARCH_KEYWORDS = [
  '貸会議室',
  '貸会議室 個室',
  '貸会議室 格安',
  '貸会議室 少人数',
  'レンタルスペース',
  'レンタルスペース 個室',
  '会議室 時間貸し',
  '会議室 レンタル',
  'コワーキングスペース',
  'コワーキング 個室',
  '公民館',
  '市民センター',
  '商工会議所',
  '図書館 会議室',
  'ホテル 会議室',
  'TKP',
  'リージャス',
  'レンタルオフィス',
  'シェアオフィス',
  'インスタベース',
  'スペースマーケット',
  'スペイシー',
  '多目的室',
  '研修室',
  '面接 貸室',
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

  document.getElementById('btn-search').addEventListener('click', startBulkSearch);
  document.getElementById('btn-fetch-prices').addEventListener('click', fetchPricesForAll);
  document.getElementById('btn-export').addEventListener('click', exportExcel);
  document.getElementById('btn-clear').addEventListener('click', clearAllResults);
  document.getElementById('btn-save-settings').addEventListener('click', saveFormData);
  document.getElementById('filter-area').addEventListener('change', renderResults);

  const stored = await chrome.storage.local.get(['venues', 'formData', 'activeTab']);
  if (stored.venues) venues = stored.venues;
  if (stored.formData) restoreFormData(stored.formData);
  if (stored.activeTab) switchTab(stored.activeTab);

  const kwEl = document.getElementById('search-keywords');
  if (kwEl && !kwEl.value.trim()) kwEl.value = DEFAULT_SEARCH_KEYWORDS.join('\n');

  updateResultCount();
  renderResults();

  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', saveFormData);
      el.addEventListener('change', saveFormData);
    }
  });

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'local' && changes.venues) {
      venues = changes.venues.newValue || [];
      updateResultCount();
      renderResults();
    }
  });
});

// ===== フォーム =====
function saveFormData() {
  const formData = {};
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) formData[id] = el.value;
  });
  chrome.storage.local.set({ formData });
}

function restoreFormData(formData) {
  Object.keys(formData).forEach(id => {
    const el = document.getElementById(id);
    if (el && formData[id] !== undefined) el.value = formData[id];
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
// 第1段階：Google検索で施設URLを大量収集
// =========================================================
async function startBulkSearch() {
  const btn = document.getElementById('btn-search');
  const statusArea = document.getElementById('search-status');
  const areasText = document.getElementById('areas').value.trim();

  if (!areasText) {
    statusArea.innerHTML = '<div class="status-line error">基準地点を入力してください</div>';
    return;
  }

  const areas = areasText.split('\n').map(a => a.trim()).filter(a => a);
  const extra = document.getElementById('extra-keywords').value.trim();
  const keywords = getSearchKeywords();

  saveFormData();
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>検索中...';
  statusArea.innerHTML = '';

  let openedCount = 0;
  for (const area of areas) {
    for (const keyword of keywords) {
      const query = extra ? `${area} ${keyword} ${extra}` : `${area} ${keyword}`;
      const url = `https://www.google.co.jp/search?q=${encodeURIComponent(query)}&num=20`;
      addStatus(statusArea, `${area} →「${keyword}」`, 'info');
      try {
        await chrome.tabs.create({ url, active: false });
        openedCount++;
        if (openedCount % 5 === 0) await sleep(800);
      } catch (e) {
        addStatus(statusArea, `タブ作成失敗: ${e.message}`, 'error');
      }
    }

    // プラットフォーム直接検索も併用
    const kw = encodeURIComponent(area);
    const cap = document.getElementById('capacity').value || 2;
    const platforms = [
      { label: 'インスタベース', url: `https://www.instabase.jp/search?keyword=${kw}&pax=${cap}&category=meetingroom` },
      { label: 'スペースマーケット', url: `https://www.spacemarket.com/spaces?keyword=${kw}&people=${cap}&types%5B%5D=meeting_room` },
      { label: 'スペイシー', url: `https://www.spacee.jp/listings?location=${kw}&capacity=${cap}` },
    ];
    for (const p of platforms) {
      try {
        await chrome.tabs.create({ url: p.url, active: false });
        openedCount++;
      } catch (e) { /* skip */ }
    }
  }

  addStatus(statusArea, `${openedCount}件のタブを開きました。自動抽出中...`, 'success');
  addStatus(statusArea, '抽出完了後「第2段階：料金を一括取得」で各施設の料金を取得してください', 'info');

  btn.disabled = false;
  btn.innerHTML = '第1段階：施設を一括検索';
}

// =========================================================
// 第2段階：各施設URLを開いて汎用料金抽出
// =========================================================
async function fetchPricesForAll() {
  const btn = document.getElementById('btn-fetch-prices');
  const statusArea = document.getElementById('search-status');

  const targets = venues.filter(v => !v.hourlyPrice && v.officialUrl);
  if (targets.length === 0) {
    addStatus(statusArea, '料金未取得の施設がありません', 'success');
    return;
  }

  btn.disabled = true;
  statusArea.innerHTML = '';
  addStatus(statusArea, `${targets.length}件の施設から料金を取得します`, 'info');

  let doneCount = 0;
  let successCount = 0;

  for (const venue of targets) {
    doneCount++;
    btn.innerHTML = `<span class="spinner"></span>${doneCount}/${targets.length}`;
    addStatus(statusArea, `[${doneCount}/${targets.length}] ${venue.name.substring(0, 30)}`, 'info');

    try {
      const tab = await chrome.tabs.create({ url: venue.officialUrl, active: false });
      await waitForTabLoad(tab.id, 12000);

      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: universalPriceExtractor
      });

      if (results?.[0]?.result) {
        const d = results[0].result;
        if (d.hourlyPrice || d.priceDetail) {
          venue.hourlyPrice = d.hourlyPrice || venue.hourlyPrice;
          venue.priceDetail = d.priceDetail || venue.priceDetail;
          venue.address = d.address || venue.address;
          venue.contactInfo = d.phone || venue.contactInfo;
          venue.capacity = d.capacity || venue.capacity;
          successCount++;
          const priceStr = d.hourlyPrice ? `¥${d.hourlyPrice}/h` : '';
          addStatus(statusArea, `  → ${priceStr} ${(d.priceDetail || '').substring(0, 60)}`, 'success');
        } else {
          addStatus(statusArea, `  → 料金情報なし`, 'error');
        }
      }

      try { await chrome.tabs.remove(tab.id); } catch (e) { }
      await sleep(1200);
    } catch (e) {
      addStatus(statusArea, `  → エラー: ${e.message}`, 'error');
    }
  }

  saveVenues();
  updateResultCount();
  renderResults();

  btn.disabled = false;
  btn.innerHTML = '第2段階：料金を一括取得';
  addStatus(statusArea, `完了: ${successCount}/${targets.length}件で料金取得`, 'success');
}

// タブの読み込み完了を待つ
function waitForTabLoad(tabId, timeout) {
  return new Promise((resolve) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      resolve();
    }, timeout);

    function listener(id, info) {
      if (id === tabId && info.status === 'complete') {
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(listener);
        setTimeout(resolve, 1500); // レンダリング待ち
      }
    }
    chrome.tabs.onUpdated.addListener(listener);
  });
}

// =========================================================
// 汎用料金抽出（どんなサイトでも動く。DOM構造に依存しない）
// =========================================================
function universalPriceExtractor() {
  const text = document.body?.innerText || '';

  // --- 時間単価の抽出 ---
  const hourlyPatterns = [
    // 「1,100円/時間」「1,100円/h」「1,100円／1時間」
    /(\d{1,3}[,，]\d{3})\s*円\s*[/／~〜]\s*(?:1\s*)?(?:時間|[hH])/,
    // 「1時間あたり1,100円」「1時間 1,100円」
    /(?:1\s*)?(?:時間)\s*[あたり:：]*\s*[^\d]{0,5}?(\d{1,3}[,，]\d{3})\s*円/,
    // 「¥1,100/h」
    /[¥￥]\s*(\d{1,3}[,，]\d{3})\s*[/／]\s*(?:時間|[hH])/,
    // 「1100円/時間」（カンマなし）
    /(\d{3,5})\s*円\s*[/／~〜]\s*(?:1\s*)?(?:時間|[hH])/,
    // 「1時間あたり1100円」（カンマなし）
    /(?:1\s*)?(?:時間)\s*[あたり:：]*\s*[^\d]{0,5}?(\d{3,5})\s*円/,
    // 「料金 1,100円」「利用料 1,100円」（最終手段）
    /(?:料金|価格|利用料|使用料|単価)[^\d\n]{0,15}?(\d{1,3}[,，]?\d{3})\s*円/,
  ];

  let hourlyPrice = null;
  for (const pat of hourlyPatterns) {
    const m = text.match(pat);
    if (m) {
      const p = parseInt(m[1].replace(/[,，]/g, ''));
      if (p > 50 && p < 50000) { hourlyPrice = p; break; }
    }
  }

  // 分単位料金→時間換算
  if (!hourlyPrice) {
    const m = text.match(/(\d{2,5})\s*円\s*[/／~〜]\s*(\d{1,2})\s*分/)
           || text.match(/(\d{1,2})\s*分\s*[あたり]*\s*(\d{2,5})\s*円/);
    if (m) {
      let price, mins;
      if (parseInt(m[1]) > 100) { price = parseInt(m[1]); mins = parseInt(m[2]); }
      else { mins = parseInt(m[1]); price = parseInt(m[2]); }
      if (price > 0 && mins > 0 && mins <= 60) {
        const calc = Math.round(price * (60 / mins));
        if (calc > 50 && calc < 50000) hourlyPrice = calc;
      }
    }
  }

  // --- 料金っぽいテキストを全て収集 ---
  const priceTexts = [];
  const seen = new Set();
  // 「○○円」を含む行を拾う
  const lines = text.split('\n');
  for (const line of lines) {
    if (/\d{3,7}円/.test(line) && line.length < 80) {
      const clean = line.trim();
      if (!seen.has(clean)) {
        seen.add(clean);
        priceTexts.push(clean);
      }
      if (priceTexts.length >= 10) break;
    }
  }

  // --- 住所 ---
  const addrMatch = text.match(/〒\d{3}-?\d{4}[^\n]{3,40}/)
    || text.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県)[^\n]{2,15}[区市町村][^\n]{1,20}[\d\-]+/);
  const address = addrMatch ? addrMatch[0].trim().substring(0, 60) : '';

  // --- 電話番号 ---
  const phoneMatch = text.match(/(?:TEL|電話|tel|Tel|☎)[^\d]{0,5}(0\d{1,4}[-ー\s]?\d{1,4}[-ー\s]?\d{3,4})/)
    || text.match(/(0\d{1,4}-\d{1,4}-\d{3,4})/);
  const phone = phoneMatch ? (phoneMatch[1] || phoneMatch[0]) : '';

  // --- 収容人数 ---
  const capMatch = text.match(/(?:定員|収容|着席)\s*[：:]*\s*(\d{1,4})\s*(?:名|人)/)
    || text.match(/(\d{1,3})\s*(?:名|人)\s*(?:まで|収容|着席)/);
  const capacity = capMatch ? capMatch[1] + '名' : '';

  return {
    hourlyPrice,
    priceDetail: priceTexts.join(' | '),
    address,
    phone,
    capacity
  };
}

// ===== データ管理 =====
function addVenues(newVenues) {
  for (const v of newVenues) {
    const isDup = venues.some(ex => ex.officialUrl === v.officialUrl);
    if (!isDup) {
      v.id = Date.now() + Math.random();
      venues.push(v);
    }
  }
  saveVenues();
  updateResultCount();
  renderResults();
}

function removeVenue(id) {
  venues = venues.filter(v => v.id !== id);
  saveVenues();
  updateResultCount();
  renderResults();
}

function clearAllResults() {
  if (!confirm('すべての結果を削除しますか？')) return;
  venues = [];
  saveVenues();
  updateResultCount();
  renderResults();
}

function saveVenues() {
  chrome.storage.local.set({ venues });
}

function updateResultCount() {
  document.getElementById('result-count').textContent = venues.length;
}

// ===== 結果表示 =====
function renderResults() {
  const list = document.getElementById('results-list');
  const filterArea = document.getElementById('filter-area').value;

  const areas = [...new Set(venues.map(v => v.area).filter(a => a))];
  const filterSelect = document.getElementById('filter-area');
  const currentFilter = filterSelect.value;
  filterSelect.innerHTML = '<option value="">全エリア</option>';
  areas.forEach(a => {
    const opt = document.createElement('option');
    opt.value = a;
    opt.textContent = a;
    if (a === currentFilter) opt.selected = true;
    filterSelect.appendChild(opt);
  });

  const filtered = filterArea ? venues.filter(v => v.area === filterArea) : [...venues];

  if (filtered.length === 0) {
    list.innerHTML = '<div class="empty-state"><p>まだ施設データがありません</p><p>「検索」タブで調査を開始してください</p></div>';
    return;
  }

  filtered.sort((a, b) => {
    const da = a.distanceKm ?? 9999;
    const db = b.distanceKm ?? 9999;
    if (da !== db) return da - db;
    return (a.hourlyPrice || 999999) - (b.hourlyPrice || 999999);
  });

  list.innerHTML = filtered.map(v => {
    const distText = v.distanceKm != null ? `${v.distanceKm}km` : '';
    const priceText = v.hourlyPrice ? `¥${v.hourlyPrice.toLocaleString()}/h` : '料金不明';
    const mapsUrl = v.address
      ? `https://www.google.com/maps/dir/${encodeURIComponent(v.area)}/${encodeURIComponent(v.address)}`
      : `https://www.google.com/maps/search/${encodeURIComponent(v.name)}`;

    return `
    <div class="venue-card ${v.hourlyPrice ? '' : 'no-price'}">
      <div class="venue-name">${escHtml(v.name)}</div>
      <div class="venue-meta">
        ${distText ? `<span class="venue-distance">${distText}</span>` : ''}
        <span class="venue-price">${priceText}</span>
        <span class="venue-platform">${escHtml(v.platform || '')}</span>
        ${v.transferPayment ? `<span>振込:${escHtml(v.transferPayment)}</span>` : ''}
      </div>
      <div class="venue-info">
        <span>${escHtml(v.address || '')}</span>
        <span>${escHtml(v.capacity || '')}</span>
      </div>
      ${v.priceDetail ? `<div class="venue-price-detail">${escHtml(v.priceDetail)}</div>` : ''}
      <div class="venue-actions">
        ${v.officialUrl ? `<button onclick="window.open('${escAttr(v.officialUrl)}')">サイト</button>` : ''}
        <button onclick="window.open('${escAttr(mapsUrl)}')">地図</button>
        <button class="btn-remove" data-id="${v.id}">削除</button>
      </div>
    </div>`;
  }).join('');

  list.querySelectorAll('.btn-remove').forEach(btn => {
    btn.addEventListener('click', () => removeVenue(parseFloat(btn.dataset.id)));
  });
}

// ===== Excel出力 =====
function exportExcel() {
  if (venues.length === 0) { alert('出力するデータがありません'); return; }

  const fulldayHours = parseInt(document.getElementById('fullday-hours')?.value) || 8;
  const wb = XLSX.utils.book_new();

  const allData = venues.map(v => ({
    '基準地点': v.area || '',
    '距離（km）': v.distanceKm ?? '',
    '施設名': v.name || '',
    '施設住所': v.address || '',
    '公式URL': v.officialUrl || '',
    '予約URL': v.bookingUrl || '',
    'Google Map経路': v.address && v.area
      ? `https://www.google.com/maps/dir/${encodeURIComponent(v.area)}/${encodeURIComponent(v.address)}` : '',
    '担当者連絡先': v.contactInfo || '',
    '商用利用可否': v.commercialUse || '要確認',
    '1時間あたり料金（円）': v.hourlyPrice || '',
    '終日利用概算（税込）': v.hourlyPrice ? v.hourlyPrice * fulldayHours : '',
    '料金詳細': v.priceDetail || '',
    '振込対応': v.transferPayment || '要確認',
    '支払方法の詳細': v.paymentDetail || '',
    'プラットフォーム': v.platform || '',
    '収容人数': v.capacity || '',
    '設備': v.equipment || '',
    '予約方法': v.bookingMethod || '',
    '備考': v.note || ''
  }));

  const ws1 = XLSX.utils.json_to_sheet(allData);
  ws1['!cols'] = [
    { wch: 20 }, { wch: 10 }, { wch: 30 }, { wch: 35 },
    { wch: 40 }, { wch: 40 }, { wch: 50 },
    { wch: 15 }, { wch: 10 },
    { wch: 15 }, { wch: 18 }, { wch: 45 },
    { wch: 10 }, { wch: 25 }, { wch: 15 },
    { wch: 10 }, { wch: 25 }, { wch: 15 }, { wch: 30 }
  ];
  const range1 = XLSX.utils.decode_range(ws1['!ref']);
  ws1['!autofilter'] = { ref: XLSX.utils.encode_range(range1) };
  XLSX.utils.book_append_sheet(wb, ws1, '全施設一覧');

  // エリア別シート
  const areas = [...new Set(venues.map(v => v.area).filter(a => a))];
  for (const area of areas) {
    const av = venues.filter(v => v.area === area).sort((a, b) => (a.hourlyPrice || 999999) - (b.hourlyPrice || 999999));
    const ad = av.map(v => ({
      '施設名': v.name || '', '距離（km）': v.distanceKm ?? '',
      '1時間あたり料金（円）': v.hourlyPrice || '',
      '終日利用概算（税込）': v.hourlyPrice ? v.hourlyPrice * fulldayHours : '',
      '料金詳細': v.priceDetail || '', '振込対応': v.transferPayment || '要確認',
      '商用利用可否': v.commercialUse || '要確認', '施設住所': v.address || '',
      '公式URL': v.officialUrl || '',
      'Google Map経路': v.address ? `https://www.google.com/maps/dir/${encodeURIComponent(area)}/${encodeURIComponent(v.address)}` : '',
      '収容人数': v.capacity || '', '設備': v.equipment || '', '備考': v.note || ''
    }));
    const ws = XLSX.utils.json_to_sheet(ad);
    ws['!cols'] = [
      { wch: 30 }, { wch: 10 }, { wch: 15 }, { wch: 18 }, { wch: 45 },
      { wch: 10 }, { wch: 10 }, { wch: 35 }, { wch: 40 }, { wch: 50 },
      { wch: 10 }, { wch: 25 }, { wch: 30 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, area.substring(0, 28));
  }

  // 支払方法まとめ
  const pd = [
    { 'プラットフォーム': 'スペースマーケット', '振込対応': '○', '条件': '法人Paid登録（審査2-3営業日）', '支払サイクル': '月末締め→翌月末払い' },
    { 'プラットフォーム': 'インスタベース', '振込対応': '○', '条件': '法人Paid登録（審査即時～3営業日）', '支払サイクル': '月末締め→翌月末払い' },
    { 'プラットフォーム': 'upnow', '振込対応': '○', '条件': '請求書・領収書発行可', '支払サイクル': '都度' },
    { 'プラットフォーム': '日本会議室', '振込対応': '○', '条件': '法人利用対応', '支払サイクル': '都度請求書' },
    { 'プラットフォーム': 'TKP', '振込対応': '○', '条件': '法人請求書払い標準対応', '支払サイクル': '都度請求書' },
    { 'プラットフォーム': '公共施設', '振込対応': '△', '条件': '窓口現金精算が基本', '支払サイクル': '利用当日精算' },
    { 'プラットフォーム': '個人運営スペース', '振込対応': '要確認', '条件': '施設による', '支払サイクル': '施設による' }
  ];
  const wsPay = XLSX.utils.json_to_sheet(pd);
  wsPay['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 35 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, wsPay, '支払方法まとめ');

  // 次のアクション
  const ai = [];
  venues.filter(v => !v.hourlyPrice).forEach(v => {
    ai.push({ '優先度': '高', 'アクション': '料金を確認', '対象施設': v.name, '連絡先': v.contactInfo || v.officialUrl || '', '備考': '' });
  });
  venues.filter(v => v.commercialUse === '要確認').forEach(v => {
    ai.push({ '優先度': '高', 'アクション': '商用利用可否を確認', '対象施設': v.name, '連絡先': v.contactInfo || v.officialUrl || '', '備考': '' });
  });
  venues.filter(v => /公民館|市民|図書館/.test(v.name)).forEach(v => {
    ai.push({ '優先度': '高', 'アクション': '面接利用が商行為に該当するか電話確認', '対象施設': v.name, '連絡先': v.contactInfo || '', '備考': '商行為該当の場合は料金2～3倍' });
  });
  if (ai.length === 0) ai.push({ '優先度': '', 'アクション': 'アクション項目なし', '対象施設': '', '連絡先': '', '備考': '' });
  const wsAct = XLSX.utils.json_to_sheet(ai);
  wsAct['!cols'] = [{ wch: 8 }, { wch: 35 }, { wch: 30 }, { wch: 30 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, wsAct, '次のアクション');

  const now = new Date();
  const ds = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(wb, `会場調査_${ds}.xlsx`);
}

// ===== ユーティリティ =====
function escHtml(s) { const d = document.createElement('div'); d.textContent = s || ''; return d.innerHTML; }
function escAttr(s) { return (s || '').replace(/'/g, "\\'").replace(/"/g, '&quot;'); }
function addStatus(c, t, type) {
  const l = document.createElement('div');
  l.className = `status-line ${type}`;
  l.textContent = t;
  c.appendChild(l);
  c.scrollTop = c.scrollHeight;
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

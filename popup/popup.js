// ===== 状態管理 =====
let venues = [];
let searchSettings = {};

// 自動保存対象のフォームフィールド
const FORM_FIELDS = {
  'purpose': 'value',
  'areas': 'value',
  'budget': 'value',
  'capacity': 'value',
  'fullday-hours': 'value',
  'transport': 'value',
  'extra-keywords': 'value'
};

// ===== 初期化 =====
document.addEventListener('DOMContentLoaded', async () => {
  // タブ切り替え
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // ボタンイベント
  document.getElementById('btn-search').addEventListener('click', startBulkSearch);
  document.getElementById('btn-extract').addEventListener('click', extractCurrentPage);
  document.getElementById('btn-export').addEventListener('click', exportExcel);
  document.getElementById('btn-clear').addEventListener('click', clearAllResults);
  document.getElementById('btn-save-settings').addEventListener('click', saveSettings);
  document.getElementById('filter-area').addEventListener('change', renderResults);

  // 保存データ読み込み（venues + settings + フォーム入力値 + チェックボックス + タブ状態）
  const stored = await chrome.storage.local.get([
    'venues', 'searchSettings', 'formData', 'platformChecks', 'activeTab'
  ]);
  if (stored.venues) venues = stored.venues;
  if (stored.searchSettings) {
    searchSettings = stored.searchSettings;
    restoreSettings();
  }
  if (stored.formData) restoreFormData(stored.formData);
  if (stored.platformChecks) restorePlatformChecks(stored.platformChecks);
  if (stored.activeTab) switchTab(stored.activeTab);

  updateResultCount();
  renderResults();

  // フォーム入力値の自動保存（input/change イベント）
  Object.keys(FORM_FIELDS).forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', saveFormData);
      el.addEventListener('change', saveFormData);
    }
  });

  // プラットフォームチェックボックスの自動保存
  document.querySelectorAll('.checkbox-group input[type="checkbox"][value]').forEach(cb => {
    if (['instabase','spacemarket','spacee','google','municipal'].includes(cb.value)) {
      cb.addEventListener('change', savePlatformChecks);
    }
  });

  // バックグラウンドからのメッセージ受信
  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === 'venues-extracted') {
      addVenues(msg.data, msg.source);
    }
  });

  // ストレージ変更を監視（コンテンツスクリプトが抽出→バックグラウンドが保存→ここで検知）
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'local' && changes.venues) {
      venues = changes.venues.newValue || [];
      updateResultCount();
      renderResults();
    }
  });
});

// ===== フォーム自動保存・復元 =====
function saveFormData() {
  const formData = {};
  Object.keys(FORM_FIELDS).forEach(id => {
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

function savePlatformChecks() {
  const checks = {};
  document.querySelectorAll('.checkbox-group input[type="checkbox"][value]').forEach(cb => {
    if (['instabase','spacemarket','spacee','google','municipal'].includes(cb.value)) {
      checks[cb.value] = cb.checked;
    }
  });
  chrome.storage.local.set({ platformChecks: checks });
}

function restorePlatformChecks(checks) {
  document.querySelectorAll('.checkbox-group input[type="checkbox"][value]').forEach(cb => {
    if (checks[cb.value] !== undefined) {
      cb.checked = checks[cb.value];
    }
  });
}

async function reloadVenuesFromStorage() {
  const stored = await chrome.storage.local.get(['venues']);
  if (stored.venues) venues = stored.venues;
  updateResultCount();
  renderResults();
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

// ===== 一括検索 =====
async function startBulkSearch() {
  const btn = document.getElementById('btn-search');
  const statusArea = document.getElementById('search-status');
  const areasText = document.getElementById('areas').value.trim();

  if (!areasText) {
    statusArea.innerHTML = '<div class="status-line error">エリアを入力してください</div>';
    return;
  }

  const areas = areasText.split('\n').map(a => a.trim()).filter(a => a);
  const platforms = Array.from(document.querySelectorAll('.checkbox-group input[type="checkbox"]:checked'))
    .filter(cb => ['instabase','spacemarket','spacee','google','municipal'].includes(cb.value))
    .map(cb => cb.value);

  if (platforms.length === 0) {
    statusArea.innerHTML = '<div class="status-line error">プラットフォームを1つ以上選択してください</div>';
    return;
  }

  const purpose = document.getElementById('purpose').value;
  const capacity = document.getElementById('capacity').value;
  const extraKeywords = document.getElementById('extra-keywords').value.trim();

  // 検索開始時にフォーム状態を保存
  saveFormData();
  savePlatformChecks();

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>検索中...';
  statusArea.innerHTML = '';

  // 各エリア×各プラットフォームの検索URLを生成して開く
  let openedCount = 0;
  for (const area of areas) {
    for (const platform of platforms) {
      const urls = buildSearchURLs(platform, area, purpose, capacity, extraKeywords);
      for (const urlInfo of urls) {
        addStatus(statusArea, `${area} → ${urlInfo.label} を検索中...`, 'info');
        try {
          await chrome.tabs.create({ url: urlInfo.url, active: false });
          openedCount++;
          // タブ開きすぎ防止
          if (openedCount % 5 === 0) {
            await sleep(1000);
          }
        } catch (e) {
          addStatus(statusArea, `${urlInfo.label} のタブ作成に失敗: ${e.message}`, 'error');
        }
      }
    }
  }

  addStatus(statusArea, `${openedCount}件のタブを開きました。各ページで自動抽出が実行されます。`, 'success');
  addStatus(statusArea, '手動抽出：各ページで「現在のページからデータ抽出」ボタンも使えます', 'info');

  btn.disabled = false;
  btn.innerHTML = '一括検索開始';
}

// ===== 検索URL生成 =====
function buildSearchURLs(platform, area, purpose, capacity, extra) {
  const kw = encodeURIComponent(area);
  const purposeKw = encodeURIComponent(purpose);
  const extraKw = extra ? '+' + encodeURIComponent(extra) : '';
  const cap = capacity || 2;

  switch (platform) {
    case 'instabase':
      return [{
        label: 'インスタベース',
        url: `https://www.instabase.jp/search?keyword=${kw}+${purposeKw}${extraKw}&pax=${cap}&category=meetingroom`
      }];
    case 'spacemarket':
      return [{
        label: 'スペースマーケット',
        url: `https://www.spacemarket.com/spaces?keyword=${kw}+${purposeKw}${extraKw}&people=${cap}&types%5B%5D=meeting_room`
      }];
    case 'spacee':
      return [{
        label: 'スペイシー',
        url: `https://www.spacee.jp/listings?location=${kw}&capacity=${cap}`
      }];
    case 'google':
      return [
        {
          label: 'Google（貸会議室）',
          url: `https://www.google.co.jp/search?q=${kw}+貸会議室+${purposeKw}${extraKw}`
        },
        {
          label: 'Google（レンタルスペース）',
          url: `https://www.google.co.jp/search?q=${kw}+レンタルスペース+${purposeKw}${extraKw}`
        }
      ];
    case 'municipal':
      return [{
        label: 'Google（公民館等）',
        url: `https://www.google.co.jp/search?q=${kw}+公民館+会議室+貸出`
      }];
    default:
      return [];
  }
}

// ===== 現在のページからデータ抽出 =====
async function extractCurrentPage() {
  const statusArea = document.getElementById('search-status');
  addStatus(statusArea, '現在のタブからデータを抽出中...', 'info');

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab) {
      addStatus(statusArea, 'アクティブなタブがありません', 'error');
      return;
    }

    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: extractVenueData,
      args: [tab.url]
    });

    if (results && results[0] && results[0].result) {
      const extracted = results[0].result;
      if (extracted.length > 0) {
        addVenues(extracted, tab.url);
        addStatus(statusArea, `${extracted.length}件の施設を抽出しました`, 'success');
      } else {
        addStatus(statusArea, '施設データが見つかりませんでした。対応していないページか、検索結果がありません。', 'error');
      }
    }
  } catch (e) {
    addStatus(statusArea, `抽出エラー: ${e.message}`, 'error');
  }
}

// ===== ページ内で実行される抽出関数 =====
function extractVenueData(pageUrl) {
  const venues = [];
  const url = pageUrl || location.href;

  // ----- インスタベース -----
  if (url.includes('instabase.jp')) {
    // 検索結果ページ
    const cards = document.querySelectorAll('[class*="SpaceCard"], [class*="spaceCard"], [class*="space-card"], a[href*="/space/"]');
    const seen = new Set();

    // リンクベースの抽出
    document.querySelectorAll('a[href*="/space/"]').forEach(link => {
      const href = link.href;
      const match = href.match(/\/space\/(\d+)/);
      if (!match || seen.has(match[1])) return;
      seen.add(match[1]);

      const card = link.closest('[class*="Card"]') || link.closest('li') || link.closest('div[class*="item"]') || link.parentElement?.parentElement;
      if (!card) return;

      const nameEl = card.querySelector('h2, h3, [class*="name"], [class*="title"]');
      const priceEl = card.querySelector('[class*="price"], [class*="Price"]');
      const areaEl = card.querySelector('[class*="area"], [class*="address"], [class*="station"]');
      const capEl = card.querySelector('[class*="capacity"], [class*="people"]');
      const imgEl = card.querySelector('img');

      const priceText = priceEl?.textContent?.trim() || '';
      const priceMatch = priceText.match(/[\d,]+/);
      const hourlyPrice = priceMatch ? parseInt(priceMatch[0].replace(/,/g, '')) : null;

      venues.push({
        name: nameEl?.textContent?.trim() || link.textContent?.trim()?.substring(0, 50) || '不明',
        address: areaEl?.textContent?.trim() || '',
        station: '',
        officialUrl: href,
        bookingUrl: href,
        hourlyPrice: hourlyPrice,
        priceDetail: priceText,
        capacity: capEl?.textContent?.trim() || '',
        photoUrl: imgEl?.src || '',
        platform: 'インスタベース',
        commercialUse: '要確認',
        transferPayment: '○',
        paymentDetail: '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い',
        equipment: '',
        bookingMethod: 'インスタベースから予約',
        note: ''
      });
    });

    // 詳細ページ
    if (url.match(/\/space\/\d+/)) {
      const name = document.querySelector('h1')?.textContent?.trim();
      const priceEl = document.querySelector('[class*="price"], [class*="Price"]');
      const priceText = priceEl?.textContent?.trim() || '';
      const priceMatch = priceText.match(/[\d,]+/);
      const addressEl = document.querySelector('[class*="address"], [class*="access"]');

      if (name) {
        venues.push({
          name: name,
          address: addressEl?.textContent?.trim() || '',
          station: '',
          officialUrl: url,
          bookingUrl: url,
          hourlyPrice: priceMatch ? parseInt(priceMatch[0].replace(/,/g, '')) : null,
          priceDetail: priceText,
          capacity: '',
          photoUrl: document.querySelector('meta[property="og:image"]')?.content || '',
          platform: 'インスタベース',
          commercialUse: '要確認',
          transferPayment: '○',
          paymentDetail: '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い',
          equipment: '',
          bookingMethod: 'インスタベースから予約',
          note: ''
        });
      }
    }
  }

  // ----- スペースマーケット -----
  if (url.includes('spacemarket.com')) {
    document.querySelectorAll('a[href*="/spaces/"]').forEach(link => {
      const href = link.href;
      if (href.includes('/search') || href.includes('/spaces?')) return;

      const card = link.closest('[class*="Card"]') || link.closest('li') || link.closest('article') || link.parentElement?.parentElement;
      if (!card) return;

      const nameEl = card.querySelector('h2, h3, [class*="name"], [class*="title"], p');
      const priceEl = card.querySelector('[class*="price"], [class*="Price"]');
      const priceText = priceEl?.textContent?.trim() || '';
      const priceMatch = priceText.match(/[\d,]+/);
      const imgEl = card.querySelector('img');

      const name = nameEl?.textContent?.trim() || '';
      if (!name || name.length < 2) return;

      venues.push({
        name: name.substring(0, 80),
        address: '',
        station: '',
        officialUrl: href,
        bookingUrl: href,
        hourlyPrice: priceMatch ? parseInt(priceMatch[0].replace(/,/g, '')) : null,
        priceDetail: priceText,
        capacity: '',
        photoUrl: imgEl?.src || '',
        platform: 'スペースマーケット',
        commercialUse: '要確認',
        transferPayment: '○',
        paymentDetail: '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い',
        equipment: '',
        bookingMethod: 'スペースマーケットから予約',
        note: ''
      });
    });
  }

  // ----- スペイシー -----
  if (url.includes('spacee.jp')) {
    document.querySelectorAll('a[href*="/listings/"]').forEach(link => {
      const href = link.href;
      const card = link.closest('[class*="card"]') || link.closest('li') || link.closest('div');
      if (!card) return;

      const nameEl = card.querySelector('h2, h3, [class*="name"], [class*="title"]');
      const priceEl = card.querySelector('[class*="price"], [class*="Price"]');
      const priceText = priceEl?.textContent?.trim() || '';
      const priceMatch = priceText.match(/[\d,]+/);

      venues.push({
        name: nameEl?.textContent?.trim() || '不明',
        address: '',
        station: '',
        officialUrl: href,
        bookingUrl: href,
        hourlyPrice: priceMatch ? parseInt(priceMatch[0].replace(/,/g, '')) : null,
        priceDetail: priceText,
        capacity: '',
        photoUrl: '',
        platform: 'スペイシー',
        commercialUse: '要確認',
        transferPayment: '要確認',
        paymentDetail: '',
        equipment: '',
        bookingMethod: 'スペイシーから予約',
        note: ''
      });
    });
  }

  // ----- 汎用抽出（Google検索結果等） -----
  if (url.includes('google.co.jp/search') || url.includes('google.com/search')) {
    document.querySelectorAll('#search .g, [data-hveid]').forEach(result => {
      const linkEl = result.querySelector('a[href^="http"]');
      const titleEl = result.querySelector('h3');
      const snippetEl = result.querySelector('[data-sncf], .VwiC3b, [class*="snippet"]');

      if (!linkEl || !titleEl) return;
      const title = titleEl.textContent.trim();
      const snippet = snippetEl?.textContent?.trim() || '';

      // 施設っぽいものだけフィルタ
      const keywords = ['貸会議室', 'レンタルスペース', '公民館', '会議室', 'コワーキング', 'ホテル', '商工会議所', 'TKP', '図書館'];
      const isVenue = keywords.some(kw => title.includes(kw) || snippet.includes(kw));
      if (!isVenue) return;

      venues.push({
        name: title.substring(0, 80),
        address: '',
        station: '',
        officialUrl: linkEl.href,
        bookingUrl: '',
        hourlyPrice: null,
        priceDetail: snippet.substring(0, 100),
        capacity: '',
        photoUrl: '',
        platform: 'Google検索',
        commercialUse: '要確認',
        transferPayment: '要確認',
        paymentDetail: '',
        equipment: '',
        bookingMethod: '要確認',
        note: 'Google検索結果から抽出。詳細は公式サイトを確認'
      });
    });
  }

  return venues;
}

// ===== データ管理 =====
function addVenues(newVenues, source) {
  const area = document.getElementById('areas')?.value?.trim()?.split('\n')[0] || source || '不明';

  for (const v of newVenues) {
    // 重複チェック（名前+プラットフォーム）
    const isDup = venues.some(existing =>
      existing.name === v.name && existing.platform === v.platform
    );
    if (!isDup) {
      v.area = v.area || area;
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

  // エリアフィルター更新
  const areas = [...new Set(venues.map(v => v.area))];
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

  const filtered = filterArea ? venues.filter(v => v.area === filterArea) : venues;

  if (filtered.length === 0) {
    list.innerHTML = `
      <div class="empty-state">
        <p>まだ施設データがありません</p>
        <p>「検索」タブで一括検索を実行してください</p>
      </div>`;
    return;
  }

  // 料金昇順ソート
  filtered.sort((a, b) => {
    const pa = a.hourlyPrice || 999999;
    const pb = b.hourlyPrice || 999999;
    return pa - pb;
  });

  list.innerHTML = filtered.map(v => `
    <div class="venue-card">
      <div class="venue-name">${escHtml(v.name)}</div>
      <div class="venue-area">${escHtml(v.platform)} | ${escHtml(v.area || '')}</div>
      <div class="venue-info">
        <span class="venue-price">${v.hourlyPrice ? '¥' + v.hourlyPrice.toLocaleString() + '/h' : '料金不明'}</span>
        <span>${escHtml(v.capacity ? '定員: ' + v.capacity : '')}</span>
        <span>${escHtml(v.address || '')}</span>
        <span>振込: ${escHtml(v.transferPayment || '不明')}</span>
      </div>
      <div class="venue-actions">
        ${v.officialUrl ? `<button onclick="window.open('${escHtml(v.officialUrl)}')">公式サイト</button>` : ''}
        <button class="btn-remove" data-id="${v.id}">削除</button>
      </div>
    </div>
  `).join('');

  // 削除ボタン
  list.querySelectorAll('.btn-remove').forEach(btn => {
    btn.addEventListener('click', () => removeVenue(parseFloat(btn.dataset.id)));
  });
}

// ===== Excel出力 =====
function exportExcel() {
  if (venues.length === 0) {
    alert('出力するデータがありません');
    return;
  }

  const fulldayHours = parseInt(document.getElementById('fullday-hours')?.value) || 8;
  const wb = XLSX.utils.book_new();

  // ----- 全施設一覧シート -----
  const allData = venues.map(v => ({
    '対象エリア': v.area || '',
    '施設名': v.name || '',
    '施設住所': v.address || '',
    '最寄駅': v.station || '',
    '公式URL': v.officialUrl || '',
    '予約URL': v.bookingUrl || '',
    'Google Mapリンク': v.address ? `https://www.google.com/maps/search/${encodeURIComponent(v.address)}` : '',
    '施設担当者名': v.contactName || '',
    '担当者連絡先': v.contactInfo || '',
    '商用利用可否': v.commercialUse || '要確認',
    '1時間あたり料金（円）': v.hourlyPrice || '',
    '終日利用概算（税込）': v.hourlyPrice ? v.hourlyPrice * fulldayHours : '',
    '料金詳細': v.priceDetail || '',
    '振込対応': v.transferPayment || '要確認',
    '支払方法の詳細': v.paymentDetail || '',
    '設備': v.equipment || '',
    '収容人数': v.capacity || '',
    '予約方法': v.bookingMethod || '',
    '写真URL': v.photoUrl || '',
    '備考': v.note || ''
  }));

  const ws1 = XLSX.utils.json_to_sheet(allData);

  // 列幅設定
  ws1['!cols'] = [
    { wch: 15 }, { wch: 30 }, { wch: 30 }, { wch: 12 },
    { wch: 35 }, { wch: 35 }, { wch: 35 },
    { wch: 12 }, { wch: 15 }, { wch: 10 },
    { wch: 15 }, { wch: 18 }, { wch: 30 },
    { wch: 10 }, { wch: 25 }, { wch: 20 },
    { wch: 10 }, { wch: 15 }, { wch: 35 }, { wch: 30 }
  ];

  // オートフィルター
  const range = XLSX.utils.decode_range(ws1['!ref']);
  ws1['!autofilter'] = { ref: XLSX.utils.encode_range(range) };

  XLSX.utils.book_append_sheet(wb, ws1, '全施設一覧');

  // ----- エリア別シート -----
  const areas = [...new Set(venues.map(v => v.area))];
  for (const area of areas) {
    const areaVenues = venues
      .filter(v => v.area === area)
      .sort((a, b) => (a.hourlyPrice || 999999) - (b.hourlyPrice || 999999));

    const areaData = areaVenues.map(v => ({
      '施設名': v.name || '',
      '1時間あたり料金（円）': v.hourlyPrice || '',
      '終日利用概算（税込）': v.hourlyPrice ? v.hourlyPrice * fulldayHours : '',
      '料金詳細': v.priceDetail || '',
      '振込対応': v.transferPayment || '要確認',
      '商用利用可否': v.commercialUse || '要確認',
      '施設住所': v.address || '',
      '公式URL': v.officialUrl || '',
      '予約URL': v.bookingUrl || '',
      '収容人数': v.capacity || '',
      '設備': v.equipment || '',
      '備考': v.note || ''
    }));

    const wsArea = XLSX.utils.json_to_sheet(areaData);
    wsArea['!cols'] = [
      { wch: 30 }, { wch: 15 }, { wch: 18 }, { wch: 30 },
      { wch: 10 }, { wch: 10 }, { wch: 30 },
      { wch: 35 }, { wch: 35 }, { wch: 10 },
      { wch: 20 }, { wch: 30 }
    ];

    // シート名は31文字以内に制限
    const sheetName = area.substring(0, 28) || 'エリア';
    XLSX.utils.book_append_sheet(wb, wsArea, sheetName);
  }

  // ----- 支払方法まとめシート -----
  const paymentData = [
    { 'プラットフォーム': 'スペースマーケット', '振込対応': '○', '条件': '法人Paid登録（審査2-3営業日）', '支払サイクル': '月末締め→翌月末払い' },
    { 'プラットフォーム': 'インスタベース', '振込対応': '○', '条件': '法人Paid登録（審査即時～3営業日）', '支払サイクル': '月末締め→翌月末払い' },
    { 'プラットフォーム': 'upnow', '振込対応': '○', '条件': '請求書・領収書発行可', '支払サイクル': '都度' },
    { 'プラットフォーム': '日本会議室', '振込対応': '○', '条件': '法人利用対応', '支払サイクル': '都度請求書' },
    { 'プラットフォーム': 'TKP', '振込対応': '○', '条件': '法人請求書払い標準対応', '支払サイクル': '都度請求書' },
    { 'プラットフォーム': '公共施設', '振込対応': '△', '条件': '窓口現金精算が基本', '支払サイクル': '利用当日精算' },
    { 'プラットフォーム': '個人運営スペース', '振込対応': '要確認', '条件': '施設による', '支払サイクル': '施設による' }
  ];
  const wsPay = XLSX.utils.json_to_sheet(paymentData);
  wsPay['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 35 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, wsPay, '支払方法まとめ');

  // ----- 次のアクションシート -----
  const actionItems = [];

  // 料金不明の施設
  venues.filter(v => !v.hourlyPrice).forEach(v => {
    actionItems.push({
      '優先度': '高',
      'アクション': '料金を確認',
      '対象施設': v.name,
      '連絡先': v.contactInfo || v.officialUrl || '',
      '備考': ''
    });
  });

  // 商用利用要確認
  venues.filter(v => v.commercialUse === '要確認').forEach(v => {
    actionItems.push({
      '優先度': '高',
      'アクション': '商用利用可否を確認',
      '対象施設': v.name,
      '連絡先': v.contactInfo || v.officialUrl || '',
      '備考': ''
    });
  });

  // 公共施設の商行為確認
  venues.filter(v => v.platform === 'Google検索' && v.name.match(/公民館|市民|図書館|ルピア/)).forEach(v => {
    actionItems.push({
      '優先度': '高',
      'アクション': '面接利用が商行為/入場料徴収に該当するか電話確認',
      '対象施設': v.name,
      '連絡先': v.contactInfo || '',
      '備考': '商行為該当の場合は料金2～3倍の可能性'
    });
  });

  if (actionItems.length === 0) {
    actionItems.push({ '優先度': '', 'アクション': 'アクション項目なし', '対象施設': '', '連絡先': '', '備考': '' });
  }

  const wsAction = XLSX.utils.json_to_sheet(actionItems);
  wsAction['!cols'] = [{ wch: 8 }, { wch: 35 }, { wch: 30 }, { wch: 30 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, wsAction, '次のアクション');

  // ファイル出力
  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `会場調査_${dateStr}.xlsx`);
}

// ===== 設定の保存・復元 =====
function saveSettings() {
  searchSettings = {
    periodStart: document.getElementById('period-start').value,
    periodEnd: document.getElementById('period-end').value,
    payTransfer: document.getElementById('pay-transfer').checked,
    payInvoice: document.getElementById('pay-invoice').checked,
    payCorporate: document.getElementById('pay-corporate').checked,
    otherConditions: document.getElementById('other-conditions').value
  };
  chrome.storage.local.set({ searchSettings });
  alert('設定を保存しました');
}

function restoreSettings() {
  if (searchSettings.periodStart) document.getElementById('period-start').value = searchSettings.periodStart;
  if (searchSettings.periodEnd) document.getElementById('period-end').value = searchSettings.periodEnd;
  if (searchSettings.payTransfer !== undefined) document.getElementById('pay-transfer').checked = searchSettings.payTransfer;
  if (searchSettings.payInvoice !== undefined) document.getElementById('pay-invoice').checked = searchSettings.payInvoice;
  if (searchSettings.payCorporate !== undefined) document.getElementById('pay-corporate').checked = searchSettings.payCorporate;
  if (searchSettings.otherConditions) document.getElementById('other-conditions').value = searchSettings.otherConditions;
}

// ===== ユーティリティ =====
function escHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}

function addStatus(container, text, type) {
  const line = document.createElement('div');
  line.className = `status-line ${type}`;
  line.textContent = text;
  container.appendChild(line);
  container.scrollTop = container.scrollHeight;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ===== 状態管理 =====
let venues = [];
let searchSettings = {};

// 自動保存対象フォームフィールド
const FORM_FIELDS = ['purpose', 'areas', 'budget', 'capacity', 'fullday-hours', 'extra-keywords'];

// デフォルト検索キーワード
const DEFAULT_SEARCH_KEYWORDS = [
  '貸会議室',
  'レンタルスペース 会議',
  '公民館 会議室 貸出',
  'コワーキングスペース 個室',
  'ホテル 会議室 貸出',
  '商工会議所 会議室',
  '市民センター 会議室',
  '図書館 会議室',
  'TKP 貸会議室',
  'レンタルオフィス 時間貸し'
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
  document.getElementById('btn-extract').addEventListener('click', extractCurrentPage);
  document.getElementById('btn-export').addEventListener('click', exportExcel);
  document.getElementById('btn-clear').addEventListener('click', clearAllResults);
  document.getElementById('btn-save-settings').addEventListener('click', saveSettings);
  document.getElementById('filter-area').addEventListener('change', renderResults);

  // 保存データ復元
  const stored = await chrome.storage.local.get([
    'venues', 'searchSettings', 'formData', 'activeTab'
  ]);
  if (stored.venues) venues = stored.venues;
  if (stored.searchSettings) {
    searchSettings = stored.searchSettings;
    restoreSettings();
  }
  if (stored.formData) restoreFormData(stored.formData);
  if (stored.activeTab) switchTab(stored.activeTab);

  updateResultCount();
  renderResults();

  // フォーム自動保存
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', saveFormData);
      el.addEventListener('change', saveFormData);
    }
  });

  // 支払フィルター変更で結果を即更新
  ['pay-transfer', 'pay-corporate'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', () => { renderResults(); });
  });

  // ストレージ変更を監視（コンテンツスクリプトの抽出結果をリアルタイム反映）
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

// ===== 一括検索 =====
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

  saveFormData();

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>検索中...';
  statusArea.innerHTML = '';

  // 基準地点のジオコーディング
  addStatus(statusArea, '基準地点の座標を取得中...', 'info');
  const areaCoords = {};
  for (const area of areas) {
    const coords = await geocode(area);
    if (coords) {
      areaCoords[area] = coords;
      addStatus(statusArea, `${area} → 座標取得OK`, 'success');
    } else {
      addStatus(statusArea, `${area} → 座標取得失敗（距離計算なしで続行）`, 'error');
    }
    await sleep(1100); // Nominatim APIレート制限: 1req/sec
  }

  // 座標をストレージに保存（コンテンツスクリプトで使う）
  await chrome.storage.local.set({ areaCoords });

  // === Google検索タブ ===
  const keywords = getSearchKeywords();
  let openedCount = 0;
  for (const area of areas) {
    for (const keyword of keywords) {
      const query = extra
        ? `${area} ${keyword} ${extra}`
        : `${area} ${keyword}`;
      const url = `https://www.google.co.jp/search?q=${encodeURIComponent(query)}&num=20`;

      addStatus(statusArea, `${area} →「${keyword}」Google検索...`, 'info');
      try {
        await chrome.tabs.create({ url, active: false });
        openedCount++;
        if (openedCount % 5 === 0) await sleep(800);
      } catch (e) {
        addStatus(statusArea, `タブ作成失敗: ${e.message}`, 'error');
      }
    }
  }

  // === プラットフォーム直接検索タブ ===
  for (const area of areas) {
    const kw = encodeURIComponent(area);
    const cap = document.getElementById('capacity').value || 2;
    const platformUrls = [
      { label: 'インスタベース', url: `https://www.instabase.jp/search?keyword=${kw}&pax=${cap}&category=meetingroom` },
      { label: 'スペースマーケット', url: `https://www.spacemarket.com/spaces?keyword=${kw}&people=${cap}&types%5B%5D=meeting_room` },
      { label: 'スペイシー', url: `https://www.spacee.jp/listings?location=${kw}&capacity=${cap}` },
    ];
    for (const p of platformUrls) {
      addStatus(statusArea, `${area} → ${p.label} 直接検索...`, 'info');
      try {
        await chrome.tabs.create({ url: p.url, active: false });
        openedCount++;
        if (openedCount % 5 === 0) await sleep(800);
      } catch (e) {
        addStatus(statusArea, `${p.label} タブ作成失敗: ${e.message}`, 'error');
      }
    }
  }

  addStatus(statusArea, `${openedCount}件のタブを開きました（Google検索＋プラットフォーム直接）`, 'success');
  addStatus(statusArea, 'Google検索タブは自動抽出。プラットフォームは「現在のページからデータ抽出」で取得できます。', 'info');

  btn.disabled = false;
  btn.innerHTML = '一括検索開始';
}

// ===== ジオコーディング（Nominatim） =====
async function geocode(address) {
  try {
    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(address)}&format=json&limit=1&countrycodes=jp`;
    const res = await fetch(url, {
      headers: { 'User-Agent': 'VenueFinderExtension/2.0' }
    });
    const data = await res.json();
    if (data && data.length > 0) {
      return { lat: parseFloat(data[0].lat), lon: parseFloat(data[0].lon) };
    }
  } catch (e) {
    console.error('Geocode error:', e);
  }
  return null;
}

// ===== Haversine距離計算（km） =====
function calcDistanceKm(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
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

    // ページの種類に応じた抽出関数を選択
    const isGoogle = tab.url.includes('google.co.jp/search') || tab.url.includes('google.com/search');
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: isGoogle ? extractFromGoogleSearch : extractFromPlatformPage,
      args: isGoogle ? [] : [tab.url]
    });

    if (results && results[0] && results[0].result) {
      const extracted = results[0].result;
      if (extracted.length > 0) {
        // 距離計算を付与
        const stored = await chrome.storage.local.get(['areaCoords']);
        const areaCoords = stored.areaCoords || {};
        const areasText = document.getElementById('areas').value.trim();
        const firstArea = areasText.split('\n').map(a => a.trim()).filter(a => a)[0] || '';

        for (const v of extracted) {
          v.area = v.area || firstArea;
        }

        addVenues(extracted);
        addStatus(statusArea, `${extracted.length}件の施設を抽出しました`, 'success');
      } else {
        addStatus(statusArea, '施設データが見つかりませんでした', 'error');
      }
    }
  } catch (e) {
    addStatus(statusArea, `抽出エラー: ${e.message}`, 'error');
  }
}

// ===== Google検索結果から施設情報を抽出（ページ内実行） =====
function extractFromGoogleSearch() {
  const venues = [];
  const seen = new Set();

  // Google検索クエリからエリア名を推測
  const searchInput = document.querySelector('input[name="q"], textarea[name="q"]');
  const query = searchInput?.value || '';

  // 検索結果を走査
  document.querySelectorAll('#search .g, #rso .g, [data-hveid] .g, #rso > div > div').forEach(result => {
    const linkEl = result.querySelector('a[href^="http"]');
    const titleEl = result.querySelector('h3');
    if (!linkEl || !titleEl) return;

    const title = titleEl.textContent.trim();
    const href = linkEl.href;

    // 重複排除
    if (seen.has(href)) return;
    seen.add(href);

    // スニペットから情報抽出
    const snippetEl = result.querySelector('.VwiC3b, [data-sncf], [class*="snippet"], .lEBKkf, span:not(h3 span)');
    const snippet = snippetEl?.textContent?.trim() || '';

    // 住所パターン抽出
    const addrMatch = snippet.match(/〒?\d{3}-?\d{4}\s*[^\d].{5,30}/) ||
                      snippet.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県).{2,20}[区市町村].{1,20}/) ||
                      snippet.match(/[^\s]{1,5}[区市町村][^\s]{1,20}/);
    const address = addrMatch ? addrMatch[0].trim() : '';

    // 電話番号パターン
    const phoneMatch = snippet.match(/(?:0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4})/);
    const phone = phoneMatch ? phoneMatch[0] : '';

    // 料金パターン（できるだけ多くのパターンを拾う）
    const allText = title + ' ' + snippet;
    // 全ての料金っぽい記述を料金詳細に残す
    const allPriceMatches = allText.match(/\d{1,3},?\d{3}円[^\s。、]{0,20}/g) || [];
    const priceDetailFromSnippet = allPriceMatches.join(' / ');

    // 時間単価の抽出（優先順）
    const priceMatch = allText.match(/(\d{1,2},?\d{3})円[/／]?(?:時間|h|1h|1時間)/i) ||
                       allText.match(/(?:時間|1h|1時間)[あたり]*[^\d]*(\d{1,2},?\d{3})円/i) ||
                       allText.match(/(\d{3,5})円[~～〜／/](?:時間|h)/i) ||
                       allText.match(/(?:¥|￥)(\d{1,2},?\d{3})[^\d]*[/／]?(?:時間|h)/i) ||
                       allText.match(/(\d{3,5})円[~～〜／/]/) ||
                       allText.match(/(\d{1,2},?\d{3})円/);
    let hourlyPrice = null;
    if (priceMatch) {
      hourlyPrice = parseInt(priceMatch[1].replace(/,/g, ''));
    }

    // 施設っぽいかフィルタ（広く取る）
    const venueKeywords = [
      '会議室', 'レンタルスペース', '貸会議室', '公民館', 'コワーキング',
      'ホテル', '商工会議所', '図書館', '市民', 'センター', 'TKP',
      'リージャス', 'スペース', '研修', 'ルーム', '個室', '多目的',
      'インスタベース', 'スペースマーケット', 'スペイシー'
    ];
    const isVenue = venueKeywords.some(kw => title.includes(kw) || snippet.includes(kw));

    // プラットフォーム系のまとめページは除外（個別施設ページのみ）
    const isAggregatorList = /まとめ|ランキング|おすすめ\d+選|比較/.test(title) && !address;

    if (!isVenue || isAggregatorList) return;

    // プラットフォーム判定
    let platform = '公式サイト';
    if (href.includes('instabase.jp')) platform = 'インスタベース';
    else if (href.includes('spacemarket.com')) platform = 'スペースマーケット';
    else if (href.includes('spacee.jp')) platform = 'スペイシー';
    else if (href.includes('upnow.jp')) platform = 'upnow';
    else if (href.includes('tkp.jp')) platform = 'TKP';
    else if (href.includes('nihonkaigishitsu')) platform = '日本会議室';
    else if (href.includes('regus.') || href.includes('regus-')) platform = 'リージャス';

    // 振込対応の推定
    let transferPayment = '要確認';
    let paymentDetail = '';
    const platformPaymentMap = {
      'インスタベース': { transfer: '○', detail: '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い' },
      'スペースマーケット': { transfer: '○', detail: '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い' },
      'upnow': { transfer: '○', detail: '請求書・領収書発行可' },
      'TKP': { transfer: '○', detail: '法人請求書払い標準対応' },
      '日本会議室': { transfer: '○', detail: '法人利用対応、都度請求書' },
      'リージャス': { transfer: '○', detail: '法人契約対応' }
    };
    if (platformPaymentMap[platform]) {
      transferPayment = platformPaymentMap[platform].transfer;
      paymentDetail = platformPaymentMap[platform].detail;
    }
    // 公共施設判定
    const isPublic = /公民館|市民|図書館|センター|商工会議所/.test(title);
    if (isPublic) {
      transferPayment = '△';
      paymentDetail = '窓口現金精算が基本。法人振込は要確認';
    }

    venues.push({
      name: title.substring(0, 80),
      address: address,
      station: '',
      officialUrl: href,
      bookingUrl: (platform !== '公式サイト') ? href : '',
      hourlyPrice: hourlyPrice,
      priceDetail: priceDetailFromSnippet || (priceMatch ? priceMatch[0] : ''),
      capacity: '',
      photoUrl: '',
      platform: platform,
      commercialUse: isPublic ? '要確認' : '可',
      transferPayment: transferPayment,
      paymentDetail: paymentDetail,
      equipment: '',
      bookingMethod: platform !== '公式サイト' ? `${platform}から予約` : '要確認',
      contactInfo: phone,
      note: isPublic ? '公共施設：商行為該当の場合は料金2～3倍の可能性。要電話確認' : '',
      area: '',
      distanceKm: null
    });
  });

  return venues;
}

// ===== プラットフォームページから施設情報を抽出（ページ内実行） =====
function extractFromPlatformPage(pageUrl) {
  const url = pageUrl || location.href;
  const venues = [];
  const seen = new Set();

  let platform = '不明';
  let linkSelector = '';
  let payment = { transfer: '要確認', detail: '' };

  if (url.includes('instabase.jp')) {
    platform = 'インスタベース';
    linkSelector = 'a[href*="/space/"]';
    payment = { transfer: '○', detail: '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い' };
  } else if (url.includes('spacemarket.com')) {
    platform = 'スペースマーケット';
    linkSelector = 'a[href*="/spaces/"]';
    payment = { transfer: '○', detail: '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い' };
  } else if (url.includes('spacee.jp')) {
    platform = 'スペイシー';
    linkSelector = 'a[href*="/listings/"]';
  } else {
    return venues;
  }

  const links = document.querySelectorAll(linkSelector);
  links.forEach(link => {
    const href = link.href;
    if (href.includes('/search') || href.includes('?keyword') || href.includes('/spaces?')) return;

    const idMatch = href.match(/\/(?:space|spaces|listings)\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) return;
    if (seen.has(idMatch[1])) return;
    seen.add(idMatch[1]);

    const card = link.closest('[class*="Card"]') || link.closest('[class*="card"]')
      || link.closest('li') || link.closest('article')
      || link.closest('[class*="item"]') || link.parentElement?.parentElement;
    if (!card) return;

    const nameEl = card.querySelector('h2, h3, h4, [class*="name"], [class*="title"], [class*="Name"], [class*="Title"]');
    const name = nameEl?.textContent?.trim() || '';
    if (!name || name.length < 2) return;

    const priceEl = card.querySelector('[class*="price"], [class*="Price"]');
    const priceText = priceEl?.textContent?.trim() || card.textContent || '';
    const allPriceMatches = priceText.match(/\d{1,3},?\d{3}円[^\s。、]{0,15}/g) || [];
    const priceDetail = allPriceMatches.join(' / ');
    const hourlyMatch = priceText.match(/(\d{1,2},?\d{3})円[/／]?(?:時間|h|1h|1時間)/i)
      || priceText.match(/(\d{3,5})円[~～〜／/]/)
      || priceText.match(/(\d{1,2},?\d{3})円/);
    const hourlyPrice = hourlyMatch ? parseInt(hourlyMatch[1].replace(/,/g, '')) : null;

    const areaEl = card.querySelector('[class*="area"], [class*="address"], [class*="station"], [class*="location"]');
    const capEl = card.querySelector('[class*="capacity"], [class*="people"]');
    const imgEl = card.querySelector('img');

    venues.push({
      name: name.substring(0, 80),
      address: areaEl?.textContent?.trim() || '',
      station: '',
      officialUrl: href,
      bookingUrl: href,
      hourlyPrice: hourlyPrice,
      priceDetail: priceDetail || (hourlyMatch ? hourlyMatch[0] : ''),
      capacity: capEl?.textContent?.trim() || '',
      photoUrl: imgEl?.src || '',
      platform: platform,
      commercialUse: '可',
      transferPayment: payment.transfer,
      paymentDetail: payment.detail,
      equipment: '',
      bookingMethod: `${platform}から予約`,
      contactInfo: '',
      note: '',
      area: '',
      distanceKm: null
    });
  });

  return venues;
}

// ===== データ管理 =====
function addVenues(newVenues) {
  for (const v of newVenues) {
    // 重複チェック（URL単位）
    const isDup = venues.some(existing => existing.officialUrl === v.officialUrl);
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

// ===== 距離を一括計算 =====
async function calcDistancesForAll() {
  const stored = await chrome.storage.local.get(['areaCoords']);
  const areaCoords = stored.areaCoords || {};
  let updated = false;

  for (const v of venues) {
    if (v.distanceKm !== null) continue; // 計算済み
    if (!v.address || !v.area) continue;

    const origin = areaCoords[v.area];
    if (!origin) continue;

    // 施設の住所をジオコーディング
    const dest = await geocode(v.address);
    if (dest) {
      v.distanceKm = Math.round(calcDistanceKm(origin.lat, origin.lon, dest.lat, dest.lon) * 10) / 10;
      updated = true;
      await sleep(1100); // Nominatimレート制限
    }
  }

  if (updated) saveVenues();
}

// ===== 結果表示 =====
function renderResults() {
  const list = document.getElementById('results-list');
  const filterArea = document.getElementById('filter-area').value;

  // エリアフィルター更新
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

  let filtered = filterArea ? venues.filter(v => v.area === filterArea) : [...venues];
  filtered = applyPaymentFilter(filtered);

  if (filtered.length === 0) {
    list.innerHTML = `
      <div class="empty-state">
        <p>まだ施設データがありません</p>
        <p>「検索」タブで一括検索を実行してください</p>
      </div>`;
    return;
  }

  // 距離→料金の順でソート
  filtered.sort((a, b) => {
    const da = a.distanceKm ?? 9999;
    const db = b.distanceKm ?? 9999;
    if (da !== db) return da - db;
    return (a.hourlyPrice || 999999) - (b.hourlyPrice || 999999);
  });

  list.innerHTML = filtered.map(v => {
    const distText = v.distanceKm !== null ? `${v.distanceKm}km` : '距離不明';
    const priceText = v.hourlyPrice ? `¥${v.hourlyPrice.toLocaleString()}/h` : '料金不明';
    const mapsUrl = v.address
      ? `https://www.google.com/maps/dir/${encodeURIComponent(v.area)}/${encodeURIComponent(v.address)}`
      : `https://www.google.com/maps/search/${encodeURIComponent(v.name)}`;

    return `
    <div class="venue-card">
      <div class="venue-name">${escHtml(v.name)}</div>
      <div class="venue-meta">
        <span class="venue-distance">${distText}</span>
        <span class="venue-price">${priceText}</span>
        <span class="venue-platform">${escHtml(v.platform)}</span>
      </div>
      <div class="venue-info">
        <span>${escHtml(v.address || '住所不明')}</span>
        <span>振込: ${escHtml(v.transferPayment || '不明')}</span>
      </div>
      ${v.priceDetail ? `<div class="venue-price-detail">${escHtml(v.priceDetail)}</div>` : ''}
      <div class="venue-actions">
        <button onclick="window.open('${escAttr(v.officialUrl)}')">サイト</button>
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
  if (venues.length === 0) {
    alert('出力するデータがありません');
    return;
  }

  const fulldayHours = parseInt(document.getElementById('fullday-hours')?.value) || 8;
  const wb = XLSX.utils.book_new();

  // 全施設一覧シート（全件出力、フィルターで除外しない）
  const allData = venues.map(v => ({
    '基準地点': v.area || '',
    '距離（km）': v.distanceKm ?? '',
    '施設名': v.name || '',
    '施設住所': v.address || '',
    '公式URL': v.officialUrl || '',
    '予約URL': v.bookingUrl || '',
    'Google Map経路': v.address && v.area
      ? `https://www.google.com/maps/dir/${encodeURIComponent(v.area)}/${encodeURIComponent(v.address)}`
      : '',
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
    { wch: 20 }, { wch: 10 }, { wch: 30 }, { wch: 30 },
    { wch: 40 }, { wch: 40 }, { wch: 50 },
    { wch: 15 }, { wch: 10 },
    { wch: 15 }, { wch: 18 }, { wch: 30 },
    { wch: 10 }, { wch: 25 }, { wch: 15 },
    { wch: 10 }, { wch: 20 }, { wch: 15 }, { wch: 30 }
  ];
  const range1 = XLSX.utils.decode_range(ws1['!ref']);
  ws1['!autofilter'] = { ref: XLSX.utils.encode_range(range1) };
  XLSX.utils.book_append_sheet(wb, ws1, '全施設一覧');

  // エリア別シート
  const areas = [...new Set(venues.map(v => v.area).filter(a => a))];
  for (const area of areas) {
    const areaVenues = venues
      .filter(v => v.area === area)
      .sort((a, b) => (a.distanceKm ?? 9999) - (b.distanceKm ?? 9999));

    const areaData = areaVenues.map(v => ({
      '施設名': v.name || '',
      '距離（km）': v.distanceKm ?? '',
      '1時間あたり料金（円）': v.hourlyPrice || '',
      '終日利用概算（税込）': v.hourlyPrice ? v.hourlyPrice * fulldayHours : '',
      '料金詳細': v.priceDetail || '',
      '振込対応': v.transferPayment || '要確認',
      '商用利用可否': v.commercialUse || '要確認',
      '施設住所': v.address || '',
      '公式URL': v.officialUrl || '',
      'Google Map経路': v.address
        ? `https://www.google.com/maps/dir/${encodeURIComponent(area)}/${encodeURIComponent(v.address)}`
        : '',
      '備考': v.note || ''
    }));

    const wsArea = XLSX.utils.json_to_sheet(areaData);
    wsArea['!cols'] = [
      { wch: 30 }, { wch: 10 }, { wch: 15 }, { wch: 18 }, { wch: 25 },
      { wch: 10 }, { wch: 10 }, { wch: 30 },
      { wch: 40 }, { wch: 50 }, { wch: 30 }
    ];
    XLSX.utils.book_append_sheet(wb, wsArea, area.substring(0, 28));
  }

  // 支払方法まとめシート
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

  // 次のアクションシート
  const actionItems = [];
  venues.filter(v => !v.hourlyPrice).forEach(v => {
    actionItems.push({ '優先度': '高', 'アクション': '料金を確認', '対象施設': v.name, '連絡先': v.contactInfo || v.officialUrl || '', '備考': '' });
  });
  venues.filter(v => v.commercialUse === '要確認').forEach(v => {
    actionItems.push({ '優先度': '高', 'アクション': '商用利用可否を確認', '対象施設': v.name, '連絡先': v.contactInfo || v.officialUrl || '', '備考': '' });
  });
  venues.filter(v => /公民館|市民|図書館/.test(v.name)).forEach(v => {
    actionItems.push({ '優先度': '高', 'アクション': '面接利用が商行為に該当するか電話確認', '対象施設': v.name, '連絡先': v.contactInfo || '', '備考': '商行為該当の場合は料金2～3倍' });
  });
  if (actionItems.length === 0) {
    actionItems.push({ '優先度': '', 'アクション': 'アクション項目なし', '対象施設': '', '連絡先': '', '備考': '' });
  }
  const wsAction = XLSX.utils.json_to_sheet(actionItems);
  wsAction['!cols'] = [{ wch: 8 }, { wch: 35 }, { wch: 30 }, { wch: 30 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(wb, wsAction, '次のアクション');

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(wb, `会場調査_${dateStr}.xlsx`);
}

// ===== 設定の保存・復元 =====
function saveSettings() {
  searchSettings = {
    periodStart: document.getElementById('period-start').value,
    periodEnd: document.getElementById('period-end').value,
    payTransfer: document.getElementById('pay-transfer').checked,
    payCorporate: document.getElementById('pay-corporate').checked,
    searchKeywords: document.getElementById('search-keywords').value,
    otherConditions: document.getElementById('other-conditions').value
  };
  chrome.storage.local.set({ searchSettings });
  renderResults(); // フィルター変更を即反映
  alert('設定を保存しました');
}

function restoreSettings() {
  if (searchSettings.periodStart) document.getElementById('period-start').value = searchSettings.periodStart;
  if (searchSettings.periodEnd) document.getElementById('period-end').value = searchSettings.periodEnd;
  if (searchSettings.payTransfer !== undefined) document.getElementById('pay-transfer').checked = searchSettings.payTransfer;
  if (searchSettings.payCorporate !== undefined) document.getElementById('pay-corporate').checked = searchSettings.payCorporate;
  if (searchSettings.searchKeywords) {
    document.getElementById('search-keywords').value = searchSettings.searchKeywords;
  } else {
    document.getElementById('search-keywords').value = DEFAULT_SEARCH_KEYWORDS.join('\n');
  }
  if (searchSettings.otherConditions) document.getElementById('other-conditions').value = searchSettings.otherConditions;
}

// 支払フィルター適用
function applyPaymentFilter(list) {
  const filterTransfer = document.getElementById('pay-transfer')?.checked;
  const filterCorporate = document.getElementById('pay-corporate')?.checked;

  return list.filter(v => {
    if (filterTransfer && v.transferPayment !== '○') return false;
    if (filterCorporate && !/法人|請求書/.test(v.paymentDetail)) return false;
    return true;
  });
}

// ===== ユーティリティ =====
function escHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}

function escAttr(str) {
  return (str || '').replace(/'/g, "\\'").replace(/"/g, '&quot;');
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

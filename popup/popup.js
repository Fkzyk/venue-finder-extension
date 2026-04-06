// ===== 状態管理 =====
let venues = [];
let searchCancelled = false;
let searchRunning = false;

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

  document.getElementById('btn-search').addEventListener('click', startBulkSearch);
  document.getElementById('btn-fetch-prices').addEventListener('click', fetchPricesForAll);
  document.getElementById('btn-stop').addEventListener('click', stopSearch);
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
    if (area === 'local' && changes.venues && !searchRunning) {
      // 検索中は自分で管理するので外部変更を無視
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

// ===== 中止ボタン =====
function stopSearch() {
  searchCancelled = true;
  document.getElementById('btn-stop').style.display = 'none';
}

function setSearchRunning(running) {
  searchRunning = running;
  const stopBtn = document.getElementById('btn-stop');
  if (running) {
    searchCancelled = false;
    stopBtn.style.display = 'block';
  } else {
    stopBtn.style.display = 'none';
  }
}

// =========================================================
// 第1段階：fetchでGoogle検索HTMLを取得→施設URLを抽出
// タブを一切開かない
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
  const delay = (parseInt(document.getElementById('request-delay').value) || 3) * 1000;

  saveFormData();
  btn.disabled = true;
  setSearchRunning(true);
  statusArea.innerHTML = '';

  // 全検索URLを生成
  const searchUrls = [];
  for (const area of areas) {
    for (const keyword of keywords) {
      const query = extra ? `${area} ${keyword} ${extra}` : `${area} ${keyword}`;
      searchUrls.push({
        url: `https://www.google.co.jp/search?q=${encodeURIComponent(query)}&num=20`,
        label: `${area} →「${keyword}」`,
        area
      });
    }
  }

  addStatus(statusArea, `全${searchUrls.length}件を${delay/1000}秒間隔で検索（バックグラウンド）`, 'info');

  let doneCount = 0;
  let foundTotal = 0;
  let consecutiveErrors = 0;

  for (const item of searchUrls) {
    if (searchCancelled) {
      addStatus(statusArea, `中止しました（${doneCount}/${searchUrls.length}完了）`, 'error');
      break;
    }

    doneCount++;
    btn.innerHTML = `<span class="spinner"></span>${doneCount}/${searchUrls.length}`;
    addStatus(statusArea, `[${doneCount}/${searchUrls.length}] ${item.label}`, 'info');

    try {
      const html = await fetchPage(item.url);
      const extracted = parseGoogleResults(html, item.area);
      if (extracted.length > 0) {
        let added = 0;
        for (const v of extracted) {
          const isDup = venues.some(ex => ex.officialUrl === v.officialUrl);
          if (!isDup) { v.id = Date.now() + Math.random(); venues.push(v); added++; }
        }
        foundTotal += added;
        if (added > 0) {
          addStatus(statusArea, `  → ${added}件追加（合計${venues.length}件）`, 'success');
          saveVenues();          // 見つかるたびに即保存
          updateResultCount();
        }
      }
      consecutiveErrors = 0;  // 成功したらリセット
    } catch (e) {
      addStatus(statusArea, `  → ${e.message}`, 'error');
      consecutiveErrors++;
      if (consecutiveErrors >= 5) {
        addStatus(statusArea, `連続エラー${consecutiveErrors}回のため自動中止（ボット判定の可能性）`, 'error');
        break;
      }
    }

    if (doneCount < searchUrls.length && !searchCancelled) await sleep(delay);
  }

  saveVenues();
  updateResultCount();
  renderResults();

  btn.disabled = false;
  setSearchRunning(false);
  btn.innerHTML = '第1段階：施設を一括検索';
  if (!searchCancelled) addStatus(statusArea, `検索完了。新規${foundTotal}件、合計${venues.length}件`, 'success');
}

// =========================================================
// 第2段階：fetchで各施設ページを取得→料金を汎用抽出
// タブを一切開かない
// =========================================================
async function fetchPricesForAll() {
  const btn = document.getElementById('btn-fetch-prices');
  const statusArea = document.getElementById('search-status');
  const maxFetch = parseInt(document.getElementById('max-fetch').value) || 20;
  const delay = (parseInt(document.getElementById('request-delay').value) || 3) * 1000;

  const allTargets = venues.filter(v => !v.hourlyPrice && v.officialUrl);
  if (allTargets.length === 0) {
    addStatus(statusArea, '料金未取得の施設がありません', 'success');
    return;
  }
  const targets = allTargets.slice(0, maxFetch);

  btn.disabled = true;
  setSearchRunning(true);
  statusArea.innerHTML = '';
  addStatus(statusArea, `料金未取得${allTargets.length}件中、${targets.length}件を取得します`, 'info');

  let doneCount = 0;
  let successCount = 0;
  let consecutiveErrors = 0;

  for (const venue of targets) {
    if (searchCancelled) {
      addStatus(statusArea, `中止しました（${doneCount}/${targets.length}完了）`, 'error');
      break;
    }

    doneCount++;
    btn.innerHTML = `<span class="spinner"></span>${doneCount}/${targets.length}`;
    addStatus(statusArea, `[${doneCount}/${targets.length}] ${venue.name.substring(0, 35)}`, 'info');

    try {
      const html = await fetchPage(venue.officialUrl);
      const d = extractPriceFromHtml(html);
      if (d.hourlyPrice || d.priceDetail) {
        venue.hourlyPrice = d.hourlyPrice || venue.hourlyPrice;
        venue.priceDetail = d.priceDetail || venue.priceDetail;
        venue.address = d.address || venue.address;
        venue.contactInfo = d.phone || venue.contactInfo;
        venue.capacity = d.capacity || venue.capacity;
        successCount++;
        const ps = d.hourlyPrice ? `¥${d.hourlyPrice}/h` : '';
        addStatus(statusArea, `  → ${ps} ${(d.priceDetail || '').substring(0, 50)}`, 'success');
      } else {
        addStatus(statusArea, `  → 料金情報なし`, 'error');
      }
      saveVenues();          // 取得するたびに即保存
      updateResultCount();
      consecutiveErrors = 0;
    } catch (e) {
      addStatus(statusArea, `  → ${e.message}`, 'error');
      consecutiveErrors++;
      if (consecutiveErrors >= 5) {
        addStatus(statusArea, `連続エラー${consecutiveErrors}回のため自動中止`, 'error');
        break;
      }
    }

    if (doneCount < targets.length && !searchCancelled) await sleep(delay);
  }

  saveVenues();
  updateResultCount();
  renderResults();

  btn.disabled = false;
  setSearchRunning(false);
  btn.innerHTML = '第2段階：料金を一括取得';
  if (!searchCancelled) addStatus(statusArea, `完了: ${successCount}/${targets.length}件で料金取得`, 'success');
}

// =========================================================
// fetchでページHTMLを取得
// =========================================================
async function fetchPage(url) {
  const res = await fetch(url, {
    headers: {
      'Accept': 'text/html,application/xhtml+xml',
      'Accept-Language': 'ja,en;q=0.9',
    }
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return await res.text();
}

// =========================================================
// Google検索結果HTMLをパースして施設情報を抽出
// =========================================================
function parseGoogleResults(html, area) {
  const venues = [];
  const seen = new Set();

  // HTMLからリンクとタイトルを抽出（正規表現ベース、DOMParser不使用でも可）
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  doc.querySelectorAll('.g, [data-hveid]').forEach(result => {
    const linkEl = result.querySelector('a[href^="http"]');
    const titleEl = result.querySelector('h3');
    if (!linkEl || !titleEl) return;

    const title = titleEl.textContent.trim();
    const href = linkEl.href;
    if (seen.has(href)) return;
    seen.add(href);

    // スニペット
    const snippetEl = result.querySelector('.VwiC3b, [data-sncf], .lEBKkf');
    const snippet = snippetEl?.textContent?.trim() || '';
    const allText = title + ' ' + snippet;

    // 施設フィルタ
    const venueKw = ['会議室','レンタルスペース','貸会議室','公民館','コワーキング','ホテル','商工会議所','図書館','市民','センター','TKP','リージャス','スペース','研修','ルーム','個室','多目的','インスタベース','スペースマーケット','スペイシー'];
    const isVenue = venueKw.some(kw => allText.includes(kw));
    const isAgg = /まとめ|ランキング|おすすめ\d+選|比較/.test(title);
    if (!isVenue || isAgg) return;

    // 住所
    const addrMatch = snippet.match(/〒?\d{3}-?\d{4}\s*[^\d].{5,30}/)
      || snippet.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県).{2,20}[区市町村].{1,20}/);
    const address = addrMatch ? addrMatch[0].trim() : '';

    // 電話
    const phoneMatch = snippet.match(/0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}/);
    const phone = phoneMatch ? phoneMatch[0] : '';

    // 料金（スニペットから）
    const priceMatches = allText.match(/\d{1,3},?\d{3}円[^\s。、]{0,20}/g) || [];
    const hourlyMatch = allText.match(/(\d{1,2},?\d{3})円[/／]?(?:時間|h|1h|1時間)/i)
      || allText.match(/(\d{3,5})円[~～〜／/]/)
      || allText.match(/(\d{1,2},?\d{3})円/);
    const hourlyPrice = hourlyMatch ? parseInt(hourlyMatch[1].replace(/,/g, '')) : null;

    // プラットフォーム判定
    let platform = '公式サイト';
    if (href.includes('instabase.jp')) platform = 'インスタベース';
    else if (href.includes('spacemarket.com')) platform = 'スペースマーケット';
    else if (href.includes('spacee.jp')) platform = 'スペイシー';
    else if (href.includes('upnow.jp')) platform = 'upnow';
    else if (href.includes('tkp.jp')) platform = 'TKP';
    else if (href.includes('nihonkaigishitsu')) platform = '日本会議室';

    // 振込
    let transferPayment = '要確認', paymentDetail = '';
    const pMap = {
      'インスタベース': ['○', '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い'],
      'スペースマーケット': ['○', '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い'],
      'upnow': ['○', '請求書・領収書発行可'],
      'TKP': ['○', '法人請求書払い標準対応'],
      '日本会議室': ['○', '法人利用対応、都度請求書'],
    };
    if (pMap[platform]) { transferPayment = pMap[platform][0]; paymentDetail = pMap[platform][1]; }
    const isPublic = /公民館|市民|図書館|センター|商工会議所/.test(title);
    if (isPublic) { transferPayment = '△'; paymentDetail = '窓口現金精算が基本'; }

    venues.push({
      name: title.substring(0, 80),
      address, station: '', officialUrl: href,
      bookingUrl: platform !== '公式サイト' ? href : '',
      hourlyPrice, priceDetail: priceMatches.join(' / '),
      capacity: '', photoUrl: '', platform,
      commercialUse: isPublic ? '要確認' : '可',
      transferPayment, paymentDetail,
      equipment: '',
      bookingMethod: platform !== '公式サイト' ? `${platform}から予約` : '要確認',
      contactInfo: phone,
      note: isPublic ? '公共施設：商行為該当の場合は料金2～3倍の可能性' : '',
      area, distanceKm: null
    });
  });

  return venues;
}

// =========================================================
// HTMLテキストから料金を汎用抽出（第2段階用）
// =========================================================
function extractPriceFromHtml(html) {
  // HTMLタグを除去してテキスト化
  const text = html.replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#\d+;/g, '')
    .replace(/\s+/g, ' ');

  // --- 時間単価 ---
  const hourlyPatterns = [
    /(\d{1,3}[,，]\d{3})\s*円\s*[/／~〜]\s*(?:1\s*)?(?:時間|[hH])/,
    /(?:1\s*)?(?:時間)\s*[あたり:：]*\s*[^\d]{0,5}?(\d{1,3}[,，]\d{3})\s*円/,
    /[¥￥]\s*(\d{1,3}[,，]\d{3})\s*[/／]\s*(?:時間|[hH])/,
    /(\d{3,5})\s*円\s*[/／~〜]\s*(?:1\s*)?(?:時間|[hH])/,
    /(?:1\s*)?(?:時間)\s*[あたり:：]*\s*[^\d]{0,5}?(\d{3,5})\s*円/,
    /(?:料金|価格|利用料|使用料|単価)[^\d]{0,15}?(\d{1,3}[,，]?\d{3})\s*円/,
  ];

  let hourlyPrice = null;
  for (const pat of hourlyPatterns) {
    const m = text.match(pat);
    if (m) {
      const p = parseInt(m[1].replace(/[,，]/g, ''));
      if (p > 50 && p < 50000) { hourlyPrice = p; break; }
    }
  }

  // 分単位→時間換算
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

  // 料金テキスト収集
  const priceTexts = [];
  const seen = new Set();
  const matches = text.match(/[^\s]{0,10}\d{1,3}[,，]?\d{3}円[^\s]{0,25}/g) || [];
  for (const m of matches) {
    const c = m.trim();
    if (!seen.has(c)) { seen.add(c); priceTexts.push(c); }
    if (priceTexts.length >= 10) break;
  }

  // 住所
  const addrMatch = text.match(/〒\d{3}-?\d{4}[^<\n]{3,40}/)
    || text.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県)[^<\n]{2,15}[区市町村][^<\n]{1,20}[\d\-]+/);
  const address = addrMatch ? addrMatch[0].trim().substring(0, 60) : '';

  // 電話
  const phoneMatch = text.match(/(?:TEL|電話|tel|Tel|☎)[^\d]{0,5}(0\d{1,4}[-ー\s]?\d{1,4}[-ー\s]?\d{3,4})/)
    || text.match(/(0\d{1,4}-\d{1,4}-\d{3,4})/);
  const phone = phoneMatch ? (phoneMatch[1] || phoneMatch[0]) : '';

  // 定員
  const capMatch = text.match(/(?:定員|収容|着席)\s*[：:]*\s*(\d{1,4})\s*(?:名|人)/);
  const capacity = capMatch ? capMatch[1] + '名' : '';

  return { hourlyPrice, priceDetail: priceTexts.join(' | '), address, phone, capacity };
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
        ${v.officialUrl ? `<button onclick="window.open('${escA(v.officialUrl)}')">サイト</button>` : ''}
        <button onclick="window.open('${escA(maps)}')">地図</button>
        <button class="btn-remove" data-id="${v.id}">削除</button>
      </div>
    </div>`;
  }).join('');

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
  XLSX.writeFile(wb, `会場調査_${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}.xlsx`);
}

// ===== ユーティリティ =====
function esc(s){const d=document.createElement('div');d.textContent=s||'';return d.innerHTML}
function escA(s){return(s||'').replace(/'/g,"\\'").replace(/"/g,'&quot;')}
function addStatus(c,t,type){const l=document.createElement('div');l.className=`status-line ${type}`;l.textContent=t;c.appendChild(l);c.scrollTop=c.scrollHeight}
function sleep(ms){return new Promise(r=>setTimeout(r,ms))}

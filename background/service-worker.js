// ===== バックグラウンド サービスワーカー =====
// 検索ロジックはすべてここで実行。ポップアップを閉じても止まらない。

// ===== 重複チェック =====

// URLからドメイン部分を取得
function getDomain(url) {
  try { return new URL(url).hostname.replace(/^www\./, '').toLowerCase(); }
  catch (_) { return ''; }
}

// URLを正規化（同じページの別バリエーションを統一）
function normalizeUrl(url) {
  try {
    const u = new URL(url);
    let path = u.pathname.replace(/\/+$/, '');
    // 末尾の /access, /price, /map, /info, /guide, /detail 等を除去（施設のサブページ）
    path = path.replace(/\/(access|price|fee|map|info|guide|detail|contact|about|review|photo|gallery|images|facilities|equipment|faq|inquiry|reserve|booking|plan)$/i, '');
    return (u.hostname.replace(/^www\./, '') + path).toLowerCase();
  } catch (_) {
    return url.toLowerCase().replace(/[?#].*$/, '').replace(/\/+$/, '');
  }
}

// 施設名を正規化
function normalizeName(name) {
  return (name || '')
    .replace(/[\s\u3000　・|｜\-ー–—―／/\\【】「」『』（）()[\]""''《》<>＜＞]/g, '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/(の|（.*?）|\(.*?\)|料金|アクセス|地図|予約|口コミ|評判|公式)$/g, '')
    .toLowerCase();
}

// プラットフォームサイトのドメイン一覧
const PLATFORM_DOMAINS = [
  'instabase.jp', 'spacemarket.com', 'spacee.jp', 'upnow.jp',
  'tkp.jp', 'nihonkaigishitsu.co.jp', 'natuluck.com', 'timeroom.jp'
];

function isDuplicate(venues, v) {
  const newUrl = normalizeUrl(v.officialUrl || '');
  const newDomain = getDomain(v.officialUrl || '');
  const newName = normalizeName(v.name);
  const isPlatform = PLATFORM_DOMAINS.some(d => newDomain.includes(d));

  return venues.some(ex => {
    const exUrl = normalizeUrl(ex.officialUrl || '');
    const exDomain = getDomain(ex.officialUrl || '');
    const exName = normalizeName(ex.name);

    // 1. URL一致（サブページ除去後）
    if (newUrl && exUrl === newUrl) return true;

    // 2. 同じドメインの別ページ（プラットフォーム以外）
    //    例：venue-x.co.jp/room-a と venue-x.co.jp/access → 同じ施設
    if (!isPlatform && newDomain && exDomain === newDomain) return true;

    // 3. 施設名一致
    if (newName.length >= 4 && exName.length >= 4) {
      if (newName === exName) return true;
      // 一方が他方を含む（6文字以上）
      if (newName.length >= 6 && exName.length >= 6) {
        if (newName.includes(exName) || exName.includes(newName)) return true;
      }
    }

    return false;
  });
}

let searchState = {
  running: false,
  cancelled: false,
  phase: '',        // 'phase1' or 'phase2'
  progress: '',     // 進捗テキスト
  done: 0,
  total: 0,
  log: [],          // ステータスログ（最新50件）
};

function addLog(text, type) {
  searchState.log.push({ text, type, time: Date.now() });
  if (searchState.log.length > 80) searchState.log = searchState.log.slice(-50);
  broadcastState();
}

function broadcastState() {
  chrome.storage.local.set({ searchState: { ...searchState } });
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// =========================================================
// タブをURLへ移動し、読み込み完了を待つ
// =========================================================
function navigateAndWait(tabId, url) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      reject(new Error('ページ読み込みタイムアウト'));
    }, 30000);

    function listener(updatedTabId, changeInfo) {
      if (updatedTabId === tabId && changeInfo.status === 'complete') {
        clearTimeout(timeout);
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    }
    chrome.tabs.onUpdated.addListener(listener);
    chrome.tabs.update(tabId, { url }).catch(e => {
      clearTimeout(timeout);
      chrome.tabs.onUpdated.removeListener(listener);
      reject(e);
    });
  });
}

// =========================================================
// 第1段階：バックグラウンドタブ1つでGoogle検索
// =========================================================
async function runPhase1(params) {
  const { areas, keywords, extra, delay } = params;

  searchState = {
    running: true, cancelled: false, phase: 'phase1',
    progress: '', done: 0, total: 0, log: [],
  };

  // 検索URL一覧を生成
  const searchUrls = [];
  for (const area of areas) {
    for (const keyword of keywords) {
      const query = extra ? `${area} ${keyword} ${extra}` : `${area} ${keyword}`;
      searchUrls.push({
        url: `https://www.google.co.jp/search?q=${encodeURIComponent(query)}&num=20&hl=ja`,
        label: `${area} →「${keyword}」`,
        area
      });
    }
  }

  searchState.total = searchUrls.length;
  const baseDelay = delay;
  let currentDelay = delay;
  addLog(`全${searchUrls.length}件を${baseDelay/1000}秒間隔で検索開始`, 'info');

  // バックグラウンドタブを1つ作成
  let searchTab;
  try {
    searchTab = await chrome.tabs.create({ url: 'about:blank', active: false });
    addLog('検索用タブを作成（非アクティブ）', 'info');
  } catch (e) {
    addLog(`タブ作成失敗: ${e.message}`, 'error');
    searchState.running = false;
    broadcastState();
    return;
  }

  // 既存データを取得
  const stored = await chrome.storage.local.get(['venues']);
  let venues = stored.venues || [];
  let foundTotal = 0;
  let consecutiveErrors = 0;

  for (const item of searchUrls) {
    if (searchState.cancelled) {
      addLog(`中止しました（${searchState.done}/${searchUrls.length}完了）`, 'error');
      break;
    }

    searchState.done++;
    searchState.progress = `[${searchState.done}/${searchUrls.length}] ${item.label}`;
    addLog(searchState.progress, 'info');
    broadcastState();

    try {
      await navigateAndWait(searchTab.id, item.url);

      const results = await chrome.scripting.executeScript({
        target: { tabId: searchTab.id },
        func: extractGoogleResultsFromPage
      });

      const extracted = results?.[0]?.result || [];
      if (extracted.length > 0) {
        let added = 0;
        for (const v of extracted) {
          v.area = item.area;
          if (!isDuplicate(venues, v)) { v.id = Date.now() + Math.random(); venues.push(v); added++; }
        }
        foundTotal += added;
        if (added > 0) {
          addLog(`  → ${added}件追加（合計${venues.length}件）`, 'success');
          await chrome.storage.local.set({ venues });
          chrome.action.setBadgeText({ text: String(venues.length) });
          chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
        }
      } else {
        addLog('  → 結果なし', 'info');
      }
      consecutiveErrors = 0;
      if (currentDelay > baseDelay) {
        currentDelay = Math.max(baseDelay, currentDelay - 1000);
        addLog(`  間隔を${currentDelay/1000}秒に短縮`, 'info');
      }
    } catch (e) {
      consecutiveErrors++;
      currentDelay = Math.min(currentDelay * 2, 60000);
      addLog(`  → ${e.message}（間隔を${currentDelay/1000}秒に延長、${consecutiveErrors}回目）`, 'error');
      if (consecutiveErrors >= 8) {
        addLog(`連続エラー${consecutiveErrors}回のため自動中止`, 'error');
        break;
      }
    }

    if (searchState.done < searchUrls.length && !searchState.cancelled) await sleep(currentDelay);
  }

  try { await chrome.tabs.remove(searchTab.id); } catch (_) {}

  if (!searchState.cancelled) {
    addLog(`検索完了。新規${foundTotal}件、合計${venues.length}件`, 'success');
  }
  searchState.running = false;
  broadcastState();
}

// =========================================================
// 第2段階：バックグラウンドタブ1つで料金取得
// =========================================================
async function runPhase2(params) {
  const { delay, maxFetch } = params;

  searchState = {
    running: true, cancelled: false, phase: 'phase2',
    progress: '', done: 0, total: 0, log: [],
  };

  const stored = await chrome.storage.local.get(['venues']);
  let venues = stored.venues || [];
  const allTargets = venues.filter(v => !v.hourlyPrice && v.officialUrl);

  if (allTargets.length === 0) {
    addLog('料金未取得の施設がありません', 'success');
    searchState.running = false;
    broadcastState();
    return;
  }

  const targets = allTargets.slice(0, maxFetch);
  searchState.total = targets.length;

  // バックグラウンドタブ作成
  let priceTab;
  try {
    priceTab = await chrome.tabs.create({ url: 'about:blank', active: false });
    addLog('料金取得用タブを作成（非アクティブ）', 'info');
  } catch (e) {
    addLog(`タブ作成失敗: ${e.message}`, 'error');
    searchState.running = false;
    broadcastState();
    return;
  }

  const baseDelay = delay;
  let currentDelay = delay;
  addLog(`料金未取得${allTargets.length}件中、${targets.length}件を取得（エラー時は自動延長）`, 'info');

  let successCount = 0;
  let consecutiveErrors = 0;

  for (const target of targets) {
    if (searchState.cancelled) {
      addLog(`中止しました（${searchState.done}/${targets.length}完了）`, 'error');
      break;
    }

    searchState.done++;
    searchState.progress = `[${searchState.done}/${targets.length}] ${target.name.substring(0, 35)}`;
    addLog(searchState.progress, 'info');
    broadcastState();

    try {
      await navigateAndWait(priceTab.id, target.officialUrl);

      const results = await chrome.scripting.executeScript({
        target: { tabId: priceTab.id },
        func: extractPriceFromPage
      });

      const d = results?.[0]?.result || {};

      // venuesの該当要素を更新
      const venue = venues.find(v => v.id === target.id);
      if (venue && (d.hourlyPrice || d.priceDetail)) {
        venue.hourlyPrice = d.hourlyPrice || venue.hourlyPrice;
        venue.priceDetail = d.priceDetail || venue.priceDetail;
        venue.address = d.address || venue.address;
        venue.contactInfo = d.phone || venue.contactInfo;
        venue.capacity = d.capacity || venue.capacity;
        successCount++;
        const ps = d.hourlyPrice ? `¥${d.hourlyPrice}/h` : '';
        addLog(`  → ${ps} ${(d.priceDetail || '').substring(0, 50)}`, 'success');
      } else {
        addLog('  → 料金情報なし', 'info');
      }
      await chrome.storage.local.set({ venues });
      consecutiveErrors = 0;
      if (currentDelay > baseDelay) {
        currentDelay = Math.max(baseDelay, currentDelay - 1000);
        addLog(`  間隔を${currentDelay/1000}秒に短縮`, 'info');
      }
    } catch (e) {
      consecutiveErrors++;
      currentDelay = Math.min(currentDelay * 2, 60000);
      addLog(`  → ${e.message}（間隔を${currentDelay/1000}秒に延長、${consecutiveErrors}回目）`, 'error');
      if (consecutiveErrors >= 8) {
        addLog(`連続エラー${consecutiveErrors}回のため自動中止`, 'error');
        break;
      }
    }

    if (searchState.done < targets.length && !searchState.cancelled) await sleep(currentDelay);
  }

  try { await chrome.tabs.remove(priceTab.id); } catch (_) {}

  if (!searchState.cancelled) {
    addLog(`完了: ${successCount}/${targets.length}件で料金取得`, 'success');
  }
  searchState.running = false;
  broadcastState();
}

// =========================================================
// Google検索結果ページで実行（注入用）
// =========================================================
function extractGoogleResultsFromPage() {
  const results = [];
  const seen = new Set();

  document.querySelectorAll('#search .g, #rso .g, [data-hveid]').forEach(el => {
    const linkEl = el.querySelector('a[href^="http"]');
    const titleEl = el.querySelector('h3');
    if (!linkEl || !titleEl) return;

    const title = titleEl.textContent.trim();
    const href = linkEl.href;
    if (!href || seen.has(href)) return;
    if (href.includes('google.com') || href.includes('google.co.jp')) return;
    seen.add(href);

    const snippetEl = el.querySelector('[data-sncf], .VwiC3b, .lEBKkf, [style*="-webkit-line-clamp"]');
    const snippet = snippetEl?.textContent?.trim() || '';
    const allText = title + ' ' + snippet;

    const venueKw = ['会議室','レンタルスペース','貸会議室','公民館','コワーキング','ホテル',
      '商工会議所','図書館','市民','センター','TKP','リージャス','スペース','研修','ルーム',
      '個室','多目的','インスタベース','スペースマーケット','スペイシー','貸室','ホール'];
    const isVenue = venueKw.some(kw => allText.includes(kw));
    const isAgg = /まとめ|ランキング|おすすめ\d+選|比較/.test(title);
    if (!isVenue || isAgg) return;

    const addrMatch = snippet.match(/〒?\d{3}-?\d{4}\s*[^\d].{5,30}/)
      || snippet.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県).{2,20}[区市町村].{1,20}/);
    const address = addrMatch ? addrMatch[0].trim() : '';

    const phoneMatch = snippet.match(/0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}/);
    const phone = phoneMatch ? phoneMatch[0] : '';

    const priceMatches = allText.match(/\d{1,3},?\d{3}円[^\s。、]{0,20}/g) || [];
    const hourlyMatch = allText.match(/(\d{1,2},?\d{3})円[/／]?(?:時間|h|1h|1時間)/i)
      || allText.match(/(\d{3,5})円[~～〜／/]/)
      || allText.match(/(\d{1,2},?\d{3})円/);
    const hourlyPrice = hourlyMatch ? parseInt(hourlyMatch[1].replace(/,/g, '')) : null;

    let platform = '公式サイト';
    if (href.includes('instabase.jp')) platform = 'インスタベース';
    else if (href.includes('spacemarket.com')) platform = 'スペースマーケット';
    else if (href.includes('spacee.jp')) platform = 'スペイシー';
    else if (href.includes('upnow.jp')) platform = 'upnow';
    else if (href.includes('tkp.jp')) platform = 'TKP';
    else if (href.includes('nihonkaigishitsu')) platform = '日本会議室';

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

    results.push({
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
      area: '', distanceKm: null
    });
  });

  return results;
}

// =========================================================
// 施設ページで実行（注入用）
// =========================================================
function extractPriceFromPage() {
  const text = document.body?.innerText || '';

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

  const priceTexts = [];
  const seen = new Set();
  const matches = text.match(/[^\s]{0,10}\d{1,3}[,，]?\d{3}円[^\s]{0,25}/g) || [];
  for (const m of matches) {
    const c = m.trim();
    if (!seen.has(c)) { seen.add(c); priceTexts.push(c); }
    if (priceTexts.length >= 10) break;
  }

  const addrMatch = text.match(/〒\d{3}-?\d{4}[^<\n]{3,40}/)
    || text.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県)[^<\n]{2,15}[区市町村][^<\n]{1,20}[\d\-]+/);
  const address = addrMatch ? addrMatch[0].trim().substring(0, 60) : '';

  const phoneMatch = text.match(/(?:TEL|電話|tel|Tel|☎)[^\d]{0,5}(0\d{1,4}[-ー\s]?\d{1,4}[-ー\s]?\d{3,4})/)
    || text.match(/(0\d{1,4}-\d{1,4}-\d{3,4})/);
  const phone = phoneMatch ? (phoneMatch[1] || phoneMatch[0]) : '';

  const capMatch = text.match(/(?:定員|収容|着席)\s*[：:]*\s*(\d{1,4})\s*(?:名|人)/);
  const capacity = capMatch ? capMatch[1] + '名' : '';

  return { hourlyPrice, priceDetail: priceTexts.join(' | '), address, phone, capacity };
}

// =========================================================
// メッセージリスナー
// =========================================================
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'start-phase1') {
    if (searchState.running) {
      sendResponse({ error: '検索が既に実行中です' });
      return;
    }
    runPhase1(msg.params);
    sendResponse({ ok: true });
  }

  if (msg.type === 'start-phase2') {
    if (searchState.running) {
      sendResponse({ error: '検索が既に実行中です' });
      return;
    }
    runPhase2(msg.params);
    sendResponse({ ok: true });
  }

  if (msg.type === 'stop-search') {
    searchState.cancelled = true;
    sendResponse({ ok: true });
  }

  if (msg.type === 'get-search-state') {
    sendResponse(searchState);
  }

  if (msg.type === 'venues-extracted' && sender.tab) {
    chrome.storage.local.get(['venues'], (result) => {
      const venues = result.venues || [];
      const newVenues = msg.data || [];
      let addedCount = 0;
      for (const v of newVenues) {
        if (!isDuplicate(venues, v)) {
          v.id = Date.now() + Math.random();
          venues.push(v);
          addedCount++;
        }
      }
      if (addedCount > 0) {
        chrome.storage.local.set({ venues });
        chrome.action.setBadgeText({ text: String(venues.length) });
        chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
      }
    });
  }
});

// バッジ初期化
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local' && changes.venues) {
    const v = changes.venues.newValue || [];
    chrome.action.setBadgeText({ text: v.length > 0 ? String(v.length) : '' });
    chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
  }
});

chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.get(['venues'], (result) => {
    const count = (result.venues || []).length;
    if (count > 0) {
      chrome.action.setBadgeText({ text: String(count) });
      chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
    }
  });
});

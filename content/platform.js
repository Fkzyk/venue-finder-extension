// ===== プラットフォーム共通 コンテンツスクリプト =====
// インスタベース・スペースマーケット・スペイシーの検索結果から自動抽出

(function () {
  let extractionDone = false;
  const url = location.href;

  function detectPlatform() {
    if (url.includes('instabase.jp')) return 'インスタベース';
    if (url.includes('spacemarket.com')) return 'スペースマーケット';
    if (url.includes('spacee.jp')) return 'スペイシー';
    return null;
  }

  function getSearchKeyword() {
    const params = new URLSearchParams(location.search);
    return params.get('keyword') || params.get('location') || params.get('q') || '';
  }

  function extractPrices(text) {
    const matches = text.match(/\d{1,3},?\d{3}円[^\s。、]{0,15}/g) || [];
    const priceDetail = matches.join(' / ');

    const hourlyMatch = text.match(/(\d{1,2},?\d{3})円[/／]?(?:時間|h|1h|1時間)/i) ||
      text.match(/(?:時間|1h|1時間)[あたり]*[^\d]*(\d{1,2},?\d{3})円/i) ||
      text.match(/(\d{3,5})円[~～〜／/]/) ||
      text.match(/(\d{1,2},?\d{3})円/);
    const hourlyPrice = hourlyMatch ? parseInt(hourlyMatch[1].replace(/,/g, '')) : null;

    return { hourlyPrice, priceDetail };
  }

  function getPaymentInfo(platform) {
    const map = {
      'インスタベース': { transfer: '○', detail: '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い' },
      'スペースマーケット': { transfer: '○', detail: '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い' },
      'スペイシー': { transfer: '要確認', detail: '' }
    };
    return map[platform] || { transfer: '要確認', detail: '' };
  }

  function tryExtract() {
    if (extractionDone) return;

    const platform = detectPlatform();
    if (!platform) return;

    const keyword = decodeURIComponent(getSearchKeyword()).replace(/\+/g, ' ');
    const area = keyword.split(/\s+/)[0] || '';
    const payment = getPaymentInfo(platform);
    const venues = [];
    const seen = new Set();

    // 共通：ページ内のリンクからスペース/施設を探す
    const linkPatterns = {
      'インスタベース': 'a[href*="/space/"]',
      'スペースマーケット': 'a[href*="/spaces/"]',
      'スペイシー': 'a[href*="/listings/"]'
    };

    const linkSelector = linkPatterns[platform];
    if (!linkSelector) return;

    const links = document.querySelectorAll(linkSelector);
    if (links.length === 0) return;

    extractionDone = true;

    links.forEach(link => {
      const href = link.href;

      // 検索ページ自身やリストページを除外
      if (href.includes('/search') || href.includes('?keyword') || href.includes('/spaces?')) return;

      // IDで重複排除
      const idMatch = href.match(/\/(?:space|spaces|listings)\/([a-zA-Z0-9_-]+)/);
      if (!idMatch) return;
      const id = idMatch[1];
      if (seen.has(id)) return;
      seen.add(id);

      // 最も近い親カード要素を探す
      const card = link.closest('[class*="Card"]')
        || link.closest('[class*="card"]')
        || link.closest('li')
        || link.closest('article')
        || link.closest('[class*="item"]')
        || link.parentElement?.parentElement;
      if (!card) return;

      // 名前
      const nameEl = card.querySelector('h2, h3, h4, [class*="name"], [class*="title"], [class*="Name"], [class*="Title"]');
      const name = nameEl?.textContent?.trim() || '';
      if (!name || name.length < 2) return;

      // 料金
      const priceEl = card.querySelector('[class*="price"], [class*="Price"], [class*="cost"], [class*="Cost"]');
      const cardText = card.textContent || '';
      const priceText = priceEl?.textContent?.trim() || '';
      const { hourlyPrice, priceDetail } = extractPrices(priceText || cardText);

      // 住所・エリア
      const areaEl = card.querySelector('[class*="area"], [class*="address"], [class*="station"], [class*="location"], [class*="Area"], [class*="Address"]');
      const address = areaEl?.textContent?.trim() || '';

      // 定員
      const capEl = card.querySelector('[class*="capacity"], [class*="people"], [class*="Capacity"], [class*="People"]');
      const capacity = capEl?.textContent?.trim() || '';

      // 画像
      const imgEl = card.querySelector('img');
      const photoUrl = imgEl?.src || '';

      venues.push({
        name: name.substring(0, 80),
        address: address,
        station: '',
        officialUrl: href,
        bookingUrl: href,
        hourlyPrice: hourlyPrice,
        priceDetail: priceDetail || priceText,
        capacity: capacity,
        photoUrl: photoUrl,
        platform: platform,
        commercialUse: '可',
        transferPayment: payment.transfer,
        paymentDetail: payment.detail,
        equipment: '',
        bookingMethod: `${platform}から予約`,
        contactInfo: '',
        note: '',
        area: area,
        distanceKm: null
      });
    });

    if (venues.length > 0) {
      chrome.runtime.sendMessage({
        type: 'venues-extracted',
        data: venues,
        source: platform
      });
      console.log(`[会場探し] ${platform}から ${venues.length} 件抽出`);
    }
  }

  // SPA対応：複数回試行 + MutationObserver
  setTimeout(tryExtract, 2000);
  setTimeout(tryExtract, 5000);
  setTimeout(tryExtract, 8000);

  const observer = new MutationObserver(() => {
    if (!extractionDone) setTimeout(tryExtract, 1500);
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();

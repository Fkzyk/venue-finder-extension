// ===== インスタベース コンテンツスクリプト =====
// 検索結果ページで自動的に施設データを抽出

(function() {
  // 検索結果ページのみで動作
  if (!location.pathname.includes('/search') && !location.href.includes('keyword=')) return;

  // ページ読み込み完了後に実行（SPA対応で遅延）
  let extractionDone = false;

  function tryExtract() {
    if (extractionDone) return;

    const links = document.querySelectorAll('a[href*="/space/"]');
    if (links.length === 0) return; // まだ読み込まれていない

    extractionDone = true;
    const venues = [];
    const seen = new Set();

    links.forEach(link => {
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

      // URLからエリア推測
      const urlParams = new URLSearchParams(location.search);
      const keyword = urlParams.get('keyword') || '';

      venues.push({
        name: nameEl?.textContent?.trim() || '不明',
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
        note: '',
        area: decodeURIComponent(keyword).replace(/\+/g, ' ').split(/\s+/)[0] || ''
      });
    });

    if (venues.length > 0) {
      chrome.runtime.sendMessage({
        type: 'venues-extracted',
        data: venues,
        source: 'instabase'
      });
      console.log(`[会場探し] インスタベースから ${venues.length} 件抽出`);
    }
  }

  // 初回実行 + MutationObserverでSPA対応
  setTimeout(tryExtract, 2000);
  setTimeout(tryExtract, 5000);

  const observer = new MutationObserver(() => {
    if (!extractionDone) {
      setTimeout(tryExtract, 1000);
    }
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();

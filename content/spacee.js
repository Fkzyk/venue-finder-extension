// ===== スペイシー コンテンツスクリプト =====

(function() {
  if (!location.pathname.includes('/listings') && !location.href.includes('location=')) return;

  let extractionDone = false;

  function tryExtract() {
    if (extractionDone) return;

    const links = document.querySelectorAll('a[href*="/listings/"]');
    if (links.length === 0) return;

    extractionDone = true;
    const venues = [];
    const seen = new Set();

    links.forEach(link => {
      const href = link.href;
      const match = href.match(/\/listings\/(\d+)/);
      if (!match || seen.has(match[1])) return;
      seen.add(match[1]);

      const card = link.closest('[class*="card"]') || link.closest('li') || link.closest('div[class*="item"]') || link.parentElement?.parentElement;
      if (!card) return;

      const nameEl = card.querySelector('h2, h3, [class*="name"], [class*="title"]');
      const priceEl = card.querySelector('[class*="price"], [class*="Price"]');
      const areaEl = card.querySelector('[class*="area"], [class*="address"]');
      const imgEl = card.querySelector('img');

      const priceText = priceEl?.textContent?.trim() || '';
      const priceMatch = priceText.match(/[\d,]+/);

      const urlParams = new URLSearchParams(location.search);
      const keyword = urlParams.get('location') || '';

      venues.push({
        name: nameEl?.textContent?.trim() || '不明',
        address: areaEl?.textContent?.trim() || '',
        station: '',
        officialUrl: href,
        bookingUrl: href,
        hourlyPrice: priceMatch ? parseInt(priceMatch[0].replace(/,/g, '')) : null,
        priceDetail: priceText,
        capacity: '',
        photoUrl: imgEl?.src || '',
        platform: 'スペイシー',
        commercialUse: '要確認',
        transferPayment: '要確認',
        paymentDetail: '',
        equipment: '',
        bookingMethod: 'スペイシーから予約',
        note: '',
        area: decodeURIComponent(keyword) || ''
      });
    });

    if (venues.length > 0) {
      chrome.runtime.sendMessage({
        type: 'venues-extracted',
        data: venues,
        source: 'spacee'
      });
      console.log(`[会場探し] スペイシーから ${venues.length} 件抽出`);
    }
  }

  setTimeout(tryExtract, 2000);
  setTimeout(tryExtract, 5000);

  const observer = new MutationObserver(() => {
    if (!extractionDone) setTimeout(tryExtract, 1000);
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();

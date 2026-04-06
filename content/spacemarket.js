// ===== スペースマーケット コンテンツスクリプト =====

(function() {
  if (!location.pathname.includes('/spaces') && !location.href.includes('keyword=')) return;

  let extractionDone = false;

  function tryExtract() {
    if (extractionDone) return;

    const links = document.querySelectorAll('a[href*="/spaces/"]');
    const validLinks = Array.from(links).filter(l =>
      !l.href.includes('/spaces?') && !l.href.includes('/search') && l.href.match(/\/spaces\/[a-zA-Z0-9_-]+/)
    );

    if (validLinks.length === 0) return;

    extractionDone = true;
    const venues = [];
    const seen = new Set();

    validLinks.forEach(link => {
      const href = link.href;
      const match = href.match(/\/spaces\/([a-zA-Z0-9_-]+)/);
      if (!match || seen.has(match[1])) return;
      seen.add(match[1]);

      const card = link.closest('[class*="Card"]') || link.closest('li') || link.closest('article') || link.parentElement?.parentElement;
      if (!card) return;

      const nameEl = card.querySelector('h2, h3, [class*="name"], [class*="title"], p');
      const priceEl = card.querySelector('[class*="price"], [class*="Price"]');
      const areaEl = card.querySelector('[class*="area"], [class*="address"]');
      const capEl = card.querySelector('[class*="capacity"], [class*="people"]');
      const imgEl = card.querySelector('img');

      const priceText = priceEl?.textContent?.trim() || '';
      const priceMatch = priceText.match(/[\d,]+/);

      const name = nameEl?.textContent?.trim() || '';
      if (!name || name.length < 2) return;

      const urlParams = new URLSearchParams(location.search);
      const keyword = urlParams.get('keyword') || '';

      venues.push({
        name: name.substring(0, 80),
        address: areaEl?.textContent?.trim() || '',
        station: '',
        officialUrl: href,
        bookingUrl: href,
        hourlyPrice: priceMatch ? parseInt(priceMatch[0].replace(/,/g, '')) : null,
        priceDetail: priceText,
        capacity: capEl?.textContent?.trim() || '',
        photoUrl: imgEl?.src || '',
        platform: 'スペースマーケット',
        commercialUse: '要確認',
        transferPayment: '○',
        paymentDetail: '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い',
        equipment: '',
        bookingMethod: 'スペースマーケットから予約',
        note: '',
        area: decodeURIComponent(keyword).replace(/\+/g, ' ').split(/\s+/)[0] || ''
      });
    });

    if (venues.length > 0) {
      chrome.runtime.sendMessage({
        type: 'venues-extracted',
        data: venues,
        source: 'spacemarket'
      });
      console.log(`[会場探し] スペースマーケットから ${venues.length} 件抽出`);
    }
  }

  setTimeout(tryExtract, 2000);
  setTimeout(tryExtract, 5000);

  const observer = new MutationObserver(() => {
    if (!extractionDone) setTimeout(tryExtract, 1000);
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();

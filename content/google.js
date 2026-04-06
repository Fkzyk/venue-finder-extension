// ===== Google検索結果ページ コンテンツスクリプト =====
// 検索結果から施設情報を自動抽出してバックグラウンドに送信

(function () {
  let extractionDone = false;

  function tryExtract() {
    if (extractionDone) return;

    const results = document.querySelectorAll('#search .g, #rso .g, #rso > div > div');
    if (results.length === 0) return;

    extractionDone = true;
    const venues = [];
    const seen = new Set();

    // 検索クエリからエリア推測
    const searchInput = document.querySelector('input[name="q"], textarea[name="q"]');
    const query = searchInput?.value || '';

    results.forEach(result => {
      const linkEl = result.querySelector('a[href^="http"]');
      const titleEl = result.querySelector('h3');
      if (!linkEl || !titleEl) return;

      const title = titleEl.textContent.trim();
      const href = linkEl.href;

      if (seen.has(href)) return;
      seen.add(href);

      const snippetEl = result.querySelector('.VwiC3b, [data-sncf], .lEBKkf, span:not(h3 span)');
      const snippet = snippetEl?.textContent?.trim() || '';

      // 住所抽出
      const addrMatch = snippet.match(/〒?\d{3}-?\d{4}\s*[^\d].{5,30}/) ||
        snippet.match(/(東京都|北海道|(?:京都|大阪)府|.{2,3}県).{2,20}[区市町村].{1,20}/) ||
        snippet.match(/[^\s]{1,5}[区市町村][^\s]{1,20}/);
      const address = addrMatch ? addrMatch[0].trim() : '';

      // 電話番号
      const phoneMatch = snippet.match(/(?:0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4})/);
      const phone = phoneMatch ? phoneMatch[0] : '';

      // 料金（できるだけ多くのパターンを拾う）
      const allText = title + ' ' + snippet;
      const allPriceMatches = allText.match(/\d{1,3},?\d{3}円[^\s。、]{0,20}/g) || [];
      const priceDetailFromSnippet = allPriceMatches.join(' / ');

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

      // 施設フィルタ
      const venueKeywords = [
        '会議室', 'レンタルスペース', '貸会議室', '公民館', 'コワーキング',
        'ホテル', '商工会議所', '図書館', '市民', 'センター', 'TKP',
        'リージャス', 'スペース', '研修', 'ルーム', '個室', '多目的',
        'インスタベース', 'スペースマーケット', 'スペイシー'
      ];
      const isVenue = venueKeywords.some(kw => title.includes(kw) || snippet.includes(kw));
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

      // 振込対応
      let transferPayment = '要確認';
      let paymentDetail = '';
      const pMap = {
        'インスタベース': { t: '○', d: '法人Paid登録（審査即時～3営業日）、月末締め→翌月末払い' },
        'スペースマーケット': { t: '○', d: '法人Paid登録（審査2-3営業日）、月末締め→翌月末払い' },
        'upnow': { t: '○', d: '請求書・領収書発行可' },
        'TKP': { t: '○', d: '法人請求書払い標準対応' },
        '日本会議室': { t: '○', d: '法人利用対応、都度請求書' }
      };
      if (pMap[platform]) {
        transferPayment = pMap[platform].t;
        paymentDetail = pMap[platform].d;
      }
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
        area: query.split(/\s+/)[0] || '',
        distanceKm: null
      });
    });

    if (venues.length > 0) {
      chrome.runtime.sendMessage({
        type: 'venues-extracted',
        data: venues,
        source: 'google'
      });
      console.log(`[会場探し] Google検索から ${venues.length} 件抽出`);
    }
  }

  setTimeout(tryExtract, 1500);
  setTimeout(tryExtract, 4000);

  const observer = new MutationObserver(() => {
    if (!extractionDone) setTimeout(tryExtract, 1000);
  });
  observer.observe(document.body, { childList: true, subtree: true });
})();

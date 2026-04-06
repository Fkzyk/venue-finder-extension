// ===== バックグラウンド サービスワーカー =====

// コンテンツスクリプトからのメッセージを中継
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'venues-extracted' && sender.tab) {
    // コンテンツスクリプトからの抽出結果 → ストレージに追加
    chrome.storage.local.get(['venues'], (result) => {
      const venues = result.venues || [];
      const newVenues = msg.data || [];
      let addedCount = 0;

      for (const v of newVenues) {
        const isDup = venues.some(existing =>
          existing.name === v.name && existing.platform === v.platform
        );
        if (!isDup) {
          v.id = Date.now() + Math.random();
          venues.push(v);
          addedCount++;
        }
      }

      if (addedCount > 0) {
        chrome.storage.local.set({ venues }, () => {
          // バッジ更新
          chrome.action.setBadgeText({ text: String(venues.length) });
          chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
        });
      }
    });
  }

  if (msg.type === 'get-settings') {
    chrome.storage.local.get(['searchSettings'], (result) => {
      sendResponse(result.searchSettings || {});
    });
    return true; // 非同期レスポンス
  }
});

// ストレージ変更を監視 → ポップアップが開いていれば通知
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local' && changes.venues) {
    // バッジ更新
    const newVenues = changes.venues.newValue || [];
    chrome.action.setBadgeText({ text: newVenues.length > 0 ? String(newVenues.length) : '' });
    chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
  }
});

// インストール時
chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.get(['venues'], (result) => {
    const count = (result.venues || []).length;
    if (count > 0) {
      chrome.action.setBadgeText({ text: String(count) });
      chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
    }
  });
});

// 西軽精機ナレッジ Service Worker
// PWAとしてホーム画面に追加できるようにするための最小限の実装。
// 積極的なキャッシュはしない（開くたびに最新版を取得＝自動更新のため）。
self.addEventListener('install', (e) => self.skipWaiting());
self.addEventListener('activate', (e) => self.clients.claim());
self.addEventListener('fetch', (e) => {});

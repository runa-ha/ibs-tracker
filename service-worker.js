/* =========================================================
   サービスワーカー：画面ファイルをキャッシュして
   電波がない場所でもアプリを開けるようにする。

   方針は「ネット優先・失敗したらキャッシュ」。
   config.js などを更新したとき、すぐ最新が反映される。
   ========================================================= */

const CACHE_NAME = "ibs-tracker-v1"; // ファイル構成を大きく変えたらここの数字を上げる

const ASSETS = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./config.js",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
];

// インストール時に一式をキャッシュ
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

// 古いキャッシュを掃除
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // GASへの通信（記録の送受信）はキャッシュしない
  if (event.request.method !== "GET" || url.origin !== location.origin) return;

  // ネット優先。成功したらキャッシュを更新、失敗したらキャッシュを返す
  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});

// 서비스워커: 앱 껍데기(화면)를 저장해뒀다가, 인터넷이 느려도 앱이 바로 켜지게 해줍니다.
// data.json(실제 리뷰 데이터)은 항상 최신 것을 새로 받아오도록 캐시하지 않습니다.
const CACHE_NAME = "ibk-review-monitor-shell-v1";
const SHELL_FILES = [
  "./index.html",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(
        names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // data.json은 항상 네트워크 우선(최신 리뷰 데이터를 보여줘야 하므로)
  if (url.pathname.endsWith("data.json")) {
    event.respondWith(
      fetch(event.request).catch(() => caches.match(event.request))
    );
    return;
  }

  // 그 외(화면 껍데기)는 캐시 우선, 없으면 네트워크
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});

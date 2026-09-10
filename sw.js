// 바디빌딩 트래커 서비스워커
// 앱 셸을 캐시해서 오프라인에서도 열리게 한다. 기록 데이터는 localStorage 와
// 앱 자체의 오프라인 큐(pending)가 담당하므로 여기서는 정적 파일만 다룬다.
const CACHE = "bb-shell-v1";
const SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable.png",
  "./icons/apple-touch-icon.png",
];

self.addEventListener("install", e => {
  e.waitUntil(
    caches.open(CACHE)
      // 하나 실패해도 나머지는 캐시되도록 개별 처리
      .then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {}))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;
  // POST(구글 시트 API) 와 크로스오리진(Chart.js CDN) 은 건드리지 않는다
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;

  const path = new URL(req.url).pathname;
  // 문서와 inbox.json 은 항상 최신이어야 한다 (inbox 는 내가 커밋하는 반영 대기 목록)
  if (req.mode === "navigate" || path.endsWith("/index.html") || path.endsWith("/inbox.json")) {
    e.respondWith(
      fetch(req).then(res => {
        if (res && res.ok) caches.open(CACHE).then(c => c.put(req, res.clone()));
        return res;
      }).catch(() => caches.match(req).then(hit => hit || caches.match("./index.html")))
    );
    return;
  }

  // 나머지 정적 파일은 stale-while-revalidate
  e.respondWith(
    caches.match(req).then(hit => {
      const fresh = fetch(req).then(res => {
        if (res && res.ok) caches.open(CACHE).then(c => c.put(req, res.clone()));
        return res;
      }).catch(() => hit);
      return hit || fresh;
    })
  );
});

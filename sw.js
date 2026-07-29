const CACHE = "break-cache-v1";
const FILES = ["/manifest.json", "/logo.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE).map(k => caches.delete(k))
    ))
  );
  e.waitUntil(clients.claim());
});

self.addEventListener("fetch", e => {
  if (e.request.mode === "navigate") {
    return;
  }
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request))
  );
});

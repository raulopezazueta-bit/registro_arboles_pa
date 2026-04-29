/* ════════════════════════════════════════════════════════════
   Service Worker — Árboles PA
   Estrategia: cache-first para shell estático, network-first para
   datos. Permite operación offline tras la primera carga.
   ════════════════════════════════════════════════════════════ */
const CACHE_VERSION = 'apa-2026-04-29-v1';
const CORE = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './favicon-32.png'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_VERSION).then(cache => cache.addAll(CORE))
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  event.respondWith(
    caches.match(req).then(cached => {
      if (cached) {
        /* Cache-first: refresca en segundo plano */
        fetch(req).then(fresh => {
          if (fresh && fresh.status === 200) {
            caches.open(CACHE_VERSION).then(c => c.put(req, fresh.clone()));
          }
        }).catch(() => {});
        return cached;
      }
      /* Sin caché: red con fallback al shell */
      return fetch(req).then(resp => {
        if (resp && resp.status === 200 && req.url.startsWith(self.location.origin)) {
          const clone = resp.clone();
          caches.open(CACHE_VERSION).then(c => c.put(req, clone));
        }
        return resp;
      }).catch(() => caches.match('./index.html'));
    })
  );
});

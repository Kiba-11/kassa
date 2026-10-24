// Офлайн-режим: приложение открывается из кэша, обновление подтягивается в фоне
// и применяется при следующем запуске. При изменении файлов увеличьте VERSION.
const VERSION = 'kassa-v1';
const ASSETS = ['./', './index.html', './xlsx.js', './manifest.webmanifest', './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(caches.open(VERSION).then(async cache => {
    const cached = await cache.match(req, { ignoreSearch: true });
    const fresh = fetch(req).then(res => {
      if (res.ok) cache.put(req, res.clone());
      return res;
    }).catch(() => null);
    if (cached) {
      e.waitUntil(fresh);
      return cached;
    }
    return (await fresh) || (req.mode === 'navigate' ? cache.match('./index.html') : Response.error());
  }));
});

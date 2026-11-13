// Trigger GitHub Pages deployment - OneSignal SDK v16
// Deployment refresh marker: 2026-11-13
// OneSignal Web SDK v16 digabung dengan service worker shell PWA GPS FM.
importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js');

const CACHE = 'gps-fm-shell-v17';
const CACHE_PREFIX = 'gps-fm-shell-';
const FILES = ['./', './index.html', './manifest.json', './config.js', './icon.png'];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      .then(cache => cache.addAll(FILES))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      // Jangan menghapus cache aplikasi lain yang berbagi origin GitHub Pages.
      .then(keys => Promise.all(keys
        .filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE)
        .map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== 'GET' || url.origin !== self.location.origin) return;

  // Konfigurasi deployment harus selalu dibaca dari jaringan.
  if (url.pathname.endsWith('/config.js')) {
    event.respondWith(
      fetch(request, {cache: 'no-store'})
        .catch(() => caches.open(CACHE).then(cache => cache.match(request)))
    );
    return;
  }

  // HTML aplikasi harus network-first agar perubahan terbaru tidak tertahan
  // oleh cache lama pada PWA yang sudah pernah dipasang.
  if (request.mode === 'navigate' || url.pathname.endsWith('/index.html')) {
    event.respondWith(
      fetch(request, {cache: 'no-store'}).then(response => {
        if (response && response.ok) {
          const copy = response.clone();
          caches.open(CACHE).then(cache => cache.put(request, copy));
        }
        return response;
      }).catch(() => caches.open(CACHE).then(cache => cache.match(request)).then(cached => {
        if (cached) return cached;
        return caches.match('./index.html');
      }))
    );
    return;
  }

  event.respondWith(
    caches.open(CACHE).then(cache =>
      cache.match(request).then(cached => cached || fetch(request))
    )
  );
});

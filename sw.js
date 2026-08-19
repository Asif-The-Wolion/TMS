// Teacher Planner — service worker
// NETWORK-FIRST: always tries to fetch the latest version when online, and
// only falls back to the cached copy if there's no connection. This matters
// a lot during active development — a "cache-first" strategy (the previous
// version of this file) would silently keep serving an old, possibly buggy
// build even after a new one was deployed, until the cache happened to
// refresh in the background. Network-first avoids that entirely: with any
// connection at all, you always get what's actually deployed right now.
//
// Prayer-time API calls are NEVER cached here — handled separately inside
// the app itself (localStorage-based cache), so Maghrib is never stale.

const CACHE_NAME = 'teacher-planner-v2'; // bumped so every existing install force-refreshes once
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-512-maskable.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  if (url.hostname.includes('aladhan.com')) return;
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        if (response && response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request)) // offline: fall back to last-known-good cache
  );
});

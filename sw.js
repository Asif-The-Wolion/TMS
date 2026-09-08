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

const CACHE_NAME = 'teacher-planner-v3'; // bumped again — this deploy also fixes the fetch() below to bypass the browser's own HTTP cache, not just the SW cache
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
    fetch(event.request, { cache: 'no-store' }) // bypass the browser's own HTTP disk-cache too, not just the SW cache — GitHub Pages sends Cache-Control: max-age on static files, so a plain fetch() can still return a stale response within that window even though this handler is "network-first". no-store forces a real network round-trip every time.
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

// Bump this string on any deploy where you want to force old cached assets
// out immediately (rare — network-first below already fetches the latest
// version whenever the person is online). This file only caches the app's
// own static files; it never touches localStorage or IndexedDB, so it has
// no bearing on whether user data survives an update.
const CACHE_NAME = "shift-priority-v3";

self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

// Network-first: always try to get the latest version; only fall back to the
// cache if the network request fails (e.g. fully offline). Only same-origin
// GETs are ever cached — a future cross-origin call (e.g. a cloud-backup API)
// must never be written into this cache.
self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    fetch(req)
      .then((response) => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
        }
        return response;
      })
      .catch(() => caches.match(req))
  );
});

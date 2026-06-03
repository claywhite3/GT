/* Service worker — caches the app shell so the PWA works offline.
   Bump CACHE_VERSION whenever you change any cached file. */
const CACHE_VERSION = 'gt-v34';
const SHELL = [
  'index.html',
  'store_type.html',
  'store_select.html',
  'flow_picker.html',
  'view_data.html',
  'category.html',
  'scan_camera.html',
  'scan_oos.html',
  'classify.html',
  'saved.html',
  'root_cause.html',
  'app.js',
  'config.js',
  'scanner.js',
  'scanner.css',
  'img/captana-logo.png',
  'img/scan-camera-hint.svg',
  'img/scan-oos-hint.svg',
  'manifest.json',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;        // never cache POSTs to the sheet
  const url = new URL(req.url);

  // Network-first for everything (own files + cross-origin): always try to
  // fetch the latest when online, fall back to cache when offline. This
  // prevents stale cached code (e.g. an old scanner.js) from masking deploys.
  event.respondWith(
    fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE_VERSION).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req))
  );
});

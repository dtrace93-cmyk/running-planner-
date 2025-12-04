const CACHE_NAME = 'rp-static-v3';
const ASSET_CACHE = [
  './',
  'index.html',
  'RP16.html',
  'index.js',
  'manifest.webmanifest',
  'icons/icon-192.png',
  'icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSET_CACHE))
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
    )
  );
});

self.addEventListener('fetch', (event) => {
  const { request } = event;

  // Skip Strava API calls entirely
  if (request.url.includes('/strava')) {
    return;
  }

  // Network-first for navigation/HTML
  const acceptHeader = request.headers.get('accept') || '';
  if (request.mode === 'navigate' || acceptHeader.includes('text/html')) {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
          return response;
        })
        .catch(() => caches.match('index.html'))
    );
    return;
  }

  // Cache-first for other GET requests
  if (request.method !== 'GET') return;

  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return fetch(request).then((response) => {
        if (response && response.status === 200) {
          const respCopy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, respCopy));
        }
        return response;
      });
    })
  );
});

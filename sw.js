const CACHE_NAME = 'rp-cache-v1';

// Relative asset paths for GitHub Pages subdirectory hosting
const CACHED_URLS = [
  './',
  'index.html',
  'RP16.html',
  'index.js',
  'manifest.webmanifest'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(CACHED_URLS);
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      );
    })
  );
});

self.addEventListener('fetch', (event) => {
  const request = event.request;

  // Don't cache Strava API calls
  if (request.url.includes('/strava')) {
    return;
  }

  event.respondWith(
    caches.match(request).then((cached) => {
      return cached || fetch(request);
    })
  );
});

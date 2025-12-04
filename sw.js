const CACHE_NAME = "rp-static-v4";

const FILES_TO_CACHE = [
  "/running-planner-/",
  "/running-planner-/index.html",
  "/running-planner-/RP16.html",
  "/running-planner-/manifest.webmanifest",
  "/running-planner-/sw.js",
  "/running-planner-/icons/icon-192.png",
  "/running-planner-/icons/icon-512.png",
  "/running-planner-/index.js"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(FILES_TO_CACHE);
    })
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  event.respondWith(
    caches.match(event.request).then((resp) => resp || fetch(event.request))
  );
});

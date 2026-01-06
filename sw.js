// Service Worker for PWA
const CACHE_NAME = "attendance-system-v1";
const urlsToCache = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "https://cdn.tailwindcss.com",
  "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css",
];

// Install event - cache resources
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log("Opened cache");
      return cache
        .addAll(
          urlsToCache.map((url) => {
            try {
              return new Request(url, { mode: "no-cors" });
            } catch (e) {
              return url;
            }
          })
        )
        .catch((err) => {
          console.log("Cache addAll failed:", err);
          // Cache index.html even if other resources fail
          return cache.add("./index.html");
        });
    })
  );
  self.skipWaiting();
});

// Activate event - clean up old caches
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log("Deleting old cache:", cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  return self.clients.claim();
});

// Fetch event - serve from cache, fallback to network
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Always fetch API requests from network (don't cache dynamic data)
  if (
    url.pathname.includes("script.google.com") ||
    url.pathname.includes("/macros/") ||
    event.request.method === "POST"
  ) {
    event.respondWith(fetch(event.request));
    return;
  }

  // Skip cross-origin requests that we can't cache
  if (
    !event.request.url.startsWith(self.location.origin) &&
    !event.request.url.includes("cdn.tailwindcss.com") &&
    !event.request.url.includes("cdnjs.cloudflare.com")
  ) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((response) => {
      // Return cached version or fetch from network
      if (response) {
        return response;
      }

      return fetch(event.request)
        .then((response) => {
          // Don't cache if not a valid response
          if (
            !response ||
            response.status !== 200 ||
            response.type !== "basic"
          ) {
            return response;
          }

          // Clone the response
          const responseToCache = response.clone();

          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });

          return response;
        })
        .catch(() => {
          // If fetch fails and it's a navigation request, return cached index.html
          if (event.request.mode === "navigate") {
            return caches.match("./index.html");
          }
        });
    })
  );
});

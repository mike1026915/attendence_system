// Service Worker for PWA
const CACHE_NAME = "attendance-system-v6";
const urlsToCache = [
  "./192x192.jpg",
  "./512x512.jpg",
  "https://cdn.tailwindcss.com",
  "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css",
];

// Check if a resource should be cached (CSS and images only)
function shouldCache(url) {
  const urlStr = url.toString().toLowerCase();
  return (
    /\.(css|jpg|jpeg|png|gif|webp|svg)(\?|$)/i.test(urlStr) ||
    urlStr.includes("cdnjs.cloudflare.com") ||
    urlStr.includes("cdn.tailwindcss.com")
  );
}

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

// Fetch event - cache CSS and images only, network-only for HTML
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Skip API requests - let browser handle them directly (no service worker overhead)
  if (
    url.pathname.includes("script.google.com") ||
    url.pathname.includes("/macros/") ||
    event.request.method === "POST"
  ) {
    return; // Don't intercept, let browser handle directly
  }

  // Skip cross-origin requests that we can't cache
  if (
    !event.request.url.startsWith(self.location.origin) &&
    !event.request.url.includes("cdn.tailwindcss.com") &&
    !event.request.url.includes("cdnjs.cloudflare.com")
  ) {
    return;
  }

  // Only cache CSS and images, everything else goes to network
  if (shouldCache(event.request.url)) {
    // Cache First strategy for CSS and images
    event.respondWith(
      caches.match(event.request).then((response) => {
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

            // Clone the response and cache it
            const responseToCache = response.clone();

            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache);
            });

            return response;
          })
          .catch((error) => {
            console.log("Fetch failed:", error);
            throw error;
          });
      })
    );
  } else {
    // Network Only strategy for HTML and other resources
    event.respondWith(fetch(event.request));
  }
});

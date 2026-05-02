self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open("resumefit-pro-v1").then((cache) => cache.addAll(["/", "/app-icon.svg", "/manifest.json"]))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") {
    return;
  }

  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});

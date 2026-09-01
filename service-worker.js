const CACHE_NAME = "masol-lecturas-v21";
const APP_SHELL = ["./", "./index.html", "./styles.css?v=21", "./auth.css?v=21", "./auth.js?v=21", "./app.js?v=21", "./config.js?v=21", "./manifest.webmanifest", "./icon.svg", "./icon-192.png", "./icon-512.png", "./logo-masol.png?v=20260828"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))));
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith(fetch(event.request).then((response) => {
    const copy = response.clone();
    caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
    return response;
  }).catch(() => caches.match(event.request).then((cached) => cached || caches.match("./index.html"))));
});

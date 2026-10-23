/* Enkel service worker: henter alltid nyeste versjon fra nettet, og bruker lagret kopi av selve appen bare når nettet er borte.
   Data fra SharePoint og innlogging går aldri via hurtigbufferen. */
const CACHE = "ordrebok-v7";
const SKALL = ["./", "index.html", "app.js", "store.js", "word.js", "config.js", "vendor/msal-browser.min.js", "vendor/pdf-lib.min.js", "vendor/xlsx.mini.min.js", "assets/logo.jpg", "assets/cmr.pdf", "manifest.webmanifest", "icons/icon-192.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SKALL)).catch(() => {})); self.skipWaiting(); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => { if (r.ok) { const kopi = r.clone(); caches.open(CACHE).then(c => c.put(e.request, kopi)); } return r; })
    .catch(() => caches.match(e.request).then(m => m || caches.match("index.html"))));
});

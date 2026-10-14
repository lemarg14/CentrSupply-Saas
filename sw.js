// Centr Supply Ops — rend l'app installable et utilisable hors ligne (page en cache).
const CACHE = 'centr-ops-v4';
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', 'index.html', 'manifest.webmanifest', 'icon.svg']).catch(() => {})));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
// Réseau d'abord pour la page (toujours la dernière version), cache en secours hors ligne.
// Les appels à Google Apps Script ne passent jamais par le cache.
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Modules de scan / QR (jsDelivr) : gardés en cache pour marcher hors ligne
  if (url.hostname === 'cdn.jsdelivr.net') {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => { const c = res.clone(); caches.open(CACHE).then(k => k.put(req, c)); return res; })));
    return;
  }
  if (url.origin !== location.origin) return;
  e.respondWith(fetch(req).then(res => {
    const copy = res.clone();
    caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
    return res;
  }).catch(() => caches.match(req).then(r => r || caches.match('index.html'))));
});

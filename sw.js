/* Склад — офлайн-кэш. При выпуске новой версии увеличь VERSION. */
const VERSION = 'sklad-v6';
const FILES = ['./', 'index.html', 'manifest.json', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'apple-touch-icon.png'];
const JSQR = 'https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js';

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(async c => {
    await c.addAll(FILES);
    try { await c.add(new Request(JSQR, {mode: 'cors'})); } catch (err) { /* без сети распознавание запасным способом недоступно */ }
  }));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith('sklad-') && k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('message', e => { if (e.data === 'skip') self.skipWaiting(); });
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.href === JSQR) {
    e.respondWith(caches.match(JSQR).then(r => r || fetch(req).then(res => {
      const copy = res.clone(); caches.open(VERSION).then(c => c.put(JSQR, copy)); return res;
    })));
    return;
  }
  if (url.origin !== self.location.origin) return;
  if (req.mode === 'navigate') { e.respondWith(caches.match('index.html').then(r => r || fetch(req))); return; }
  e.respondWith(caches.match(req, {ignoreSearch: true}).then(r => r || fetch(req)));
});

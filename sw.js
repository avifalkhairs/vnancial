/* Vnancial service worker: the app opens instantly and works offline; data always comes fresh from the network. */
const CACHE = 'vnancial-v1.0.0';
const SHELL = ['./', 'index.html', 'app.css', 'app.js', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.hostname.endsWith('google.com') || u.hostname.endsWith('googleusercontent.com')) return; // API: never cached
  if (u.origin === location.origin) {
    // app files: cache first, refresh in the background
    e.respondWith(caches.match(e.request).then(hit => {
      const net = fetch(e.request).then(r => { if (r.ok) caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; }).catch(() => hit);
      return hit || net;
    }));
  } else if (u.hostname.includes('fonts.')) {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(r => { caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; })));
  }
});

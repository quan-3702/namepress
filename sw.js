const CACHE = 'namepress-v9';
const ASSETS = ['./', 'index.html', 'help.html', 'privacy.html', 'style.css', 'app.js', 'manifest.json', 'icon-192.png', 'icon-512.png', 'icon-512-maskable.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS.map((u) => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) =>
    Promise.all(keys.filter((k) => k !== CACHE && k !== 'namepress-libs').map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  if (url.hostname === 'cdn.jsdelivr.net') { // the Excel reader: keep a copy after first use so Excel import works offline
    e.respondWith(caches.open('namepress-libs').then(async (c) => { const hit = await c.match(e.request); if (hit) return hit; const res = await fetch(e.request); if (res.ok || res.type === 'opaque') c.put(e.request, res.clone()); return res; }));
    return;
  }
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request)));
});

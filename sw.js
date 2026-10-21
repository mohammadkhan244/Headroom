// headroom-v3: network-first for HTML, stale-while-revalidate for assets
const CACHE = 'headroom-v3';
const SHELL = ['/', '/index.html', '/manifest.json', '/icons/icon-192.svg', '/icons/icon-512.svg'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // Cross-origin (fonts, CDN): network with cache fallback
  if (url.origin !== self.location.origin) {
    e.respondWith(
      fetch(e.request).then(r => {
        if (r.ok) caches.open(CACHE).then(c => c.put(e.request, r.clone()));
        return r;
      }).catch(() => caches.match(e.request))
    );
    return;
  }

  // index.html / root: always network-first so stale code is never served
  if (url.pathname === '/' || url.pathname === '/index.html') {
    e.respondWith(
      fetch(new Request(e.request, { cache: 'no-store' }))
        .then(r => {
          if (r.ok) caches.open(CACHE).then(c => c.put(e.request, r.clone()));
          return r;
        })
        .catch(() => caches.match(e.request)) // offline fallback
    );
    return;
  }

  // Never cache /api/* — auth-gated, must always be fresh
  if (url.pathname.startsWith('/api/')) return;

  // Other same-origin assets: stale-while-revalidate
  e.respondWith(
    caches.open(CACHE).then(cache =>
      cache.match(e.request).then(cached => {
        const network = fetch(e.request).then(fresh => {
          if (fresh.ok) cache.put(e.request, fresh.clone());
          return fresh;
        }).catch(() => null);
        return cached || network;
      })
    )
  );
});

self.addEventListener('message', e => {
  if (e.data?.type === 'SKIP_WAITING') self.skipWaiting();
});

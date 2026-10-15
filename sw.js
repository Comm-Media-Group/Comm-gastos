// Bump VERSION on every deploy — it forces phones to fetch the new files.
const VERSION = 'comm-gastos-v1.3.0';
const CORE = ['./', './index.html', './styles.css', './app.js', './config.js', './mail.js', './xlsx-builder.js', './xlsx-assets.js', './manifest.json',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png', './icons/favicon-32.png', './icons/apple-touch-icon.png'];
const CDN = ['https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(VERSION);
    await c.addAll(CORE);
    await Promise.all(CDN.map(u => c.add(u).catch(() => { }))); // optional; also cached at runtime
  })());
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION && k !== 'cmg-share').map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  // Android "Share → COMM Gastos": park the shared files in a cache; the app turns them into "Por completar" drafts on open
  if (req.method === 'POST' && new URL(req.url).pathname.endsWith('/share-target')) {
    e.respondWith((async () => {
      try {
        const fd = await req.formData(), c = await caches.open('cmg-share');
        let i = 0;
        for (const f of fd.getAll('files')) {
          if (!f || !f.size) continue;
          await c.put(new URL('./_shared/' + Date.now() + '-' + (i++), self.registration.scope).href,
            new Response(f, { headers: { 'Content-Type': f.type || 'application/octet-stream', 'X-Name': encodeURIComponent(f.name || '') } }));
        }
      } catch (err) { /* fall through to the app anyway */ }
      return Response.redirect(new URL('./index.html?shared=1', self.registration.scope).href, 303);
    })());
    return;
  }
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Google sign-in / Gmail API must always hit the network
  if (/(^|\.)(accounts\.google\.com|googleapis\.com)$/.test(url.hostname) && !/fonts\./.test(url.hostname)) return;

  const sameOrigin = url.origin === self.location.origin;
  if (sameOrigin) {
    // network-first: always pick up new code when online, fall back to cache offline
    e.respondWith(fetch(req, { cache: 'no-cache' }).then(res => { // revalidate with the server every time (not the browser's 10-min HTTP cache)
      if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('./index.html'))));
  } else {
    // fonts + JSZip: cache-first, filled at runtime
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
      if (res.ok || res.type === 'opaque') { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
      return res;
    })));
  }
});

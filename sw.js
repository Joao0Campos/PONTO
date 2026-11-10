const C = 'ponto-vic-v12';
const ASSETS = ['./', 'index.html', 'manifest.json', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'apple-touch-icon.png', 'img/vic-marca.png', 'img/vic-logo.png', 'img/vic-logo-escuro.png', 'css/app.css',
  'js/core.js', 'js/importar.js', 'js/cartao.js', 'js/app.js', 'vendor/exceljs.min.js'];
self.addEventListener('install', e => e.waitUntil(caches.open(C).then(c => c.addAll(ASSETS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k)))).then(() => self.clients.claim())));
// rede primeiro, sempre revalidando no servidor (evita HTML novo com CSS/JS velho do cache HTTP); cache como reserva offline
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  // navegação não aceita opções no fetch(request, init): refaz o pedido pela URL
  const req = e.request.mode === 'navigate' ? new Request(e.request.url, { cache: 'no-cache', credentials: 'same-origin' }) : e.request;
  e.respondWith(fetch(req, e.request.mode === 'navigate' ? undefined : { cache: 'no-cache' }).then(r => { if (r.ok) { const c = r.clone(); caches.open(C).then(x => x.put(e.request, c)); } return r; })
    .catch(() => caches.match(e.request, { ignoreSearch: true })));
});

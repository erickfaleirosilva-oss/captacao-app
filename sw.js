// Your Vacation — Service Worker (PWA offline)
// Estratégia: network-first com fallback cache. Ao detectar update, notifica o app para recarregar.
const CACHE_NAME = 'yv-captacao-v18';
const ASSETS = [
  '/captacao-app/',
  '/captacao-app/index.html',
  '/captacao-app/manifest.json',
  '/captacao-app/icon-192.png',
  '/captacao-app/icon-512.png',
];

// Instalação
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Ativação — remove caches antigos e avisa todos os clientes para recarregar
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => {
      // Notifica todos os clientes abertos que há nova versão disponível
      return self.clients.matchAll({ type: 'window' }).then(clients => {
        clients.forEach(client => client.postMessage({ type: 'APP_UPDATED' }));
      });
    })
  );
  self.clients.claim();
});

// Fetch — network-first: tenta a rede, cai no cache se offline
self.addEventListener('fetch', e => {
  // Não intercepta chamadas ao Google Apps Script
  if (e.request.url.includes('script.google.com')) return;
  // Só intercepta assets do app
  if (!e.request.url.includes('/captacao-app/') && !e.request.url.endsWith('/captacao-app')) return;

  e.respondWith(
    fetch(e.request).then(response => {
      // Atualiza cache com versão nova
      if (response.ok) {
        caches.open(CACHE_NAME).then(cache => cache.put(e.request, response.clone()));
      }
      return response;
    }).catch(() => {
      // Sem rede — serve do cache
      return caches.match(e.request);
    })
  );
});

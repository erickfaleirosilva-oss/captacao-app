// Your Vacation — Service Worker (PWA offline)
const CACHE_NAME = 'yv-captacao-v1';
const ASSETS = [
  '/captacao-app/',
  '/captacao-app/index.html',
  '/captacao-app/manifest.json',
];

// Instalação — faz cache dos arquivos principais
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Ativação — remove caches antigos
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch — serve do cache se offline, senão busca na rede e atualiza cache
self.addEventListener('fetch', e => {
  // Não intercepta chamadas ao Google Apps Script (webhook)
  if (e.request.url.includes('script.google.com')) return;

  e.respondWith(
    caches.match(e.request).then(cached => {
      const network = fetch(e.request).then(response => {
        // Atualiza cache com versão nova se for um asset do app
        if (response.ok && e.request.url.includes('/captacao-app/')) {
          caches.open(CACHE_NAME).then(cache => cache.put(e.request, response.clone()));
        }
        return response;
      }).catch(() => null);

      // Retorna cache imediatamente se disponível, senão aguarda rede
      return cached || network;
    })
  );
});

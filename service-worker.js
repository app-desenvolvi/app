/** 🚀 SERVICE WORKER PASSIVO DO PORTAL DESENVOLVI */
const CACHE_NAME = 'desenvolvi-portal-v1';

// Ativa o SW imediatamente
self.addEventListener('install', function(event) {
  self.skipWaiting();
});

// Toma controle da página na hora
self.addEventListener('activate', function(event) {
  event.waitUntil(self.clients.claim());
});

// Intercepta e repassa requisições de rede
self.addEventListener('fetch', function(event) {
  event.respondWith(
    fetch(event.request).catch(function() {
      return caches.match(event.request);
    })
  );
});

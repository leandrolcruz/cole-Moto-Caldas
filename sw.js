var CACHE_NAME = 'inv-motocaldas-v25';
var urlsToCache = [
  './',
  './index.html',
  './manifest.json',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Outfit:wght@300;400;500;600;700;800;900&display=swap'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(urlsToCache);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(
        names.filter(function(name) { return name !== CACHE_NAME; })
             .map(function(name) { return caches.delete(name); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(event) {
  // API do Apps Script: NUNCA interceptar — resposta transiente cacheada envenena
  // o login/dados ("Falha ao falar com o backend" persistente)
  if (event.request.url.indexOf('script.google') !== -1 ||
      event.request.url.indexOf('googleusercontent') !== -1) {
    return;
  }
  // Catálogo embutido e dashboard: sempre da rede (dashboard é ferramenta de PC
  // online — cache-first segurava updates até o bump de versão)
  if (event.request.url.indexOf('catalogo.') !== -1 ||
      event.request.url.indexOf('dashboard') !== -1) {
    event.respondWith(fetch(event.request));
    return;
  }
  event.respondWith(
    caches.match(event.request).then(function(response) {
      if (response) return response;
      return fetch(event.request).then(function(fetchResponse) {
        if (fetchResponse && fetchResponse.status === 200) {
          var responseClone = fetchResponse.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(event.request, responseClone);
          });
        }
        return fetchResponse;
      }).catch(function() {
        // Offline fallback
        return caches.match('./index.html');
      });
    })
  );
});

/* Service worker — Panel de Seguimiento (caché para uso sin conexión) */
var CACHE = 'panel-seguimiento-v7'; /* subir versión cuando cambien app.js/styles.css para que las instalaciones existentes actualicen */
var ARCHIVOS = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './xlsx.full.min.js',
  './manifest.webmanifest',
  './icon.svg',
  './icon-512.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (cache) { return cache.addAll(ARCHIVOS); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then(function (resp) {
      return resp || fetch(e.request).catch(function () { return caches.match('./index.html'); });
    })
  );
});

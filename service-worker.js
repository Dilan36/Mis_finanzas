// Service Worker — Control de Gastos
// Cachea el "cascarón" de la app (HTML/CSS/JS/íconos) para que abra sin
// internet. Los datos (Sheets) SIEMPRE necesitan conexión; esto solo
// evita la pantalla en blanco cuando no hay señal.

const CACHE_NAME = 'control-gastos-v3';
const ASSETS = [
  './',
  './index.html',
  './config.js',
  './app.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-512-maskable.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Solo controlamos el cascarón de la app (GET, mismo origen).
  // Las llamadas a la API de Google (POST) pasan de largo, siempre por red.
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const fetchPromise = fetch(event.request)
        .then((networkResponse) => {
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, networkResponse.clone()));
          return networkResponse;
        })
        .catch(() => cached);
      return cached || fetchPromise;
    })
  );
});
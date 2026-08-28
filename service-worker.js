// Service worker mínimo - solo lo necesario para que la página cumpla
// los requisitos de "PWA instalable" que pide TWA. No hace caché
// agresivo a propósito, porque la app real vive en Apps Script y
// siempre tiene que traer los datos frescos (pedidos, productos, etc).

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});

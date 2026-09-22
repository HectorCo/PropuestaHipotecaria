// Service Worker minimalista.
// No intercepta peticiones: todo va directo a la red.
// Su única misión es habilitar la instalación PWA y limpiar cachés antiguas.

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// Sin handler de 'fetch' → el navegador hace peticiones normales a la red.
// propuesta-pwa/sw.js

// 1. Cambia la versión de la caché cada vez que hagas un despliegue importante.
//    Esto fuerza al navegador a ver este archivo como nuevo.
const CACHE = 'propuesta-hipotecaria-v6'; 

const ASSETS = [
  './',
  './manifest.json',
  './icon.svg',
  'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  // 2. Estrategia "Network-First" para archivos críticos.
  //    Intenta buscar en la red primero. Si falla, usa la caché.
  //    Esto se aplica a la navegación, HTML, JS y CSS.
  if (
    event.request.mode === 'navigate' ||
    url.pathname.endsWith('.html') ||
    url.pathname.endsWith('.js') ||
    url.pathname.endsWith('.css')
  ) {
    event.respondWith(
      fetch(event.request)
        .then((res) => {
          // Guarda una copia en caché para futuras peticiones offline.
          const clone = res.clone();
          caches.open(CACHE).then((cache) => cache.put(event.request, clone));
          return res;
        })
        .catch(() => {
          // Si falla la red, intenta servir desde la caché.
          return caches.match(event.request).then((r) => r || caches.match('./index.html'));
        })
    );
    return;
  }

  // 3. Estrategia "Cache-First" para el resto (iconos, librerías CDN).
  //    Estos cambian muy poco, así que priorizamos la velocidad de la caché.
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((res) => {
        if (res && res.ok && (res.type === 'basic' || res.type === 'cors')) {
          const clone = res.clone();
          caches.open(CACHE).then((cache) => cache.put(event.request, clone));
        }
        return res;
      });
    })
  );
});

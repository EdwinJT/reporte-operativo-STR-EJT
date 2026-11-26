/**
 * Service Worker Offline-First — Control de Horómetros y Producción STRACON
 * Precachea el App Shell completo (HTML, JS, CSS, íconos y manifest) para funcionar 100% sin señal.
 */

const CACHE_VERSION = 'v1.3.3-supabase-live';
const APP_SHELL_CACHE = `stracon-produccion-appshell-${CACHE_VERSION}`;

const PRECACHE_ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './supabase_setup.sql',
  './public/pwa-192x192.png',
  './public/pwa-512x512.png',
  './public/apple-touch-icon.png',
  './src/main.js',
  './src/App.js',
  './src/index.css',
  './src/config/supabaseClient.js',
  './src/db/localDb.js',
  './src/db/storageManager.js',
  './src/data/demoSeed.js',
  './src/services/pwaManager.js',
  './src/services/connectivity.js',
  './src/services/catalogService.js',
  './src/services/productionService.js',
  './src/services/syncQueue.js',
  './src/services/excelExporter.js',
  './src/utils/datePeru.js',
  './src/utils/codeNormalizer.js',
  './src/components/CapturaTurno.js',
  './src/components/MatrizControl.js',
  './src/components/HistorialCatalogo.js',
  './src/components/GerenciaDashboard.js',
  './src/components/ConfigModal.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(APP_SHELL_CACHE).then(async (cache) => {
      for (const asset of PRECACHE_ASSETS) {
        try {
          const response = await fetch(asset, { cache: 'reload' });
          if (response && response.ok) {
            await cache.put(asset, response);
          }
        } catch {
          // Ignorar fallo individual offline para no abortar la instalación
        }
      }
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys.map((key) => {
            if (key.startsWith('stracon-produccion-appshell-') && key !== APP_SHELL_CACHE) {
              return caches.delete(key);
            }
            return Promise.resolve();
          })
        )
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (event) => {
  if (!event.data) return;

  if (event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }

  if (event.data.type === 'PRECACHING_VERIFY') {
    event.waitUntil(
      caches.open(APP_SHELL_CACHE).then(async (cache) => {
        for (const asset of PRECACHE_ASSETS) {
          try {
            const existing = await cache.match(asset);
            if (!existing) {
              const res = await fetch(asset);
              if (res && res.ok) await cache.put(asset, res);
            }
          } catch {
            // ignore offline
          }
        }
      })
    );
  }
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then((networkRes) => {
          const copy = networkRes.clone();
          caches.open(APP_SHELL_CACHE).then((cache) => cache.put('./index.html', copy));
          return networkRes;
        })
        .catch(async () => {
          const cachedIndex =
            (await caches.match(request)) ||
            (await caches.match('./index.html')) ||
            (await caches.match('./'));
          return cachedIndex || Response.error();
        })
    );
    return;
  }

  event.respondWith(
    fetch(request)
      .then((networkResponse) => {
        if (networkResponse && networkResponse.ok) {
          const copy = networkResponse.clone();
          caches.open(APP_SHELL_CACHE).then((cache) => cache.put(request, copy));
        }
        return networkResponse;
      })
      .catch(() => caches.match(request))
  );
});

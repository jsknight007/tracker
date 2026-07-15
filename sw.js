const APP_VERSION = '3.1.0';
const CACHE_NAME = 'shopping-list-v' + APP_VERSION;
const SHELL_FILES = ['index.html', 'kids.html', 'manifest.json', 'kids.json'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});

self.addEventListener('message', (event) => {
  if (event.data === 'getVersion') {
    event.source.postMessage({ type: 'version', version: APP_VERSION });
  }
});

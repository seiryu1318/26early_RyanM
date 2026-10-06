const CACHE_NAMESPACE = 'admissions-result-sync';
const CACHE_VERSION = 'v7';
const SCOPE_URL = new URL(self.registration.scope);
const SCOPE_PATH = SCOPE_URL.pathname.endsWith('/')
  ? SCOPE_URL.pathname
  : `${SCOPE_URL.pathname}/`;
const SCOPE_KEY = SCOPE_PATH
  .replace(/^\/+|\/+$/g, '')
  .replace(/[^a-zA-Z0-9._-]+/g, '-') || 'root';
const CACHE_PREFIX = `${CACHE_NAMESPACE}:${SCOPE_KEY}:`;
const CACHE_NAME = `${CACHE_PREFIX}shell:${CACHE_VERSION}`;
const APP_SHELL = [
  '',
  'index.html',
  'manifest.webmanifest',
  'icons/app-icon-48.png',
  'icons/app-icon-180.png',
  'icons/app-icon-192.png',
  'icons/app-icon-512.png',
].map(path => new URL(path, SCOPE_URL).href);
const OFFLINE_URL = new URL('index.html', SCOPE_URL).href;

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys
          .filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE_NAME)
          .map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const requestUrl = new URL(event.request.url);
  const isOwnedRequest = requestUrl.origin === SCOPE_URL.origin
    && requestUrl.pathname.startsWith(SCOPE_PATH);
  if (event.request.method !== 'GET' || !isOwnedRequest) return;
  event.respondWith(
    fetch(event.request)
      .then(async response => {
        if (response.ok) {
          const copy = response.clone();
          const cache = await caches.open(CACHE_NAME);
          await cache.put(event.request, copy);
        }
        return response;
      })
      .catch(async () => {
        const cache = await caches.open(CACHE_NAME);
        return (await cache.match(event.request)) || cache.match(OFFLINE_URL);
      })
  );
});

// SW v36 — self-destruct: unregisters all old SWs and clears all caches,
// then re-registers clean. Fixes stuck SW on Android Chrome.
var CACHE_NAME = 'plusone-v36';
var urlsToCache = ['/offices/index.html', '/offices/manifest.json', '/offices/icon.png'];

self.addEventListener('install', function(event) {
  self.skipWaiting();
  event.waitUntil(
    // Wipe ALL caches first, then re-cache fresh
    caches.keys().then(function(names) {
      return Promise.all(names.map(function(name) { return caches.delete(name); }));
    }).then(function() {
      return caches.open(CACHE_NAME).then(function(cache) {
        return Promise.all(urlsToCache.map(function(url) {
          return fetch(new Request(url, { cache: 'no-store' }))
            .then(function(res) { if (res.ok) cache.put(url, res); })
            .catch(function() {});
        }));
      });
    })
  );
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(
        names.filter(function(name) { return name !== CACHE_NAME; })
          .map(function(name) { return caches.delete(name); })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function(event) {
  var url = event.request.url;

  // Never intercept Google Script calls (JSONP)
  if (url.indexOf('script.google.com') !== -1 || url.indexOf('script.googleusercontent.com') !== -1) {
    return;
  }

  // HTML navigation: always network-first, never serve stale HTML
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(new Request(event.request, { cache: 'no-store' }))
        .then(function(response) {
          var copy = response.clone();
          caches.open(CACHE_NAME).then(function(cache) { cache.put(event.request, copy); });
          return response;
        })
        .catch(function() {
          return caches.match(event.request)
            .then(function(cached) { return cached || caches.match('/offices/index.html'); });
        })
    );
    return;
  }

  // Everything else: network-first, cache fallback
  event.respondWith(
    fetch(event.request).then(function(response) {
      var copy = response.clone();
      caches.open(CACHE_NAME).then(function(cache) {
        if (event.request.method === 'GET') cache.put(event.request, copy);
      });
      return response;
    }).catch(function() {
      return caches.match(event.request).then(function(cached) {
        return cached || Response.error();
      });
    })
  );
});

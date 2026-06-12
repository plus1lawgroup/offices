var CACHE_NAME = 'plusone-v27';
var urlsToCache = ['/offices/', '/offices/index.html', '/offices/dashboard_pwa.html', '/offices/manifest.json', '/offices/icon.png'];

self.addEventListener('install', function(event) {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(urlsToCache);
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
  // Never intercept Google Script calls (fetch or JSONP)
  if (url.indexOf('script.google.com') !== -1 || url.indexOf('script.googleusercontent.com') !== -1) {
    return;
  }
  event.respondWith(
    fetch(event.request).then(function(response) {
      var copy = response.clone();
      caches.open(CACHE_NAME).then(function(cache) {
        if (event.request.method === 'GET') cache.put(event.request, copy);
      });
      return response;
    }).catch(function() {
      return caches.match(event.request);
    })
  );
});

// E-Laundry PWA Service Worker – Dynamic Base Path

var CACHE_NAME = 'elaundry-shell-v2';

// Ambil base path dari lokasi sw.js (misal: /laundry/ atau /)
var basePath = self.location.pathname.replace('sw.js', '');

// Asset statis yang di-cache (relative terhadap basePath)
var SHELL_ASSETS = [
  basePath,
  basePath + 'index.html',
  basePath + 'manifest.json',
  basePath + 'config.js',
  basePath + 'icon-192.svg',
  basePath + 'icon-512.svg'
];

// Install
self.addEventListener('install', function(event) {
  console.log('[SW] Install, basePath:', basePath);
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(SHELL_ASSETS);
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

// Activate – hapus cache lama
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(key) {
          return key !== CACHE_NAME;
        }).map(function(key) {
          console.log('[SW] Hapus cache lama:', key);
          return caches.delete(key);
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// Fetch strategy
self.addEventListener('fetch', function(event) {
  var url = event.request.url;
  var req = event.request;

  // JANGAN cache Google Apps Script
  if (url.includes('script.google.com')) {
    event.respondWith(fetch(req));
    return;
  }

  // Jangan cache juga jika request ke Google Fonts / CDN eksternal? 
  // Biar network first, tapi kita tetap coba cache untuk offline
  if (url.includes('fonts.googleapis.com') || url.includes('fonts.gstatic.com') || url.includes('cdn.jsdelivr.net')) {
    event.respondWith(
      fetch(req).then(function(resp) {
        var clone = resp.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(req, clone);
        });
        return resp;
      }).catch(function() {
        return caches.match(req);
      })
    );
    return;
  }

  // Shell assets – Cache first
  event.respondWith(
    caches.match(req).then(function(cached) {
      if (cached) return cached;
      return fetch(req).then(function(resp) {
        if (resp && resp.status === 200 && resp.type === 'basic') {
          var clone = resp.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(req, clone);
          });
        }
        return resp;
      }).catch(function() {
        // Offline fallback ke index.html
        if (req.mode === 'navigate') {
          return caches.match(basePath + 'index.html');
        }
      });
    })
  );
});

// Message: skip waiting atau clear cache
self.addEventListener('message', function(event) {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  if (event.data && event.data.type === 'CLEAR_CACHE') {
    caches.delete(CACHE_NAME).then(function() {
      console.log('[SW] Cache dihapus manual');
    });
  }
});

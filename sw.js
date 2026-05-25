// ============================================================
// E-LAUNDRY MANAGEMENT SYSTEM — Service Worker
// sw.js — Cache shell statis GitHub Pages
// ============================================================

var CACHE_NAME = 'elaundry-shell-v1';

// Asset statis yang di-cache (hanya file GitHub Pages, bukan GAS)
var SHELL_ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icon-192.svg',
  '/icon-512.svg',
  // Tambahkan aset statis lain jika ada (misal CSS/JS terpisah)
];

// ── INSTALL: cache shell assets ──
self.addEventListener('install', function(event) {
  console.log('[SW] Install — caching shell assets');
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(SHELL_ASSETS);
    }).then(function() {
      // Langsung aktif tanpa menunggu tab lama tutup
      return self.skipWaiting();
    })
  );
});

// ── ACTIVATE: hapus cache lama ──
self.addEventListener('activate', function(event) {
  console.log('[SW] Activate — membersihkan cache lama');
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

// ── FETCH: strategi per tipe request ──
self.addEventListener('fetch', function(event) {
  var url = event.request.url;
  var req = event.request;

  // 1. Request ke GAS (script.google.com) — ALWAYS network, jangan di-cache
  if (url.includes('script.google.com')) {
    event.respondWith(fetch(req));
    return;
  }

  // 2. Request ke Google Fonts & CDN eksternal — network first, fallback cache
  if (
    url.includes('fonts.googleapis.com') ||
    url.includes('fonts.gstatic.com') ||
    url.includes('cdn.jsdelivr.net')
  ) {
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

  // 3. Shell assets (index.html, manifest, icon) — cache first, fallback network
  event.respondWith(
    caches.match(req).then(function(cached) {
      if (cached) return cached;

      return fetch(req).then(function(resp) {
        // Cache hanya response OK dari origin yang sama
        if (resp && resp.status === 200 && resp.type === 'basic') {
          var clone = resp.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(req, clone);
          });
        }
        return resp;
      }).catch(function() {
        // Offline fallback ke index.html untuk navigasi
        if (req.mode === 'navigate') {
          return caches.match('/index.html');
        }
      });
    })
  );
});

// ── MESSAGE: update cache manual jika diperlukan ──
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

/* Naim Baba Yarış — çevrimdışı kabuk + son gün/istatistik yanıtları */
const SURUM = 'naim-at-v4';
const KABUK = [
  './',
  './manifest.webmanifest',
  './icons/icon-180.png',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './logo.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(SURUM).then(function (c) {
      return Promise.all(KABUK.map(function (u) {
        return c.add(u).catch(function () {});
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (ks) {
      return Promise.all(ks.map(function (k) {
        if (k !== SURUM) return caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

function anahtar(url) {
  const u = new URL(url);
  // API isteklerinde tarih/koşu parametresi anahtara girmeli; yoksa bütün
  // günler tek önbellek girdisini paylaşır ve "hep Salı" görünür.
  const q = u.pathname.indexOf('/api/') === 0 ? u.search : '';
  return u.origin + u.pathname + q;
}

self.addEventListener('fetch', function (e) {
  const req = e.request;
  if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (u.origin !== location.origin) return;

  const json = /\.json$/i.test(u.pathname) || u.pathname.indexOf('/veri/') >= 0;
  const sayfa = req.mode === 'navigate';

  if (json || sayfa) {
    // Ağ önce: yeni veri geldiyse göster ve sakla; çevrimdışıysa son iyi kopya.
    e.respondWith(
      fetch(req).then(function (r) {
        if (r && r.ok) {
          const kopya = r.clone();
          caches.open(SURUM).then(function (c) { c.put(anahtar(req.url), kopya); });
        }
        return r;
      }).catch(function () {
        return caches.match(anahtar(req.url)).then(function (c) {
          return c || caches.match(new URL('./', self.location.href).href);
        });
      })
    );
  } else {
    e.respondWith(
      caches.match(anahtar(req.url)).then(function (c) {
        return c || fetch(req).then(function (r) {
          if (r && r.ok) {
            const kopya = r.clone();
            caches.open(SURUM).then(function (cc) { cc.put(anahtar(req.url), kopya); });
          }
          return r;
        });
      })
    );
  }
});

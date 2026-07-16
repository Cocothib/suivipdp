const CACHE_NAME = 'suivi-pdp-v188';
// Cache séparé pour les libs CDN : leurs URLs sont versionnées (immuables),
// inutile de les re-télécharger à chaque bump du cache applicatif.
// N'incrémenter LIBS_CACHE que si la liste CDN_ASSETS change.
const LIBS_CACHE = 'suivi-pdp-libs-1';

// App shell local
const ASSETS = [
  './',
  './index.html',
  './manifest.json'
];

// Libs CDN précachées (P0.2) : évite l'écran blanc au lancement en chantier.
// Dexie + Bootstrap sont prioritaires (l'app ne démarre pas sans eux). Les autres
// (export PDF/DOCX/XLSX) sont précachés best-effort. Toute URL absente du précache
// reste servie par le mécanisme cache-first/stale-while-revalidate ci-dessous.
const CDN_ASSETS = [
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js',
  'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css',
  'https://unpkg.com/dexie@3.2.4/dist/dexie.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://unpkg.com/docx@8.5.0/build/index.umd.js',
  'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'
];

// Une URL CDN va dans le cache libs (stable), le reste dans le cache applicatif (versionné).
function _cacheFor(url) {
  return /^https:\/\/(cdn\.jsdelivr\.net|unpkg\.com|cdnjs\.cloudflare\.com|alcdn\.msauth\.net)\//.test(url)
    ? LIBS_CACHE : CACHE_NAME;
}

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    // App shell : obligatoire (échec = install échoue, comportement voulu).
    const shell = await caches.open(CACHE_NAME);
    await shell.addAll(ASSETS);
    // Libs CDN : best-effort, une indispo CDN ne doit pas bloquer l'install.
    const libs = await caches.open(LIBS_CACHE);
    await Promise.all(CDN_ASSETS.map(async u => {
      if (await libs.match(u)) return; // déjà en cache : pas de re-téléchargement
      await libs.add(new Request(u, { mode: 'cors' })).catch(() => { /* précache best-effort */ });
    }));
  })());
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME && k !== LIBS_CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Détermine si une requête est un asset statique cacheable (app shell + libs CDN).
// On NE cache QUE les GET d'assets ; tout le reste (proxy PHP, Graph, Sellsy, POST/PUT)
// passe en réseau direct pour ne JAMAIS servir des données dynamiques depuis le cache.
function _isStaticAsset(url) {
  // Libs CDN connues
  if (/^https:\/\/(cdn\.jsdelivr\.net|unpkg\.com|cdnjs\.cloudflare\.com|alcdn\.msauth\.net)\//.test(url)) return true;
  // Assets locaux (même origine) hors API : html, css, js, json statique, images, fontes
  return /\.(html|css|js|json|png|jpg|jpeg|gif|svg|webp|ico|woff2?|ttf|eot)(\?.*)?$/i.test(url);
}

self.addEventListener('fetch', event => {
  const req = event.request;

  // 1) NE JAMAIS intercepter autre chose que GET : les écritures (POST/PUT/PATCH/DELETE)
  //    vers le proxy PHP / Graph / Sellsy doivent toujours passer par le réseau réel.
  if (req.method !== 'GET') return; // laisse le navigateur gérer normalement

  const url = req.url;

  // 2) Ne pas cacher les appels dynamiques (proxy PHP, Graph, Sellsy) même en GET :
  //    risque de servir des données périmées / incohérentes.
  const isDynamic = /\.php(\?.*)?$/i.test(url)
    || /graph\.microsoft\.com|login\.microsoftonline\.com|sellsy/i.test(url)
    || /[?&]api=|sellsy-proxy/i.test(url);
  if (isDynamic) return; // réseau direct, pas de cache

  // 3) Assets statiques : CACHE-FIRST + stale-while-revalidate.
  if (_isStaticAsset(url)) {
    event.respondWith(
      caches.match(req).then(cached => {
        const network = fetch(req).then(response => {
          // Ne met en cache que les réponses exploitables (ok ou opaque CDN).
          if (response && (response.ok || response.type === 'opaque')) {
            const clone = response.clone();
            caches.open(_cacheFor(url)).then(cache => cache.put(req, clone)).catch(() => {});
          }
          return response;
        }).catch(() => cached); // hors-ligne : retombe sur le cache
        // Cache d'abord (affichage instantané), rafraîchit en arrière-plan.
        return cached || network;
      })
    );
    return;
  }

  // 4) Reste (GET non statique non dynamique) : network-first avec repli cache.
  event.respondWith(
    fetch(req)
      .then(response => {
        if (response && response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(req, clone)).catch(() => {});
        }
        return response;
      })
      .catch(() => caches.match(req))
  );
});

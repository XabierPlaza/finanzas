/* Funcionamiento sin conexión de la app.
 * Guarda en caché los archivos de la app (no los datos: esos están en IndexedDB y, si se elige, en Drive).
 * Estrategia: responde desde la caché al instante y actualiza en segundo plano.
 * VERSION y APP los rellena scripts/publicar.mjs en cada publicación (versión con fecha y lista real de archivos).
 * Las llamadas a Google no se tocan. */
const VERSION = "finanzas-20261026-1221";
const APP = ["./","assets/index-De3MCHpz.css","assets/index-DHglBPir.js","iconos/apple-touch-icon.png","iconos/icon-192.png","iconos/icon-512.png","iconos/icon-maskable-512.png","index.html","manifest.webmanifest","vendor/exceljs.min.js","vendor/xlsx.full.min.js"];

self.addEventListener("install", e => {
  // cache: "reload": baja los archivos del servidor, no de la caché HTTP del navegador (podría tener un index.html antiguo)
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(APP.map(u => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;  // Google y el resto, directos a la red
  e.respondWith(caches.open(VERSION).then(async cache => {
    const cached = await cache.match(e.request, { ignoreSearch: true });
    // los archivos de assets/ llevan huella en el nombre y no cambian nunca; el resto (index.html…) se revalida siempre
    const req = url.pathname.includes("/assets/") ? e.request : new Request(e.request.url, { cache: "no-cache" });
    const net = fetch(req).then(r => { if (r.ok) cache.put(e.request, r.clone()); return r; }).catch(() => null);
    return cached || (await net) || (e.request.mode === "navigate" ? cache.match("index.html") : Response.error());
  }));
});

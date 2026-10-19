/* Funcionamiento sin conexión de la app del móvil.
 * Guarda en caché los archivos de la app (no los datos: esos están en IndexedDB y, si se elige, en Drive).
 * Estrategia: responde desde la caché al instante y actualiza en segundo plano; al publicar una versión
 * nueva basta con cambiar VERSION. Las llamadas a Google no se tocan. */
const VERSION = "finanzas-v14";
const APP = ["./", "index.html", "datos-local.js", "drive.js", "xlsx.full.min.js", "exceljs.min.js", "manifest.webmanifest",
  "iconos/icon-192.png", "iconos/icon-512.png", "iconos/apple-touch-icon.png", "iconos/icon-maskable-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(APP)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;  // Google y el resto, directos a la red
  e.respondWith(caches.open(VERSION).then(async cache => {
    const cached = await cache.match(e.request, { ignoreSearch: true });
    const net = fetch(e.request).then(r => { if (r.ok) cache.put(e.request, r.clone()); return r; }).catch(() => null);
    return cached || (await net) || (e.request.mode === "navigate" ? cache.match("index.html") : Response.error());
  }));
});

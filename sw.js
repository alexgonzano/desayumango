// Ronda de cafés: guarda la "carcasa" de la app para que se pueda instalar y abra rápido.
// Los datos nunca se guardan aquí: siempre se piden al servidor (Apps Script) en directo.
const CACHE = 'ronda-cafes-v4';
const SHELL = ['./', './index.html', './config.js', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  // El servidor de datos y el inicio de sesión de Google van siempre en directo.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  // Primero la red, para ver siempre la última versión; sin conexión, lo guardado.
  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); }
        return res;
      })
      .catch(() => caches.match(req).then((r) => r || caches.match('./index.html')))
  );
});

// Avisos: el servidor manda el texto ya preparado ({title, body, tag}).
// Con la misma tag, un aviso sustituye al anterior (por ejemplo, «Ronda anulada»).
self.addEventListener('push', (event) => {
  const d = event.data ? event.data.json() : {};
  event.waitUntil(self.registration.showNotification(d.title || 'DesayuMango', {
    body: d.body || '', tag: d.tag, icon: 'icon-192.png', lang: 'es'
  }));
});

// Al tocar el aviso: a la app si ya está abierta; si no, se abre.
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true })
    .then((list) => (list.length ? list[0].focus() : self.clients.openWindow('./'))));
});

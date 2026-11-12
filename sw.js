/* Mis Vehículos — Service Worker v3 */
const VERSION = 'vtk-v3.3.0';
const SHELL = ['./', './index.html', './core.js', './manifest.webmanifest',
  './icon-192.png', './icon-512.png', './maskable-512.png', './badge-96.png'];
const STATE_CACHE = 'vtk-state';
const STATE_URL = './__state';

importScripts('./core.js');

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION)
    .then(c => Promise.allSettled(SHELL.map(u => fetch(u, { cache: 'reload' }).then(r => r.ok && c.put(u, r)))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== STATE_CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// App shell: cache primero, actualiza en segundo plano. La API (Google) nunca se cachea acá.
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  if (url.pathname.endsWith('/__state')) return;
  // Red primero (así cada cambio subido a GitHub se ve al toque); sin señal, usa la copia guardada
  e.respondWith((async () => {
    const cache = await caches.open(VERSION);
    try {
      const r = await Promise.race([
        fetch(e.request, { cache: 'no-cache' }),
        new Promise((_, ko) => setTimeout(() => ko(new Error('timeout')), 5000))
      ]);
      if (r && r.ok) { cache.put(e.request, r.clone()); return r; }
      throw new Error('bad');
    } catch (err) {
      const hit = await cache.match(e.request, { ignoreSearch: true });
      return hit || (e.request.mode === 'navigate' ? cache.match('./index.html') : Response.error());
    }
  })());
});

// ── Estado compartido con la página (config + datos + último aviso) ──
async function readState() {
  const c = await caches.open(STATE_CACHE);
  const r = await c.match(STATE_URL);
  return r ? r.json() : null;
}
async function writeState(st) {
  const c = await caches.open(STATE_CACHE);
  await c.put(STATE_URL, new Response(JSON.stringify(st), { headers: { 'Content-Type': 'application/json' } }));
}

function hoyStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

async function checkAndNotify(force) {
  const st = await readState();
  if (!st || !st.cfg) return;
  if (!force && st.lastNotified === hoyStr()) return;
  if (!force && st.cfg.notif === false) return;

  // intenta traer datos frescos
  if (st.cfg.apiUrl) {
    try {
      const u = new URL(st.cfg.apiUrl);
      u.searchParams.set('action', 'getAll');
      if (st.cfg.apiKey) u.searchParams.set('key', st.cfg.apiKey);
      const r = await fetch(u.toString(), { redirect: 'follow' });
      const d = await r.json();
      if (d && !d.error) st.data = { vehiculos: d.vehiculos || [], servicios: d.servicios || [] };
    } catch (e) { /* offline: usa los datos guardados */ }
  }
  const list = VTK.alertas(st.data || {}, st.cfg);
  st.lastNotified = hoyStr();
  await writeState(st);
  if (!list.length && !force) return;

  const venc = list.filter(x => x.e.level === 'venc').length;
  const title = list.length
    ? (venc ? `⚠️ ${venc} mantenimiento${venc > 1 ? 's' : ''} vencido${venc > 1 ? 's' : ''}` : `🔧 ${list.length} mantenimiento${list.length > 1 ? 's' : ''} por vencer`)
    : '✅ Todo al día';
  const body = list.length
    ? list.slice(0, 4).map(x => `${x.s.patente}: ${x.s.tipo.split('+')[0].trim()}${VTK.partes(x.s.tipo).length > 1 ? ' +' : ''} — ${VTK.textoEstado(x.e)}`).join('\n') +
      (list.length > 4 ? `\n…y ${list.length - 4} más` : '')
    : 'No hay mantenimientos pendientes.';
  await self.registration.showNotification(title, {
    body, tag: 'vtk-alertas', renotify: true,
    icon: './icon-192.png', badge: './badge-96.png',
    data: { url: './?tab=home' }
  });
}

self.addEventListener('periodicsync', e => {
  if (e.tag === 'vtk-check') e.waitUntil(checkAndNotify(false));
});
self.addEventListener('sync', e => {
  if (e.tag === 'vtk-check') e.waitUntil(checkAndNotify(false));
});
self.addEventListener('message', e => {
  const m = e.data || {};
  if (m.type === 'check') e.waitUntil(checkAndNotify(!!m.force));
  if (m.type === 'skipWaiting') self.skipWaiting();
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const target = new URL((e.notification.data && e.notification.data.url) || './', self.registration.scope).href;
  e.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
      for (const c of list) { if ('focus' in c) { c.postMessage({ type: 'goto', tab: 'home' }); return c.focus(); } }
      return self.clients.openWindow(target);
    })
  );
});

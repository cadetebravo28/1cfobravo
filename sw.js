/* CFOBRAVO · service worker
   Sempre busca a versão mais nova na internet (tempo real) e guarda uma cópia
   para abrir sem sinal com os últimos dados. */
// notificações (OneSignal); sem sinal, o app continua funcionando
try { importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js'); } catch (e) {}

const CACHE = 'cfobravo-v5';
const BASE = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  if (/\.(m4a|mp3|ogg|wav|aac)$/i.test(new URL(req.url).pathname)) return;   // áudio: direto (o iPhone precisa de Range)
  const url = new URL(req.url);
  const site = url.origin === self.location.origin;
  const api = /script\.google(usercontent)?\.com$/.test(url.hostname) || url.hostname === 'api.open-meteo.com';
  const fontes = /fonts\.(googleapis|gstatic)\.com$/.test(url.hostname);
  if (!site && !api && !fontes) return;
  // rede primeiro (dados e páginas sempre atuais); sem sinal, usa a última cópia
  e.respondWith(
    fetch(site ? req.url : req, site ? { cache: 'no-cache', credentials: 'same-origin' } : undefined).then(res => {
      if (res && (res.ok || res.type === 'opaque')) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(api ? keyApi(req) : req, cp)); }
      return res;
    }).catch(() => caches.match(api ? keyApi(req) : req, { ignoreSearch: !api }).then(r => r || (req.mode === 'navigate' ? caches.match('./index.html') : Response.error())))
  );
});
// a API redireciona para outro endereço; guarda pela URL pedida
function keyApi(req) { return new Request(req.url, { method: 'GET' }); }

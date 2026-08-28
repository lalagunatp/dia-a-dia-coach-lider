const CACHE='sv2-shell-v1';
const SHELL=['./','./index.html','./manifest.json'];

self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});

// Solo se cachea el shell de la app (HTML/manifest). Todo lo demás (Sheets, Apps Script,
// fotos) va directo a red para no servir datos desactualizados.
self.addEventListener('fetch',e=>{
  const url=new URL(e.request.url);
  if(url.origin!==location.origin)return;

  if(e.request.mode==='navigate'){
    e.respondWith(fetch(e.request).catch(()=>caches.match('./index.html')));
    return;
  }
  if(SHELL.some(s=>url.pathname.endsWith(s.replace('./','')))){
    e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));
  }
});

const CACHE='sv2-shell-v2';
const SHELL=['./','./index.html','./manifest.json'];
// Versión fijada de React/ReactDOM/Babel standalone (antes sin versión / "@18" sin fijar): permite
// cachearlas aquí y arrancar la app sin depender de que unpkg.com responda en ese momento — clave
// para que entre rápido y funcione igual con cualquier compañía celular, no solo con wifi.
const CDN=[
  'https://unpkg.com/react@18.3.1/umd/react.production.min.js',
  'https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js',
  'https://unpkg.com/@babel/standalone@7.24.7/babel.min.js',
];

self.addEventListener('install',e=>{
  e.waitUntil(Promise.all([
    caches.open(CACHE).then(c=>c.addAll(SHELL)),
    caches.open(CACHE).then(c=>Promise.all(CDN.map(u=>
      fetch(u,{mode:'no-cors'}).then(r=>c.put(u,r)).catch(()=>{})
    ))),
  ]));
  self.skipWaiting();
});

self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});

// Solo se cachea el shell de la app (HTML/manifest), React/ReactDOM/Babel (versión fija, nunca
// cambian) y el catálogo de Clusters/PDV (ver index.html: fetchClustersColonias/fetchPDV, que
// tienen su propio respaldo en localStorage). Todo lo demás (Sheets, Apps Script, fotos) va
// directo a red para no servir datos desactualizados.
self.addEventListener('fetch',e=>{
  // React/ReactDOM/Babel: caché primero — arranca la app al instante sin esperar a unpkg.com.
  if(CDN.includes(e.request.url)){
    e.respondWith(
      caches.match(e.request).then(cached=>cached||fetch(e.request,{mode:'no-cors'}).then(r=>{
        caches.open(CACHE).then(c=>c.put(e.request,r.clone()));
        return r;
      }))
    );
    return;
  }

  const url=new URL(e.request.url);
  if(url.origin!==location.origin)return;

  if(e.request.mode==='navigate'){
    // Red primero, pero con tope de 3s: si la señal está lenta (no caída, solo lenta) no vale la
    // pena que la app se sienta trabada esperando — se usa la última copia guardada mientras la
    // red sigue intentando en segundo plano y actualiza la caché para la próxima vez.
    e.respondWith((async()=>{
      const netPromise=fetch(e.request).then(r=>{
        if(r&&r.ok)caches.open(CACHE).then(c=>c.put('./index.html',r.clone()));
        return r;
      }).catch(()=>null);
      const timeout=new Promise(res=>setTimeout(()=>res(null),3000));
      const fast=await Promise.race([netPromise,timeout]);
      if(fast)return fast;
      const cached=await caches.match('./index.html');
      return cached||netPromise.then(r=>r||new Response('Sin conexión. Intenta de nuevo.',{status:503,headers:{'Content-Type':'text/plain;charset=utf-8'}}));
    })());
    return;
  }
  if(SHELL.some(s=>url.pathname.endsWith(s.replace('./','')))){
    e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));
  }
});

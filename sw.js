const CACHE='sv2-shell-v4';
// React/ReactDOM/Babel ahora viven en ./vendor (mismo origen, servidos por GitHub Pages) en vez de
// bajarse de unpkg.com — un solo origen decide si la app arranca, no dos. Se cachean aquí junto con
// el resto del shell, con el mismo camino simple que ./index.html (ya no hace falta el manejo
// especial "no-cors" que necesitaban cuando venían de un origen distinto).
const SHELL=['./','./index.html','./manifest.json','./vendor/react.production.min.js','./vendor/react-dom.production.min.js','./vendor/babel.min.js'];

// Con datos móviles inestables un fetch puede quedarse "colgado" (ni resuelve ni truena) en vez de
// fallar rápido — sin este tope, eso bloqueaba la instalación completa del service worker (si UN
// solo archivo no respondía, quedaba detenido) o dejaba una carga de script esperando para siempre.
// Con el tope, lo peor que pasa es que ese archivo en particular no queda pre-cacheado esta vez, no
// que toda la app se quede en blanco.
function fetchConTope(req,opts,ms){
  return Promise.race([
    fetch(req,opts),
    new Promise((_,rej)=>setTimeout(()=>rej(new Error('timeout')),ms)),
  ]);
}

self.addEventListener('install',e=>{
  e.waitUntil(
    caches.open(CACHE).then(c=>Promise.all(SHELL.map(u=>
      fetchConTope(u,undefined,10000).then(r=>c.put(u,r)).catch(()=>{})
    )))
  );
  self.skipWaiting();
});

self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});

// Solo se cachea el shell de la app (HTML/manifest/React/ReactDOM/Babel, todo mismo origen y
// versión fija). Todo lo demás (Sheets, Apps Script, fotos) va directo a red para no servir datos
// desactualizados — esos tienen su propio respaldo en localStorage (ver index.html).
self.addEventListener('fetch',e=>{
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
    e.respondWith(caches.match(e.request).then(r=>r||fetchConTope(e.request,undefined,15000)));
  }
});

const CACHE='gps-fm-shell-v9';
const FILES=['./','./index.html','./manifest.json','./config.js','./icon.png'];

self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(FILES)).then(()=>self.skipWaiting()));
});

self.addEventListener('activate',event=>{
  event.waitUntil(caches.keys()
    .then(keys=>Promise.all(keys.filter(key=>key.startsWith('gps-fm-shell-')&&key!==CACHE).map(key=>caches.delete(key))))
    .then(()=>self.clients.claim()));
});

self.addEventListener('fetch',event=>{
  const request=event.request;
  const url=new URL(request.url);
  if(request.method!=='GET'||url.origin!==self.location.origin)return;

  const isNavigation=request.mode==='navigate'||url.pathname.endsWith('/index.html');
  const isConfig=url.pathname.endsWith('/config.js');

  if(isNavigation||isConfig){
    event.respondWith(fetch(request,{cache:'no-store'}).then(response=>{
      if(response&&response.ok){
        const copy=response.clone();
        return caches.open(CACHE).then(cache=>cache.put(request,copy)).then(()=>response);
      }
      return response;
    }).catch(()=>caches.match(request).then(cached=>cached||caches.match('./index.html'))));
    return;
  }

  event.respondWith(caches.match(request).then(cached=>{
    if(cached)return cached;
    return fetch(request).then(response=>{
      if(response&&response.ok){
        const copy=response.clone();
        return caches.open(CACHE).then(cache=>cache.put(request,copy)).then(()=>response);
      }
      return response;
    });
  }));
});

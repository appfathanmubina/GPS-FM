const CACHE='gps-fm-shell-v3';
const FILES=['./','./index.html','./manifest.json','./config.js'];
self.addEventListener('install',e=>e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES)).then(()=>self.skipWaiting())));
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',e=>{const u=new URL(e.request.url);if(e.request.method!=='GET'||u.origin!==self.location.origin)return;if(u.pathname.endsWith('/config.js')){e.respondWith(fetch(e.request,{cache:'no-store'}).catch(()=>caches.match(e.request)));return;}e.respondWith(caches.match(e.request).then(c=>c||fetch(e.request)));});

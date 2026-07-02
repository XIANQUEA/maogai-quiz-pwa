const CACHE="maogai-quiz-v1";
const ASSETS=["./","./index.html","./manifest.webmanifest","./assets/icon.svg","./styles/tokens.css","./styles/base.css","./styles/components.css","./src/app.js","./src/config.js","./src/domain/questions.js","./src/domain/session.js","./src/domain/progress.js","./src/data/question-repository.js","./src/data/progress-store.js","./src/data/backup.js","./src/ui/screens.js","./src/ui/controller.js","./data/questions.v1.json"];
self.addEventListener("install",event=>event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(ASSETS))));
self.addEventListener("activate",event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(key=>key!==CACHE).map(key=>caches.delete(key)))).then(()=>self.clients.claim())));
self.addEventListener("fetch",event=>{if(event.request.method!=="GET")return;event.respondWith(caches.match(event.request).then(cached=>cached||fetch(event.request)));});

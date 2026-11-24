/* 투자일지 앱 화면 파일을 휴대폰에 저장해 두고 바로 띄움. 엑셀 데이터는 저장하지 않음 */
const CACHE = 'tujailji-v16';
const SHELL = ['./', './index.html', './xlsx.mini.min.js', './worker.js', './excelimport.js', './manifest.webmanifest', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  const isPage = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('index.html');
  const key = isPage ? './index.html' : req;
  e.respondWith(caches.open(CACHE).then(c => c.match(key).then(hit => {
    const net = fetch(req).then(res => { if (res.ok) c.put(key, res.clone()); return res; }).catch(() => hit);
    return isPage ? (hit || net) : (hit || net);
  })));
});

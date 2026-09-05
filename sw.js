/* Service worker app KHO — mở được app khi không có mạng (chỉ giao diện; dữ liệu vẫn cần mạng).
   Đổi CACHE mỗi lần phát hành bản mới để điện thoại tự nạp bản mới. */
const CACHE = 'l13fc-kho-v1.7.0';
const FILES = ['./', './index.html', './config.js', './xlsx-lite.js', './icons/header-cty.jpg', './huong-dan-cai-dat.html', './manifest.webmanifest',
  './icons/logo.png', './icons/logo-login.png', './icons/icon-192.png', './icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks =>
    Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.hostname.includes('script.google.com')) return;   // API luôn đi mạng
  if (url.hostname.includes('drive.google.com') || url.hostname.includes('googleusercontent.com')) return; // ảnh phiếu không cache
  e.respondWith(
    fetch(e.request).then(r => {
      const copy = r.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return r;
    }).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
  );
});

/**
 * File: sw.js
 * Dự Án: HKD Đào Thanh Yến (PWA Offline-first)
 */
// Cập nhật Timestamp hiện tại: 19/07/2026 - 11:16
const CACHE_NAME = 'hkd-dty-20260719-1116';

// Các tài nguyên tĩnh cục bộ và link hình ảnh cần bộ đệm hoạt động ngoại tuyến
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './manifest.json',
  './logo-192.png',
  './logo-512.jpg',
  'https://raw.githubusercontent.com/Nguyennth19/hkd-dao-thanh-yen/refs/heads/main/load-phone-hkd-dty.png',
  'https://raw.githubusercontent.com/Nguyennth19/hkd-dao-thanh-yen/refs/heads/main/load-computer-hkd-dty.png',
  'https://raw.githubusercontent.com/Nguyennth19/hkd-dao-thanh-yen/refs/heads/main/icon-htk-dty.png'
];

// Sự kiện cài đặt (Install) - Nạp tài nguyên vào Cache
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] Đang nạp tài nguyên tĩnh vào bộ đệm...');
      return cache.addAll(ASSETS_TO_CACHE);
    }).then(() => self.skipWaiting())
  );
});

// Sự kiện kích hoạt (Activate) - Giải phóng và xóa bỏ cache cũ
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            console.log('[Service Worker] Đang dọn dẹp bộ đệm cũ:', cache);
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Sự kiện fetch - Đọc bộ đệm (Offline-first), bỏ qua API Google Sheets
self.addEventListener('fetch', (event) => {
  // Bỏ qua các lệnh API gửi lên Google Apps Script
  if (event.request.url.includes('script.google.com')) return;

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      // Trả về cache nếu có sẵn, nếu không sẽ tải từ mạng
      return cachedResponse || fetch(event.request);
    })
  );
});

// Lắng nghe sự kiện 'message' để thực thi skipWaiting (Cập nhật phiên bản mới)
self.addEventListener('message', (event) => {
  if (event.data && event.data.action === 'skipWaiting') {
    console.log('[Service Worker] Nhận được tín hiệu skipWaiting. Đang ép cập nhật phiên bản mới...');
    self.skipWaiting();
  }
});

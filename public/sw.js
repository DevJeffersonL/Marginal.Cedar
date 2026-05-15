// Simple Service Worker for PWA compliance
const CACHE_NAME = 'marginal-v1';

self.addEventListener('install', (event) => {
  console.log('[SW] Installed');
});

self.addEventListener('fetch', (event) => {
  // Pass-through for now to ensure app works while being installable
  event.respondWith(fetch(event.request));
});

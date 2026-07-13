// sw.js - Basic Service Worker to enable PWA installation on Android Chrome

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(clients.claim());
});

// A fetch listener is required by Android Chrome to trigger the "Add to Home Screen" prompt.
// We just pass the request straight through to the network.
self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});

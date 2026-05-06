self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

// Пустой fetch-хендлер: сеть работает как обычно, но приложение готово к PWA-режиму.
self.addEventListener("fetch", () => {});

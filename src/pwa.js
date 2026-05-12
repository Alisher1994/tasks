export function isStandalonePwaMode() {
  try {
    const byDisplayMode = typeof window.matchMedia === "function"
      && window.matchMedia("(display-mode: standalone)").matches;
    const byNavigator = window.navigator?.standalone === true;
    return byDisplayMode || byNavigator;
  } catch (_) {
    return false;
  }
}

export function applyStandalonePwaClass() {
  const enabled = isStandalonePwaMode();
  document.documentElement.classList.toggle("pwa-standalone", enabled);
  document.body.classList.toggle("pwa-standalone", enabled);
}

export async function registerPwaServiceWorker() {
  try {
    if (!("serviceWorker" in navigator)) return;
    const isLocal = location.protocol === "http:" && /^(localhost|127\.0\.0\.1)$/i.test(location.hostname);
    const isSecure = location.protocol === "https:" || isLocal;
    if (!isSecure) return;
    await navigator.serviceWorker.register("/sw.js");
  } catch (_) {
    // noop
  }
}

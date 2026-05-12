import { getAuthToken } from "./token.js";

const SESSION_LAST_ACTIVITY_KEY = "mbc_task_last_activity_ts";
const SESSION_IDLE_TIMEOUT_MS = 12 * 60 * 60 * 1000;
const SESSION_IDLE_CHECK_INTERVAL_MS = 60 * 1000;

let lastSessionActivityWriteAt = 0;
let sessionActivityListenersBound = false;
let sessionIdleCheckTimer = null;
let onIdleExpiredCallback = null;

export function configureSessionIdle({ onExpired } = {}) {
  onIdleExpiredCallback = typeof onExpired === "function" ? onExpired : null;
}

function readLastSessionActivityTs() {
  try {
    const raw = sessionStorage.getItem(SESSION_LAST_ACTIVITY_KEY);
    const n = Number(raw);
    return Number.isFinite(n) && n > 0 ? n : 0;
  } catch (_) {
    return 0;
  }
}

function writeLastSessionActivityTs(ts = Date.now()) {
  try {
    sessionStorage.setItem(SESSION_LAST_ACTIVITY_KEY, String(Math.max(0, Math.floor(ts))));
  } catch (_) {
    /* noop */
  }
}

export function clearLastSessionActivityTs() {
  try {
    sessionStorage.removeItem(SESSION_LAST_ACTIVITY_KEY);
  } catch (_) {
    /* noop */
  }
}

export function trackSessionActivity(force = false) {
  if (!getAuthToken()) return;
  const now = Date.now();
  if (!force && now - lastSessionActivityWriteAt < 5000) return;
  lastSessionActivityWriteAt = now;
  writeLastSessionActivityTs(now);
}

function bindSessionActivityListeners() {
  if (sessionActivityListenersBound) return;
  sessionActivityListenersBound = true;
  const handler = () => trackSessionActivity(false);
  const opts = { passive: true, capture: true };
  ["pointerdown", "keydown", "touchstart", "wheel", "scroll"].forEach((eventName) => {
    window.addEventListener(eventName, handler, opts);
  });
}

function checkSessionIdleTimeout() {
  if (!getAuthToken()) return;
  const last = readLastSessionActivityTs();
  if (!last) {
    trackSessionActivity(true);
    return;
  }
  if (Date.now() - last >= SESSION_IDLE_TIMEOUT_MS && onIdleExpiredCallback) {
    onIdleExpiredCallback();
  }
}

export function startSessionIdleWatcher() {
  if (!getAuthToken()) return;
  bindSessionActivityListeners();
  trackSessionActivity(true);
  clearInterval(sessionIdleCheckTimer);
  sessionIdleCheckTimer = setInterval(() => {
    checkSessionIdleTimeout();
  }, SESSION_IDLE_CHECK_INTERVAL_MS);
}

export function stopSessionIdleWatcher() {
  clearInterval(sessionIdleCheckTimer);
  sessionIdleCheckTimer = null;
}

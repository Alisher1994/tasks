/**
 * Клиентский WebSocket-канал к /ws/state.
 *
 * Архитектура:
 *  - Один сокет на всё приложение, живёт пока пользователь авторизован.
 *  - Авторизация: JWT в query (?token=...). Браузерный WebSocket не умеет слать
 *    кастомные заголовки, поэтому это единственный безопасный способ.
 *  - Сервер шлёт два типа сообщений: { type: "hello", rev } сразу после connect,
 *    и { type: "state_changed", rev } после каждого успешного апдейта payload.
 *  - При получении state_changed мы НЕ принимаем payload по WS — это сделано
 *    специально: WS-канал даёт только сигнал "что-то поменялось", а данные забираем
 *    через обычный GET /api/data, где уже работает 3-way merge на 409.
 *  - Reconnect: экспоненциальный backoff с jitter (1с, 2с, 4с, 8с, 16с, 30с max).
 *  - Heartbeat: каждые 25 сек шлём ping в виде JSON-сообщения; сервер также пингует
 *    нас своим WS-ping/pong на низком уровне.
 *
 * Контракт с main.js — модуль публикует configureRealtime/startRealtimeSync/stopRealtimeSync.
 */

let socket = null;
let reconnectTimer = null;
let reconnectAttempt = 0;
let pingTimer = null;
let shouldRun = false;

let getTokenFn = null;
let onStateChangedFn = null;
let getCurrentRevFn = null;
let onConnectFn = null;
let onDisconnectFn = null;

function buildWebSocketUrl(token) {
  const protocol = typeof location !== "undefined" && location.protocol === "https:" ? "wss:" : "ws:";
  const host = typeof location !== "undefined" ? location.host : "localhost";
  return `${protocol}//${host}/ws/state?token=${encodeURIComponent(token)}`;
}

export function configureRealtime(deps) {
  if (deps && typeof deps === "object") {
    if (typeof deps.getToken === "function") getTokenFn = deps.getToken;
    if (typeof deps.onStateChanged === "function") onStateChangedFn = deps.onStateChanged;
    if (typeof deps.getCurrentRev === "function") getCurrentRevFn = deps.getCurrentRev;
    if (typeof deps.onConnect === "function") onConnectFn = deps.onConnect;
    if (typeof deps.onDisconnect === "function") onDisconnectFn = deps.onDisconnect;
  }
}

export function startRealtimeSync() {
  shouldRun = true;
  if (socket && (socket.readyState === 0 || socket.readyState === 1)) {
    return; // уже работает
  }
  connect();
}

export function stopRealtimeSync() {
  shouldRun = false;
  if (reconnectTimer) {
    clearTimeout(reconnectTimer);
    reconnectTimer = null;
  }
  if (pingTimer) {
    clearInterval(pingTimer);
    pingTimer = null;
  }
  if (socket) {
    try { socket.close(1000, "client_logout"); } catch {}
    socket = null;
  }
  reconnectAttempt = 0;
}

export function realtimeStatus() {
  if (!socket) return "disconnected";
  if (typeof WebSocket === "undefined") return "unsupported";
  switch (socket.readyState) {
    case WebSocket.CONNECTING: return "connecting";
    case WebSocket.OPEN: return "open";
    case WebSocket.CLOSING: return "closing";
    case WebSocket.CLOSED: return "closed";
    default: return "unknown";
  }
}

function connect() {
  if (!shouldRun) return;
  if (typeof WebSocket === "undefined") return;
  if (typeof getTokenFn !== "function") return;
  const token = getTokenFn();
  if (!token) {
    // Без токена соединяться нельзя — попробуем позже, когда логин завершится.
    scheduleReconnect();
    return;
  }
  let ws;
  try {
    ws = new WebSocket(buildWebSocketUrl(token));
  } catch {
    scheduleReconnect();
    return;
  }
  socket = ws;

  ws.addEventListener("open", () => {
    reconnectAttempt = 0;
    if (pingTimer) clearInterval(pingTimer);
    pingTimer = setInterval(() => {
      if (!ws || ws.readyState !== WebSocket.OPEN) return;
      try { ws.send(JSON.stringify({ type: "ping", t: Date.now() })); } catch {}
    }, 25_000);
    if (typeof onConnectFn === "function") {
      try { onConnectFn(); } catch {}
    }
  });

  ws.addEventListener("message", (event) => {
    let data = null;
    try { data = JSON.parse(typeof event.data === "string" ? event.data : ""); } catch { return; }
    if (!data || typeof data !== "object") return;
    const type = String(data.type || "");
    if (type === "pong") return;
    if (type === "hello" || type === "state_changed") {
      const rev = Number(data.rev) || 0;
      const currentRev = typeof getCurrentRevFn === "function" ? Number(getCurrentRevFn()) || 0 : 0;
      // Подтягиваем только если сервер ушёл вперёд относительно нашего последнего pull/push.
      // Это игнорирует наш собственный broadcast после успешного PUT (rev уже равен).
      if (rev > currentRev && typeof onStateChangedFn === "function") {
        try { onStateChangedFn({ rev, source: data.source, by: data.by, type }); } catch {}
      }
    }
  });

  ws.addEventListener("close", (event) => {
    if (pingTimer) { clearInterval(pingTimer); pingTimer = null; }
    socket = null;
    if (typeof onDisconnectFn === "function") {
      try { onDisconnectFn(event); } catch {}
    }
    if (event.code === 1000) return; // штатное закрытие
    if (event.code === 1008 || event.code === 4401 || event.code === 4403) {
      // Авторизация отвалилась — не дёргаемся, пусть auth-flow разберётся.
      shouldRun = false;
      return;
    }
    scheduleReconnect();
  });

  ws.addEventListener("error", () => {
    // close-событие всё равно прилетит, ничего не делаем
  });
}

function scheduleReconnect() {
  if (!shouldRun) return;
  if (reconnectTimer) clearTimeout(reconnectTimer);
  reconnectAttempt += 1;
  const step = Math.min(reconnectAttempt - 1, 5);
  const baseDelay = Math.min(30_000, 1000 * 2 ** step);
  const jitter = Math.floor(Math.random() * 1000);
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connect();
  }, baseDelay + jitter);
}

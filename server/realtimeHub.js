/**
 * WebSocket-хаб для real-time синхронизации app_state между клиентами.
 *
 * Архитектура:
 *  - Один WebSocketServer, прицеплен к существующему http-серверу через upgrade-событие.
 *  - При подключении клиент передаёт JWT в query (?token=...) — других безопасных способов
 *    отдать заголовок Authorization для браузерного WebSocket нет.
 *  - После аутентификации хаб шлёт клиенту { type: "hello", rev } — текущая ревизия БД.
 *  - Когда любой клиент успешно делает PUT /api/data, сервер вызывает broadcastStateChanged(rev, exceptUserId)
 *    и хаб рассылает { type: "state_changed", rev } всем остальным.
 *  - Клиент при получении state_changed сравнивает rev с локальным и при необходимости
 *    вызывает GET /api/data — никакого payload в WS-сообщении нет (мелкие сообщения, безопасно).
 *  - Heartbeat: каждые 30 сек сервер шлёт ping, клиент отвечает pong. Если pong не пришёл —
 *    соединение признаётся мёртвым и закрывается (клиент переподключится).
 */

import { WebSocketServer } from "ws";
import jwt from "jsonwebtoken";

const HEARTBEAT_INTERVAL_MS = 30_000;

let wssInstance = null;
let heartbeatTimer = null;

function safeSend(ws, payload) {
  if (ws.readyState !== 1) return;
  try {
    ws.send(JSON.stringify(payload));
  } catch (e) {
    // Соединение могло закрыться между проверкой и send — игнорируем.
  }
}

function parseTokenFromRequest(req) {
  try {
    const url = new URL(req.url, "http://localhost");
    const t = url.searchParams.get("token");
    if (t) return t;
  } catch {}
  // Fallback: Sec-WebSocket-Protocol — некоторые клиенты используют этот канал для токена.
  const proto = req.headers["sec-websocket-protocol"];
  if (typeof proto === "string") {
    const parts = proto.split(",").map((s) => s.trim());
    const bearerPart = parts.find((p) => p.startsWith("bearer.") || p.startsWith("token."));
    if (bearerPart) return bearerPart.split(".").slice(1).join(".");
  }
  return null;
}

export function attachRealtimeHub({ httpServer, jwtSecret, isSessionVersionCurrent, getCurrentRevision }) {
  if (wssInstance) {
    return wssInstance;
  }
  const wss = new WebSocketServer({ noServer: true });

  httpServer.on("upgrade", async (req, socket, head) => {
    // Принимаем только наш путь, остальные апгрейды отдаём другим возможным WS-серверам/middleware.
    let pathname = "/";
    try {
      pathname = new URL(req.url, "http://localhost").pathname;
    } catch {}
    if (pathname !== "/ws/state") {
      return;
    }
    const token = parseTokenFromRequest(req);
    if (!token) {
      socket.write("HTTP/1.1 401 Unauthorized\r\n\r\n");
      socket.destroy();
      return;
    }
    let payload;
    try {
      payload = jwt.verify(token, jwtSecret);
    } catch {
      socket.write("HTTP/1.1 401 Unauthorized\r\n\r\n");
      socket.destroy();
      return;
    }
    const subject = payload?.sub != null ? String(payload.sub) : "";
    try {
      const ok = await isSessionVersionCurrent(subject, payload?.sv);
      if (!ok) {
        socket.write("HTTP/1.1 401 Unauthorized\r\n\r\n");
        socket.destroy();
        return;
      }
    } catch {
      socket.write("HTTP/1.1 500 Internal Server Error\r\n\r\n");
      socket.destroy();
      return;
    }
    wss.handleUpgrade(req, socket, head, (ws) => {
      ws.userId = subject;
      ws.isAlive = true;
      wss.emit("connection", ws, req);
    });
  });

  wss.on("connection", async (ws) => {
    ws.on("pong", () => {
      ws.isAlive = true;
    });
    ws.on("message", (raw) => {
      // Сейчас клиент может присылать только pong-like сообщения; payload игнорируем.
      // На будущее: presence, typing-indicators и т.п.
      let msg = null;
      try {
        msg = JSON.parse(raw.toString());
      } catch {
        return;
      }
      if (!msg || typeof msg !== "object") return;
      if (msg.type === "ping") {
        safeSend(ws, { type: "pong", t: Date.now() });
      }
    });
    ws.on("error", () => {
      // Мёртвое соединение — пусть закрывается естественным путём.
    });

    try {
      const rev = await getCurrentRevision();
      safeSend(ws, { type: "hello", rev });
    } catch {
      safeSend(ws, { type: "hello", rev: 0 });
    }
  });

  heartbeatTimer = setInterval(() => {
    wss.clients.forEach((ws) => {
      if (ws.isAlive === false) {
        try { ws.terminate(); } catch {}
        return;
      }
      ws.isAlive = false;
      try { ws.ping(); } catch {}
    });
  }, HEARTBEAT_INTERVAL_MS);

  wss.on("close", () => {
    if (heartbeatTimer) {
      clearInterval(heartbeatTimer);
      heartbeatTimer = null;
    }
  });

  wssInstance = wss;
  return wss;
}

/**
 * Рассылка события "state_changed" всем подключённым клиентам.
 * exceptUserId — если задан, инициатор изменения пропускается (его клиент уже знает свой ревизию по ответу PUT).
 * Однако оставлять одно и то же устройство без апдейта опасно — пользователь может иметь
 * несколько вкладок/устройств. Поэтому шлём ВСЕМ, а клиент сравнит revision с локальным.
 */
export function broadcastStateChanged(rev, meta = {}) {
  if (!wssInstance) return 0;
  let count = 0;
  const msg = { type: "state_changed", rev: Number(rev) || 0, ...(meta && typeof meta === "object" ? meta : {}) };
  wssInstance.clients.forEach((ws) => {
    if (ws.readyState !== 1) return;
    safeSend(ws, msg);
    count += 1;
  });
  return count;
}

export function getRealtimeConnectionCount() {
  if (!wssInstance) return 0;
  return wssInstance.clients.size;
}

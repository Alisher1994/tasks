/**
 * API + раздача статики для Railway.
 * Переменные: DATABASE_URL, JWT_SECRET, ADMIN_PHONE, ADMIN_PASSWORD (первый админ), PORT
 */
import express from "express";
import { promises as fsp } from "fs";
import path from "path";
import { fileURLToPath } from "url";
import jwt from "jsonwebtoken";
import pg from "pg";
import bcrypt from "bcryptjs";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import { randomBytes } from "crypto";
import { configureTelegramWebhook, handleTelegramWebhook } from "./telegramWebhook.js";
import { runGoogleSheetsSync, startGoogleSheetsAutoSync } from "./googleSheetsSync.js";
import { runMigrations } from "./migrate.js";
import { validateAppPayload } from "./validatePayload.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.join(__dirname, "..");

const { Pool } = pg;
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes("localhost") ? false : { rejectUnauthorized: false }
});

const JWT_SECRET = process.env.JWT_SECRET || "change-me-in-production";
const ADMIN_PHONE = process.env.ADMIN_PHONE || "";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "";
const NODE_ENV = process.env.NODE_ENV || "development";
const MEDIA_STORAGE_PATH = String(process.env.MEDIA_STORAGE_PATH || "").trim()
  || path.join(rootDir, "storage", "media");
const TASK_NUMBER_COL = 0;
const TASK_STATUS_COL = 2;
const TASK_PRIORITY_COL = 3;
const TASK_ADDED_DATE_COL = 4;
const TASK_PHASE_COL = 5;
const TASK_SECTION_COL = 6;
const TASK_SUBSECTION_COL = 7;
const TASK_TITLE_COL = 8;
const TASK_RESPONSIBLE_COL = 9;
const TASK_ASSIGNED_COL = 10;
const TASK_NOTE_COL = 11;
const TASK_PLAN_COL = 12;
const TASK_CLOSED_DATE_COL = 15;
const TASK_MEDIA_AFTER_COL = 17;
const TASK_READ_STATE_COL = 18;
const TASK_LAST_SENT_AT_COL = 19;
const TASK_DELAY_REASON_COL = 20;
const TASK_CREATED_BY_COL = 21;
const TASK_CREATED_AT_COL = 22;
const TASK_REASSIGN_REASON_COL = 23;
const EMPLOYEE_FULL_NAME_COL = 1;
const EMPLOYEE_TELEGRAM_COL = 5;
const EMPLOYEE_CHAT_ID_COL = 6;
const CHAT_CLEAR_CODE_TTL_MS = 10 * 60 * 1000;
const CHAT_CLEAR_DELETE_SCAN_BEFORE = 3000;
const CHAT_CLEAR_DELETE_SCAN_AFTER = 40;

function normalizePhone(raw) {
  const src = String(raw || "").trim();
  if (!src) return "";
  let s = src.replace(/[^\d+]/g, "");
  if (s.startsWith("00")) s = `+${s.slice(2)}`;
  if (!s.startsWith("+")) s = `+${s.replace(/\D/g, "")}`;
  const digits = s.slice(1).replace(/\D/g, "").slice(0, 15);
  return digits ? `+${digits}` : "";
}

function last4DigitsPasswordFromPhone(rawPhone) {
  const digits = String(rawPhone || "").replace(/\D/g, "");
  if (digits.length < 4) return "";
  return digits.slice(-4);
}

function getEmployeesSection(payload) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  return sections.find((s) => s && s.id === "employees");
}

function findEmployeeByIdInPayload(payload, employeeId) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  const want = String(employeeId || "").trim();
  if (!want) return null;
  return rows.find((row) => String(row?.[0] || "").trim() === want) || null;
}

function findEmployeeByPhoneInPayload(payload, phone) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  for (const row of rows) {
    const rowPhone = normalizePhone(row?.[4] || "");
    if (rowPhone === phone) return row;
  }
  return null;
}

function findEmployeeByDisplayNameInPayload(payload, displayName) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  const want = String(displayName || "").trim().replace(/\s+/g, " ").toLowerCase();
  if (!want) return null;
  for (const row of rows) {
    const name = String(row?.[1] || "").trim().replace(/\s+/g, " ").toLowerCase();
    if (name === want) return row;
  }
  return null;
}

function normalizeLabelKey(raw) {
  return String(raw || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[^a-zа-я0-9]+/giu, "");
}

function resolveEmployeeAdminAccessColumnIndex(payload) {
  const employees = getEmployeesSection(payload);
  const cols = Array.isArray(employees?.columns) ? employees.columns : [];
  const aliases = new Set([
    "админдоступ",
    "админ",
    "доступадмина",
    "рольдоступа",
    "accessadmin",
    "adminaccess"
  ]);
  for (let i = 0; i < cols.length; i += 1) {
    if (aliases.has(normalizeLabelKey(cols[i]))) return i;
  }
  return 8;
}

function isEmployeeAdminAccessEnabled(payload, employeeRow) {
  if (!Array.isArray(employeeRow)) return false;
  const idx = resolveEmployeeAdminAccessColumnIndex(payload);
  const raw = String(employeeRow[idx] || "").trim().toLowerCase();
  return ["да", "true", "1", "yes", "on", "admin", "админ"].includes(raw);
}

function ensureObjectStore(payload, key) {
  if (!payload[key] || typeof payload[key] !== "object" || Array.isArray(payload[key])) {
    payload[key] = {};
  }
  return payload[key];
}

async function loadAppPayload() {
  const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
  return rows[0]?.payload && typeof rows[0].payload === "object"
    ? JSON.parse(JSON.stringify(rows[0].payload))
    : {};
}

async function saveAppPayload(payload) {
  const nextPayload = payload && typeof payload === "object"
    ? JSON.parse(JSON.stringify(payload))
    : {};
  try {
    const currentPayload = await loadAppPayload();
    const currentSections = Array.isArray(currentPayload?.sections) ? currentPayload.sections : [];
    const nextSections = Array.isArray(nextPayload?.sections) ? nextPayload.sections : [];
    const currentTasks = currentSections.find((s) => s && s.id === "tasks");
    const nextTasksIndex = nextSections.findIndex((s) => s && s.id === "tasks");
    if (currentTasks && nextTasksIndex >= 0) {
      nextSections[nextTasksIndex] = JSON.parse(JSON.stringify(currentTasks));
    }
  } catch (_) {
    /* noop */
  }
  await pool.query(
    `INSERT INTO app_state (id, payload, updated_at)
     VALUES (1, $1::jsonb, NOW())
     ON CONFLICT (id) DO UPDATE SET payload = EXCLUDED.payload, updated_at = NOW()`,
    [JSON.stringify(nextPayload)]
  );
}

function buildBotApiUrl(token, method) {
  return `https://api.telegram.org/bot${token}/${method}`;
}

async function tgSendMessage(token, chatId, text) {
  const r = await fetch(buildBotApiUrl(token, "sendMessage"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: String(chatId || "").trim(), text: String(text || "") })
  });
  const j = await r.json().catch(() => ({}));
  return {
    ok: r.ok && j?.ok === true,
    description: String(j?.description || ""),
    messageId: Number(j?.result?.message_id) || 0
  };
}

async function tgDeleteMessage(token, chatId, messageId) {
  const r = await fetch(buildBotApiUrl(token, "deleteMessage"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: String(chatId || "").trim(), message_id: Number(messageId) || 0 })
  });
  const j = await r.json().catch(() => ({}));
  return { ok: r.ok && j?.ok === true, description: String(j?.description || "") };
}

async function resolveInitiatorEmployeeRow(payload, reqUser) {
  const sub = String(reqUser?.sub || "").trim();
  if (!sub) return null;
  if (sub.startsWith("emp:")) {
    const phone = normalizePhone(sub.slice(4));
    return phone ? findEmployeeByPhoneInPayload(payload, phone) : null;
  }
  if (/^\d+$/.test(sub)) {
    const { rows } = await pool.query("SELECT phone, display_name FROM users WHERE id = $1", [Number(sub)]);
    if (rows.length) {
      const phone = normalizePhone(rows[0]?.phone || "");
      if (phone) {
        const byPhone = findEmployeeByPhoneInPayload(payload, phone);
        if (byPhone) return byPhone;
      }
      const byName = findEmployeeByDisplayNameInPayload(payload, rows[0]?.display_name || "");
      if (byName) return byName;
    }
  }
  return findEmployeeByDisplayNameInPayload(payload, reqUser?.name || "");
}

function getChatClearMessageAnchors(payload, targetChatId) {
  const chat = String(targetChatId || "").trim();
  const out = new Set();
  const lastCtx = payload?.telegramLastTaskByChat?.[chat];
  const prompt = Number(lastCtx?.promptMessageId) || 0;
  if (prompt > 0) out.add(prompt);
  const lastSeen = Number(payload?.telegramLastSeenMessageByChat?.[chat]) || 0;
  if (lastSeen > 0) out.add(lastSeen);
  return Array.from(out).sort((a, b) => b - a);
}

async function clearTelegramChatMessagesBestEffort(payload, token, targetChatId) {
  const chat = String(targetChatId || "").trim();
  const anchors = getChatClearMessageAnchors(payload, chat);
  // Пробный message_id из текущего состояния чата: помогает чистить даже если
  // задачи уже удалены из таблицы и нет внутренних привязок.
  const probe = await tgSendMessage(token, chat, "🧹 Выполняется очистка чата...");
  if (probe.ok && probe.messageId > 0) {
    anchors.push(probe.messageId);
  }
  const tested = new Set();
  const ids = [];
  Array.from(new Set(anchors.filter((x) => Number.isFinite(x) && x > 0))).forEach((anchor) => {
    for (let mid = anchor + CHAT_CLEAR_DELETE_SCAN_AFTER; mid >= Math.max(1, anchor - CHAT_CLEAR_DELETE_SCAN_BEFORE); mid -= 1) {
      if (tested.has(mid)) continue;
      tested.add(mid);
      ids.push(mid);
    }
  });
  if (!ids.length) {
    for (let mid = 3500; mid >= 1; mid -= 1) ids.push(mid);
  }
  let deleted = 0;
  let failed = 0;
  let tooOld = 0;
  let notFound = 0;
  for (const mid of ids) {
    // eslint-disable-next-line no-await-in-loop
    const r = await tgDeleteMessage(token, chat, mid);
    if (r.ok) {
      deleted += 1;
    } else {
      failed += 1;
      const d = String(r.description || "").toLowerCase();
      if (d.includes("can't be deleted") || d.includes("message can") || d.includes("delete for everyone")) {
        tooOld += 1;
      } else if (d.includes("not found") || d.includes("message to delete")) {
        notFound += 1;
      }
    }
  }
  if (payload?.telegramSessions && typeof payload.telegramSessions === "object") {
    delete payload.telegramSessions[chat];
  }
  if (payload?.telegramLastTaskByChat && typeof payload.telegramLastTaskByChat === "object") {
    delete payload.telegramLastTaskByChat[chat];
  }
  if (payload?.telegramLastSeenMessageByChat && typeof payload.telegramLastSeenMessageByChat === "object") {
    delete payload.telegramLastSeenMessageByChat[chat];
  }
  if (payload?.telegramCloseRequests && typeof payload.telegramCloseRequests === "object") {
    Object.keys(payload.telegramCloseRequests).forEach((taskId) => {
      const req = payload.telegramCloseRequests[taskId];
      if (String(req?.chatId || "").trim() === chat) delete payload.telegramCloseRequests[taskId];
    });
  }
  return {
    deleted,
    failed,
    scanned: ids.length,
    usedAnchors: anchors.length,
    tooOld,
    notFound
  };
}

function refreshEmployeeChatIdsByPhoneBindings(payload) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  const map = payload?.telegramPhoneChatBindings && typeof payload.telegramPhoneChatBindings === "object"
    ? payload.telegramPhoneChatBindings
    : {};
  let connectedCount = 0;
  let clearedCount = 0;
  for (const row of rows) {
    const phone = normalizePhone(row?.[4] || "");
    const mappedChatId = phone ? String(map[phone] || "").trim() : "";
    if (mappedChatId) {
      row[5] = "Подключен";
      row[6] = mappedChatId;
      row[7] = "Активен";
      connectedCount += 1;
      continue;
    }
    row[5] = "Не подключен";
    row[6] = "";
    row[7] = "Не активен";
    clearedCount += 1;
  }
  return { total: rows.length, connectedCount, clearedCount };
}

function getTaskRows(payload) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  const tasks = sections.find((s) => s && s.id === "tasks");
  return Array.isArray(tasks?.rows) ? tasks.rows : [];
}

function getSectionRows(payload, sectionId) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  const section = sections.find((s) => s && s.id === sectionId);
  return Array.isArray(section?.rows) ? section.rows : [];
}

function parseTaskAssigneeNames(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw) return [];
  const out = [];
  const seen = new Set();
  raw.split(",").map((x) => x.trim()).filter(Boolean).forEach((name) => {
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    out.push(name);
  });
  return out;
}

function replaceAssigneeName(rawValue, oldName, newName) {
  const list = parseTaskAssigneeNames(rawValue);
  if (!list.length) return String(newName || "").trim();
  const oldKey = String(oldName || "").trim().toLowerCase();
  const newClean = String(newName || "").trim();
  if (!newClean) return list.join(", ");
  let replaced = false;
  const next = list.map((name) => {
    if (!replaced && String(name || "").trim().toLowerCase() === oldKey) {
      replaced = true;
      return newClean;
    }
    return name;
  });
  if (!replaced) return newClean;
  return Array.from(new Set(next.map((x) => String(x || "").trim()).filter(Boolean))).join(", ");
}

function isReadStateValue(value) {
  const firstLine = String(value || "").split(/\r?\n/)[0].trim().toLowerCase();
  return firstLine.startsWith("прочитано");
}

function hasLastSentValue(value) {
  const s = String(value || "").trim();
  return Boolean(s && s !== "—");
}

function isBlankTaskCell(value) {
  const s = String(value ?? "").trim();
  return s === "" || s === "-" || s === "—" || s.toLowerCase() === "null" || s.toLowerCase() === "undefined";
}

function latestTelegramHistoryTs(store, taskId) {
  const arr = Array.isArray(store?.[taskId]) ? store[taskId] : [];
  let maxTs = 0;
  for (const item of arr) {
    const action = String(item?.action || "");
    if (!action.startsWith("Telegram:")) continue;
    const t = Number(item?.t) || 0;
    if (t > maxTs) maxTs = t;
  }
  return maxTs;
}

function appendTaskHistory(payload, taskId, who, actionText) {
  const id = String(taskId ?? "").trim() || "—";
  const actor = String(who || "").trim() || "Система";
  const action = String(actionText || "").trim() || "—";
  if (!payload.taskHistory || typeof payload.taskHistory !== "object" || Array.isArray(payload.taskHistory)) {
    payload.taskHistory = {};
  }
  if (!Array.isArray(payload.taskHistory[id])) payload.taskHistory[id] = [];
  payload.taskHistory[id].unshift({ t: Date.now(), who: actor, action });
  if (payload.taskHistory[id].length > 300) payload.taskHistory[id].length = 300;
}

function mergeTaskSyncSafeFields(currentPayload, incomingPayload) {
  const next = incomingPayload && typeof incomingPayload === "object"
    ? JSON.parse(JSON.stringify(incomingPayload))
    : incomingPayload;
  const incomingRows = getTaskRows(next);
  const currentRows = getTaskRows(currentPayload);
  if (!currentRows.length) return next;
  const currentById = new Map();
  for (const row of currentRows) {
    const taskId = String(row?.[TASK_NUMBER_COL] || "").trim();
    if (taskId) currentById.set(taskId, row);
  }
  for (const row of incomingRows) {
    const taskId = String(row?.[TASK_NUMBER_COL] || "").trim();
    if (!taskId) continue;
    const currentRow = currentById.get(taskId);
    if (!currentRow) continue;
    const curTgTs = latestTelegramHistoryTs(currentPayload?.taskHistory, taskId);
    const inTgTs = latestTelegramHistoryTs(next?.taskHistory, taskId);
    const currentHasNewerTelegramUpdates = curTgTs > inTgTs;
    if (isReadStateValue(currentRow[TASK_READ_STATE_COL]) && !isReadStateValue(row[TASK_READ_STATE_COL])) {
      row[TASK_READ_STATE_COL] = currentRow[TASK_READ_STATE_COL];
    }
    if (hasLastSentValue(currentRow[TASK_LAST_SENT_AT_COL]) && !hasLastSentValue(row[TASK_LAST_SENT_AT_COL])) {
      row[TASK_LAST_SENT_AT_COL] = currentRow[TASK_LAST_SENT_AT_COL];
    }
    if (String(currentRow[TASK_DELAY_REASON_COL] || "").trim() && !String(row[TASK_DELAY_REASON_COL] || "").trim()) {
      row[TASK_DELAY_REASON_COL] = currentRow[TASK_DELAY_REASON_COL];
    }
    if (currentHasNewerTelegramUpdates) {
      row[TASK_STATUS_COL] = currentRow[TASK_STATUS_COL];
      row[TASK_PLAN_COL] = currentRow[TASK_PLAN_COL];
      row[TASK_CLOSED_DATE_COL] = currentRow[TASK_CLOSED_DATE_COL];
      row[TASK_MEDIA_AFTER_COL] = currentRow[TASK_MEDIA_AFTER_COL];
      row[TASK_READ_STATE_COL] = currentRow[TASK_READ_STATE_COL];
      row[TASK_LAST_SENT_AT_COL] = currentRow[TASK_LAST_SENT_AT_COL];
      row[TASK_DELAY_REASON_COL] = currentRow[TASK_DELAY_REASON_COL];
    }

    // Защита от "обнуления" при частичном импорте/синхронизации:
    // если входящее значение пустое, а текущее заполнено — сохраняем текущее.
    [
      TASK_STATUS_COL,
      TASK_PRIORITY_COL,
      TASK_ADDED_DATE_COL,
      TASK_PHASE_COL,
      TASK_SECTION_COL,
      TASK_SUBSECTION_COL,
      TASK_TITLE_COL,
      TASK_RESPONSIBLE_COL,
      TASK_ASSIGNED_COL,
      TASK_NOTE_COL,
      TASK_PLAN_COL,
      TASK_CLOSED_DATE_COL,
      TASK_READ_STATE_COL,
      TASK_LAST_SENT_AT_COL,
      TASK_DELAY_REASON_COL,
      TASK_CREATED_BY_COL,
      TASK_CREATED_AT_COL,
      TASK_REASSIGN_REASON_COL
    ].forEach((col) => {
      if (isBlankTaskCell(row[col]) && !isBlankTaskCell(currentRow[col])) {
        row[col] = currentRow[col];
      }
    });
  }

  return next;
}

function authMiddleware(req, res, next) {
  const h = req.headers.authorization || "";
  const m = h.match(/^Bearer\s+(.+)$/i);
  if (!m) {
    return res.status(401).json({ error: "Требуется авторизация" });
  }
  try {
    const payload = jwt.verify(m[1], JWT_SECRET);
    req.user = payload;
    next();
  } catch {
    return res.status(401).json({ error: "Сессия недействительна" });
  }
}

function isAdminUser(req) {
  const u = req.user;
  if (!u) return false;
  if (u.role === "admin") return true;
  if (u.sub === "admin") return true;
  return false;
}

function requireAdmin(req, res, next) {
  if (!isAdminUser(req)) {
    return res.status(403).json({ error: "Недостаточно прав" });
  }
  next();
}

async function ensureMediaStorageDir() {
  await fsp.mkdir(MEDIA_STORAGE_PATH, { recursive: true });
}

function mimeToExt(mime) {
  const m = String(mime || "").toLowerCase();
  if (m === "image/jpeg") return "jpg";
  if (m === "image/png") return "png";
  if (m === "image/webp") return "webp";
  if (m === "image/gif") return "gif";
  if (m === "image/svg+xml") return "svg";
  if (m === "image/bmp") return "bmp";
  if (m === "video/mp4") return "mp4";
  if (m === "video/webm") return "webm";
  if (m === "video/ogg") return "ogv";
  if (m === "video/quicktime") return "mov";
  if (m === "video/x-m4v") return "m4v";
  if (m === "application/pdf") return "pdf";
  return "bin";
}

function extFromFileName(fileName) {
  const src = String(fileName || "").trim();
  if (!src) return "";
  const clean = src.replace(/[?#].*$/, "");
  const dot = clean.lastIndexOf(".");
  if (dot <= 0 || dot >= clean.length - 1) return "";
  const ext = clean.slice(dot + 1).toLowerCase().replace(/[^a-z0-9]/g, "");
  if (!ext) return "";
  const allow = new Set([
    "jpg", "jpeg", "png", "webp", "gif", "svg", "bmp",
    "mp4", "webm", "ogv", "mov", "m4v",
    "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx",
    "txt", "csv", "zip", "rar", "7z"
  ]);
  return allow.has(ext) ? ext : "";
}

function parseDataUrl(dataUrl) {
  const m = /^data:([^;]+);base64,(.+)$/s.exec(String(dataUrl || "").trim());
  if (!m) return null;
  return { mime: String(m[1] || "").trim(), base64: String(m[2] || "").trim() };
}

/** Базовый HTTPS-URL приложения для setWebhook (без завершающего /). */
function getPublicBaseUrl(req) {
  const envUrl = String(process.env.PUBLIC_APP_URL || "").trim().replace(/\/$/, "");
  if (envUrl) return envUrl;
  const railway = String(process.env.RAILWAY_PUBLIC_DOMAIN || "").trim();
  if (railway) return `https://${railway}`;
  const host = String(req.get("x-forwarded-host") || req.get("host") || "").trim();
  if (host) {
    const proto = String(req.get("x-forwarded-proto") || "").trim().split(",")[0] || "https";
    return `${proto}://${host}`;
  }
  return "";
}

function resolveServerSidePhotoRef(rawRef, req, token) {
  const raw = String(rawRef || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  if (raw.startsWith("/")) {
    const base = getPublicBaseUrl(req);
    if (!base) return "";
    return `${base}${raw}`;
  }
  if (raw.startsWith("media/")) {
    const base = getPublicBaseUrl(req);
    if (!base) return "";
    return `${base}/${raw}`;
  }
  if (/^[\w.-]+\.(jpg|jpeg|png|gif|webp|bmp|svg)$/i.test(raw)) {
    const base = getPublicBaseUrl(req);
    if (!base) return "";
    return `${base}/media/${encodeURIComponent(raw)}`;
  }
  if (token && raw.includes("/")) {
    const clean = raw.replace(/^\/+/, "");
    return `https://api.telegram.org/file/bot${token}/${clean}`;
  }
  return "";
}

function fileNameFromUrlOrFallback(url, fallback = "photo.jpg") {
  try {
    const u = new URL(String(url || ""));
    const name = decodeURIComponent(String(u.pathname.split("/").pop() || "").trim());
    return name || fallback;
  } catch {
    return fallback;
  }
}

function detectImageMimeFromBytes(buf) {
  if (!buf || buf.length < 4) return "";
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return "image/jpeg";
  if (
    buf.length >= 8 &&
    buf[0] === 0x89 &&
    buf[1] === 0x50 &&
    buf[2] === 0x4e &&
    buf[3] === 0x47 &&
    buf[4] === 0x0d &&
    buf[5] === 0x0a &&
    buf[6] === 0x1a &&
    buf[7] === 0x0a
  ) return "image/png";
  if (buf.length >= 6) {
    const h = buf.subarray(0, 6).toString("ascii");
    if (h === "GIF87a" || h === "GIF89a") return "image/gif";
  }
  if (
    buf.length >= 12 &&
    buf[0] === 0x52 &&
    buf[1] === 0x49 &&
    buf[2] === 0x46 &&
    buf[3] === 0x46 &&
    buf[8] === 0x57 &&
    buf[9] === 0x45 &&
    buf[10] === 0x42 &&
    buf[11] === 0x50
  ) return "image/webp";
  if (buf.length >= 2 && buf[0] === 0x42 && buf[1] === 0x4d) return "image/bmp";
  const headText = buf.subarray(0, Math.min(buf.length, 256)).toString("utf8").trimStart().toLowerCase();
  if (headText.startsWith("<svg")) return "image/svg+xml";
  return "";
}

const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 40,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Слишком много попыток входа, попробуйте позже" }
});

const apiLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 200,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Слишком много запросов" }
});

const app = express();
app.set("trust proxy", 1);

app.use(
  helmet({
    contentSecurityPolicy: false,
    crossOriginEmbedderPolicy: false
  })
);

app.use("/api/", apiLimiter);
app.use(
  "/media",
  express.static(MEDIA_STORAGE_PATH, {
    index: false,
    maxAge: "30d",
    immutable: true,
    fallthrough: false
  })
);

app.use(express.json({ limit: "50mb" }));

app.get("/health", async (_req, res) => {
  try {
    await pool.query("SELECT 1");
    return res.json({ ok: true, service: "mbc-task-api", db: "up" });
  } catch (e) {
    console.error(e);
    return res.status(503).json({ ok: false, service: "mbc-task-api", db: "down" });
  }
});

app.post("/api/auth/login", loginLimiter, async (req, res) => {
  const phone = normalizePhone(req.body?.phone || "");
  const password = String(req.body?.password || "");

  try {
    const { rows } = await pool.query(
      "SELECT id, phone, password_hash, display_name, role FROM users WHERE phone = $1",
      [phone]
    );
    if (rows.length > 0) {
      const u = rows[0];
      const ok = await bcrypt.compare(password, u.password_hash);
      if (!ok) {
        return res.status(401).json({ error: "Неверный логин или пароль" });
      }
      let effectiveRole = u.role;
      try {
        const { rows: stateRows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
        const payload = stateRows[0]?.payload;
        const byPhone = findEmployeeByPhoneInPayload(payload, normalizePhone(u.phone || phone));
        if (isEmployeeAdminAccessEnabled(payload, byPhone)) {
          effectiveRole = "admin";
        }
      } catch (_) {
        // Если app_state временно недоступен — используем роль из users.
      }
      const token = jwt.sign(
        { sub: String(u.id), role: effectiveRole, name: u.display_name },
        JWT_SECRET,
        { expiresIn: "30d" }
      );
      return res.json({ token, displayName: u.display_name });
    }

    /** Фолбэк-вход сотрудника: пароль = последние 4 цифры телефона из справочника. */
    const { rows: stateRows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const payload = stateRows[0]?.payload;
    const emp = findEmployeeByPhoneInPayload(payload, phone);
    if (emp) {
      const expectedPass = last4DigitsPasswordFromPhone(phone);
      if (expectedPass && password === expectedPass) {
        const displayName = String(emp?.[1] || "").trim() || "Пользователь";
        const role = isEmployeeAdminAccessEnabled(payload, emp) ? "admin" : "user";
        const token = jwt.sign(
          { sub: `emp:${phone}`, role, name: displayName },
          JWT_SECRET,
          { expiresIn: "30d" }
        );
        return res.json({ token, displayName });
      }
    }

    const { rows: cntRows } = await pool.query("SELECT COUNT(*)::int AS c FROM users");
    const userCount = cntRows[0]?.c ?? 0;
    const envPhone = normalizePhone(ADMIN_PHONE);
    if (
      userCount === 0 &&
      envPhone &&
      ADMIN_PASSWORD &&
      phone === envPhone &&
      password === ADMIN_PASSWORD
    ) {
      const token = jwt.sign({ sub: "admin", role: "admin", name: "Пользователь" }, JWT_SECRET, { expiresIn: "30d" });
      return res.json({ token, displayName: "Пользователь" });
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сервера" });
  }

  return res.status(401).json({ error: "Неверный логин или пароль" });
});

app.get("/api/auth/me", authMiddleware, (req, res) => {
  const name = String(req.user?.name || "").trim() || "Пользователь";
  const role = req.user?.role === "admin" || req.user?.sub === "admin" ? "admin" : req.user?.role || "user";
  const id = req.user?.sub != null ? String(req.user.sub) : "";
  return res.json({ id, displayName: name, role });
});

app.get("/api/data", authMiddleware, async (_req, res) => {
  try {
    const { rows } = await pool.query("SELECT payload, updated_at FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload ?? null;
    const rev = rows[0]?.updated_at ? Date.parse(rows[0].updated_at) : 0;
    return res.json({ data: payload, rev: Number.isFinite(rev) ? rev : 0 });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка чтения данных" });
  }
});

app.post("/api/telegram/webhook", express.json(), async (req, res) => {
  await handleTelegramWebhook(req, res, pool);
});

/** После сохранения токена в приложении: зарегистрировать webhook на этом домене. */
app.post("/api/telegram/set-webhook", authMiddleware, async (req, res) => {
  try {
    const fromBody = String(req.body?.publicBaseUrl || "").trim().replace(/\/$/, "");
    const base = fromBody || getPublicBaseUrl(req);
    const result = await configureTelegramWebhook(pool, base);
    if (!result.ok) {
      return res.status(400).json({ error: result.error || result.description || "setWebhook failed" });
    }
    return res.json({
      ok: true,
      webhookUrl: result.webhookUrl,
      botUsername: result.botUsername || "",
      botDisplayName: result.botDisplayName || ""
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сервера" });
  }
});

app.post("/api/employees/refresh-chat-ids", authMiddleware, async (_req, res) => {
  try {
    const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload && typeof rows[0].payload === "object"
      ? JSON.parse(JSON.stringify(rows[0].payload))
      : {};
    const stats = refreshEmployeeChatIdsByPhoneBindings(payload);
    await saveAppPayload(payload);
    return res.json({ ok: true, ...stats });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Не удалось обновить Chat ID сотрудников." });
  }
});

app.post("/api/telegram/send-photo-proxy", authMiddleware, async (req, res) => {
  try {
    const chatId = String(req.body?.chatId || "").trim();
    const token = String(req.body?.token || "").trim();
    const photoRefRaw = String(req.body?.photoRef || "").trim();
    const caption = String(req.body?.caption || "");
    const replyMarkup = req.body?.replyMarkup && typeof req.body.replyMarkup === "object" ? req.body.replyMarkup : null;
    if (!chatId) return res.status(400).json({ error: "chatId обязателен" });
    if (!token) return res.status(400).json({ error: "token обязателен" });
    const photoRef = resolveServerSidePhotoRef(photoRefRaw, req, token);
    if (!photoRef) {
      return res.status(400).json({ error: "Не удалось сформировать ссылку на фото" });
    }

    const src = await fetch(photoRef, { method: "GET" });
    if (!src.ok) {
      return res.status(400).json({ error: `Источник фото недоступен: HTTP ${src.status}` });
    }
    const ab = await src.arrayBuffer();
    const buf = Buffer.from(ab);
    const headerType = String(src.headers.get("content-type") || "").toLowerCase();
    const sniffedType = detectImageMimeFromBytes(buf);
    const effectiveType = (headerType.startsWith("image/") ? headerType : "") || sniffedType || "image/jpeg";
    if (!headerType.startsWith("image/") && !sniffedType) {
      return res.status(400).json({ error: `Источник не изображение: ${headerType || "unknown"}` });
    }
    const blob = new Blob([buf], { type: effectiveType });
    const form = new FormData();
    form.append("chat_id", chatId);
    form.append("photo", blob, fileNameFromUrlOrFallback(photoRef));
    if (caption && caption.length <= 1024) {
      form.append("caption", caption);
    }
    if (replyMarkup) {
      form.append("reply_markup", JSON.stringify(replyMarkup));
    }

    const tgResp = await fetch(`https://api.telegram.org/bot${token}/sendPhoto`, {
      method: "POST",
      body: form
    });
    const tgJson = await tgResp.json().catch(() => ({}));
    if (!(tgResp.ok && tgJson?.ok === true)) {
      return res.status(400).json({
        ok: false,
        error: String(tgJson?.description || `Telegram sendPhoto error ${tgResp.status}`)
      });
    }
    return res.json({ ok: true, result: tgJson.result || null });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: `Ошибка прокси-отправки фото: ${String(e?.message || e)}` });
  }
});

app.post("/api/telegram/send-media-group-proxy", authMiddleware, async (req, res) => {
  try {
    const chatId = String(req.body?.chatId || "").trim();
    const token = String(req.body?.token || "").trim();
    const refs = Array.isArray(req.body?.photoRefs) ? req.body.photoRefs : [];
    if (!chatId) return res.status(400).json({ error: "chatId обязателен" });
    if (!token) return res.status(400).json({ error: "token обязателен" });
    if (!refs.length) return res.status(400).json({ error: "photoRefs обязателен" });

    const resolved = refs
      .map((x) => resolveServerSidePhotoRef(String(x || "").trim(), req, token))
      .filter(Boolean)
      .slice(0, 10);
    if (!resolved.length) {
      return res.status(400).json({ error: "Не удалось сформировать ссылки на фото" });
    }

    const media = [];
    const form = new FormData();
    form.append("chat_id", chatId);

    for (let i = 0; i < resolved.length; i += 1) {
      const ref = resolved[i];
      const src = await fetch(ref, { method: "GET" });
      if (!src.ok) {
        return res.status(400).json({ error: `Источник фото #${i + 1} недоступен: HTTP ${src.status}` });
      }
      const ab = await src.arrayBuffer();
      const buf = Buffer.from(ab);
      const headerType = String(src.headers.get("content-type") || "").toLowerCase();
      const sniffedType = detectImageMimeFromBytes(buf);
      const effectiveType = (headerType.startsWith("image/") ? headerType : "") || sniffedType || "image/jpeg";
      if (!headerType.startsWith("image/") && !sniffedType) {
        return res.status(400).json({ error: `Источник #${i + 1} не изображение: ${headerType || "unknown"}` });
      }

      const attachName = `file${i}`;
      media.push({ type: "photo", media: `attach://${attachName}` });
      form.append(attachName, new Blob([buf], { type: effectiveType }), fileNameFromUrlOrFallback(ref, `photo-${i + 1}.jpg`));
    }

    form.append("media", JSON.stringify(media));
    const tgResp = await fetch(`https://api.telegram.org/bot${token}/sendMediaGroup`, {
      method: "POST",
      body: form
    });
    const tgJson = await tgResp.json().catch(() => ({}));
    if (!(tgResp.ok && tgJson?.ok === true)) {
      return res.status(400).json({
        ok: false,
        error: String(tgJson?.description || `Telegram sendMediaGroup error ${tgResp.status}`)
      });
    }

    const result = Array.isArray(tgJson.result) ? tgJson.result : [];
    const firstMessageId = Number(result[0]?.message_id) || null;
    return res.json({ ok: true, result, firstMessageId });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: `Ошибка прокси-отправки группы фото: ${String(e?.message || e)}` });
  }
});

app.post("/api/tasks/reassign/decision", authMiddleware, requireAdmin, async (req, res) => {
  try {
    const requestId = String(req.body?.requestId || "").trim();
    const decision = String(req.body?.decision || "").trim().toLowerCase();
    if (!requestId) return res.status(400).json({ ok: false, error: "requestId обязателен" });
    if (!["approve", "reject"].includes(decision)) {
      return res.status(400).json({ ok: false, error: "decision должен быть approve/reject" });
    }

    const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload && typeof rows[0].payload === "object"
      ? JSON.parse(JSON.stringify(rows[0].payload))
      : {};
    const reqStore = ensureObjectStore(payload, "telegramReassignRequests");
    const reqEntry = reqStore[requestId];
    if (!reqEntry) return res.status(404).json({ ok: false, error: "Заявка не найдена" });
    if (String(reqEntry.status || "").trim() !== "pending") {
      return res.json({ ok: true, alreadyProcessed: true, status: reqEntry.status || "unknown" });
    }

    const tasks = getTaskRows(payload);
    const taskId = String(reqEntry.taskId || "").trim();
    const row = tasks.find((r) => String(r?.[TASK_NUMBER_COL] || "").trim() === taskId);
    if (!row) return res.status(404).json({ ok: false, error: "Задача не найдена" });

    const actorName = String(req.user?.name || "").trim() || "Администратор";
    const nowIso = new Date().toISOString();
    const logStore = ensureObjectStore(payload, "taskReassignLog");
    if (!Array.isArray(logStore[taskId])) logStore[taskId] = [];

    if (decision === "approve") {
      const fromName = String(reqEntry.fromEmployeeName || "").trim();
      const toName = String(reqEntry.toEmployeeName || "").trim();
      row[TASK_STATUS_COL] = "Передано";
      row[TASK_REASSIGN_REASON_COL] = String(reqEntry.reasonText || "").trim();
      reqEntry.status = "approved";
      reqEntry.decidedAt = nowIso;
      reqEntry.decidedBy = actorName;
      appendTaskHistory(payload, taskId, actorName, `Переназначение подтверждено: ${fromName || "—"} → ${toName || "—"}`);
    } else {
      reqEntry.status = "rejected";
      reqEntry.decidedAt = nowIso;
      reqEntry.decidedBy = actorName;
      appendTaskHistory(payload, taskId, actorName, "Переназначение отклонено");
    }

    logStore[taskId].unshift({
      id: requestId,
      code: String(reqEntry.code || `${taskId}/1`).trim(),
      status: reqEntry.status,
      currentStatus: String(reqEntry.status === "approved" ? "В процессе" : "").trim(),
      from: reqEntry.fromEmployeeName || "",
      to: reqEntry.toEmployeeName || "",
      reasonType: reqEntry.reasonType || "",
      reasonText: reqEntry.reasonText || "",
      department: reqEntry.departmentName || "",
      requestedBy: reqEntry.requesterName || "",
      decidedBy: actorName,
      createdAt: reqEntry.createdAt || nowIso,
      decidedAt: nowIso
    });
    if (logStore[taskId].length > 100) logStore[taskId].length = 100;

    await saveAppPayload(payload);
    return res.json({ ok: true, status: reqEntry.status, taskId });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка обработки переназначения" });
  }
});

app.post("/api/telegram/chat-clear/request", authMiddleware, requireAdmin, async (req, res) => {
  try {
    const employeeId = String(req.body?.employeeId || "").trim();
    const payload = await loadAppPayload();
    const token = String(payload?.displaySettings?.telegramBotToken || "").trim();
    if (!token) return res.status(400).json({ ok: false, error: "Не задан токен Telegram-бота." });

    const targetEmployee = employeeId ? findEmployeeByIdInPayload(payload, employeeId) : null;
    if (!targetEmployee) {
      return res.status(404).json({ ok: false, error: "Сотрудник не найден." });
    }
    const targetChatId = String(targetEmployee?.[6] || "").trim();
    const targetTg = String(targetEmployee?.[5] || "").trim();
    if (!targetChatId || targetTg !== "Подключен") {
      return res.status(400).json({ ok: false, error: "У сотрудника нет активного Telegram Chat ID." });
    }

    const initiator = await resolveInitiatorEmployeeRow(payload, req.user);
    const initiatorChatId = String(initiator?.[6] || "").trim();
    const initiatorTg = String(initiator?.[5] || "").trim();
    if (!initiatorChatId || initiatorTg !== "Подключен") {
      return res.status(400).json({
        ok: false,
        error: "Ваш Telegram не подключен в справочнике сотрудников. Подключите его и повторите."
      });
    }

    const code = String(Math.floor(100000 + Math.random() * 900000));
    const requestId = randomBytes(12).toString("hex");
    const expiresAt = Date.now() + CHAT_CLEAR_CODE_TTL_MS;
    const targetName = String(targetEmployee?.[1] || "").trim() || "—";
    const initiatorName = String(initiator?.[1] || req.user?.name || "").trim() || "Администратор";
    const msg = [
      "🔐 Подтверждение очистки чата",
      "",
      `Инициатор: ${initiatorName}`,
      `Сотрудник: ${targetName}`,
      `Chat ID: ${targetChatId}`,
      `Код: ${code}`,
      "Срок действия: 10 минут"
    ].join("\n");
    const sendResult = await tgSendMessage(token, initiatorChatId, msg);
    if (!sendResult.ok) {
      return res.status(400).json({
        ok: false,
        error: `Не удалось отправить код в Telegram: ${sendResult.description || "ошибка API"}`
      });
    }

    const reqStore = ensureObjectStore(payload, "telegramChatClearRequests");
    reqStore[requestId] = {
      code,
      createdAt: Date.now(),
      expiresAt,
      createdBySub: String(req.user?.sub || ""),
      createdByChatId: initiatorChatId,
      targetEmployeeId: String(targetEmployee?.[0] || ""),
      targetEmployeeName: targetName,
      targetChatId
    };
    await saveAppPayload(payload);
    return res.json({ ok: true, requestId, expiresAt });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка создания запроса на очистку чата." });
  }
});

app.post("/api/telegram/chat-clear/confirm", authMiddleware, requireAdmin, async (req, res) => {
  try {
    const requestId = String(req.body?.requestId || "").trim();
    const code = String(req.body?.code || "").trim();
    if (!requestId || !/^\d{6}$/.test(code)) {
      return res.status(400).json({ ok: false, error: "Укажите корректный код подтверждения." });
    }
    const payload = await loadAppPayload();
    const token = String(payload?.displaySettings?.telegramBotToken || "").trim();
    if (!token) return res.status(400).json({ ok: false, error: "Не задан токен Telegram-бота." });
    const reqStore = ensureObjectStore(payload, "telegramChatClearRequests");
    const pending = reqStore[requestId];
    if (!pending) return res.status(404).json({ ok: false, error: "Запрос подтверждения не найден." });
    const ownerSub = String(pending?.createdBySub || "").trim();
    if (ownerSub !== String(req.user?.sub || "").trim()) {
      return res.status(403).json({ ok: false, error: "Этот код выдан другому администратору." });
    }
    if (Date.now() > Number(pending?.expiresAt || 0)) {
      delete reqStore[requestId];
      await saveAppPayload(payload);
      return res.status(400).json({ ok: false, error: "Срок действия кода истек. Запросите новый код." });
    }
    if (String(pending?.code || "") !== code) {
      return res.status(400).json({ ok: false, error: "Неверный код подтверждения." });
    }
    const targetChatId = String(pending?.targetChatId || "").trim();
    if (!targetChatId) {
      delete reqStore[requestId];
      await saveAppPayload(payload);
      return res.status(400).json({ ok: false, error: "У сотрудника отсутствует Chat ID для очистки." });
    }

    const result = await clearTelegramChatMessagesBestEffort(payload, token, targetChatId);
    delete reqStore[requestId];
    await saveAppPayload(payload);
    return res.json({ ok: true, ...result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка очистки чата." });
  }
});

app.post("/api/google-sheets/sync", authMiddleware, async (_req, res) => {
  try {
    const result = await runGoogleSheetsSync(pool, { mode: "manual" });
    if (!result.ok && result.busy) {
      return res.status(409).json({ ok: false, error: "Синхронизация уже выполняется." });
    }
    if (!result.ok) {
      return res.status(400).json({ ok: false, error: result.error || "Не удалось синхронизировать Google Sheets." });
    }
    return res.json({ ok: true, ...result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка ручной синхронизации Google Sheets." });
  }
});

app.post("/api/media/upload", authMiddleware, async (req, res) => {
  try {
    await ensureMediaStorageDir();
    const dataUrl = String(req.body?.dataUrl || "").trim();
    const sourceFileName = String(req.body?.fileName || "").trim();
    const parsed = parseDataUrl(dataUrl);
    if (!parsed) {
      return res.status(400).json({ error: "Неверный формат файла (ожидается data URL)." });
    }
    if (parsed.base64.length > 24 * 1024 * 1024) {
      return res.status(413).json({ error: "Файл слишком большой (максимум ~18MB)." });
    }
    const ext = extFromFileName(sourceFileName) || mimeToExt(parsed.mime);
    const fileName = `${Date.now()}-${randomBytes(6).toString("hex")}.${ext}`;
    const absPath = path.join(MEDIA_STORAGE_PATH, fileName);
    const buf = Buffer.from(parsed.base64, "base64");
    await fsp.writeFile(absPath, buf);
    const base = getPublicBaseUrl(req);
    const url = `${base}/media/${encodeURIComponent(fileName)}`;
    return res.json({
      ok: true,
      url,
      fileName,
      mime: parsed.mime
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: `Ошибка загрузки медиа: ${String(e?.message || e)}` });
  }
});

app.put("/api/data", authMiddleware, async (req, res) => {
  try {
    const incomingData = req.body?.data;
    if (incomingData === undefined) {
      return res.status(400).json({ error: "Поле data обязательно" });
    }
    const v = validateAppPayload(incomingData);
    if (!v.ok) {
      return res.status(400).json({ error: v.error || "Некорректные данные" });
    }
    const data = typeof incomingData === "object" && incomingData ? { ...incomingData } : incomingData;
    const baseRev = Number(req.body?.baseRev) || 0;
    const { rows: currentRows } = await pool.query("SELECT payload, updated_at FROM app_state WHERE id = 1");
    const currentPayload = currentRows[0]?.payload && typeof currentRows[0].payload === "object"
      ? currentRows[0].payload
      : {};
    const currentRev = currentRows[0]?.updated_at ? Date.parse(currentRows[0].updated_at) : 0;
    if (baseRev > 0 && Number.isFinite(currentRev) && currentRev > 0 && baseRev < currentRev) {
      return res.status(409).json({
        error: "stale_data",
        message: "Данные устарели. Обновите таблицу и повторите действие.",
        data: currentPayload,
        rev: currentRev
      });
    }
    const mergedData = mergeTaskSyncSafeFields(currentPayload, data);
    // Служебные поля Telegram живут только на сервере. На клиентском PUT /api/data
    // мы всегда сохраняем их из текущего payload в БД (без доверия клиентскому слепку),
    // чтобы гонки между webhook и авто-sync не ломали сценарии комментариев/фото.
    const { rows: savedRows } = await pool.query(
      `INSERT INTO app_state (id, payload, updated_at)
       VALUES (1, $1::jsonb, NOW())
       ON CONFLICT (id) DO UPDATE SET
         payload = jsonb_set(
           jsonb_set(
             jsonb_set(
               jsonb_set(
                 jsonb_set(
                   jsonb_set(
                     jsonb_set(
                       jsonb_set(
                         jsonb_set(
                           EXCLUDED.payload,
                           '{telegramSessions}',
                           COALESCE(app_state.payload->'telegramSessions', '{}'::jsonb),
                           true
                         ),
                         '{telegramPhoneChatBindings}',
                         COALESCE(app_state.payload->'telegramPhoneChatBindings', '{}'::jsonb),
                         true
                       ),
                       '{telegramCloseRequests}',
                       COALESCE(app_state.payload->'telegramCloseRequests', '{}'::jsonb),
                       true
                     ),
                     '{telegramLastTaskByChat}',
                     COALESCE(app_state.payload->'telegramLastTaskByChat', '{}'::jsonb),
                     true
                   ),
                   '{telegramChatClearRequests}',
                   COALESCE(app_state.payload->'telegramChatClearRequests', '{}'::jsonb),
                   true
                 ),
                 '{telegramLastSeenMessageByChat}',
                 COALESCE(app_state.payload->'telegramLastSeenMessageByChat', '{}'::jsonb),
                 true
               ),
               '{telegramReassignRequests}',
               COALESCE(app_state.payload->'telegramReassignRequests', '{}'::jsonb),
               true
             ),
             '{taskReassignLog}',
             COALESCE(app_state.payload->'taskReassignLog', '{}'::jsonb),
             true
           ),
           '{taskCloseMeta}',
           COALESCE(app_state.payload->'taskCloseMeta', '{}'::jsonb),
           true
         ),
         updated_at = NOW()
       RETURNING updated_at`,
      [JSON.stringify(mergedData)]
    );
    const savedRev = savedRows[0]?.updated_at ? Date.parse(savedRows[0].updated_at) : Date.now();
    return res.json({ ok: true, rev: Number.isFinite(savedRev) ? savedRev : Date.now() });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сохранения" });
  }
});

app.post("/api/share", authMiddleware, async (req, res) => {
  try {
    const pin = String(req.body?.pin || "");
    const expiresAt = Number(req.body?.expiresAt);
    const rows = req.body?.rows;
    if (!/^\d{4}$/.test(pin) || !Number.isFinite(expiresAt) || !Array.isArray(rows)) {
      return res.status(400).json({ error: "Некорректные данные" });
    }
    const id = randomBytes(16).toString("hex");
    await pool.query(`INSERT INTO report_shares (id, pin, expires_at, rows) VALUES ($1, $2, $3, $4::jsonb)`, [
      id,
      pin,
      expiresAt,
      JSON.stringify(rows)
    ]);
    return res.json({ id });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Не удалось создать ссылку" });
  }
});

app.get("/api/share/:id/meta", async (req, res) => {
  try {
    const { rows } = await pool.query("SELECT expires_at FROM report_shares WHERE id = $1", [req.params.id]);
    if (!rows.length) return res.status(404).json({ error: "not_found" });
    return res.json({ expiresAt: Number(rows[0].expires_at) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "server" });
  }
});

app.post("/api/share/:id/verify", async (req, res) => {
  try {
    const pin = String(req.body?.pin || "").replace(/\D/g, "").slice(0, 4);
    const { rows } = await pool.query("SELECT pin, expires_at, rows FROM report_shares WHERE id = $1", [req.params.id]);
    if (!rows.length) return res.status(404).json({ error: "not_found" });
    const r = rows[0];
    if (Date.now() > Number(r.expires_at)) {
      return res.status(410).json({ error: "expired" });
    }
    if (pin !== r.pin) {
      return res.status(403).json({ error: "bad_pin" });
    }
    return res.json({ rows: r.rows, expiresAt: Number(r.expires_at) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "server" });
  }
});

app.get("/api/admin/users", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const { rows } = await pool.query(
      "SELECT id, phone, display_name, role, created_at FROM users ORDER BY id ASC"
    );
    return res.json({ users: rows });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сервера" });
  }
});

app.post("/api/admin/users", authMiddleware, requireAdmin, async (req, res) => {
  try {
    const phone = normalizePhone(req.body?.phone || "");
    const password = String(req.body?.password || "");
    const displayName = String(req.body?.displayName || "Пользователь").trim() || "Пользователь";
    const role = req.body?.role === "admin" ? "admin" : "user";
    if (!phone || phone.length < 5) {
      return res.status(400).json({ error: "Укажите корректный телефон" });
    }
    if (password.length < 6) {
      return res.status(400).json({ error: "Пароль не короче 6 символов" });
    }
    const passwordHash = await bcrypt.hash(password, 11);
    const { rows } = await pool.query(
      `INSERT INTO users (phone, password_hash, display_name, role)
       VALUES ($1, $2, $3, $4)
       RETURNING id, phone, display_name, role, created_at`,
      [phone, passwordHash, displayName, role]
    );
    return res.status(201).json({ user: rows[0] });
  } catch (e) {
    if (e.code === "23505") {
      return res.status(409).json({ error: "Пользователь с таким телефоном уже есть" });
    }
    console.error(e);
    return res.status(500).json({ error: "Не удалось создать пользователя" });
  }
});

app.delete("/api/admin/users/:id", authMiddleware, requireAdmin, async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) {
      return res.status(400).json({ error: "Некорректный id" });
    }
    const selfId = Number(req.user?.sub);
    if (Number.isFinite(selfId) && selfId === id) {
      return res.status(400).json({ error: "Нельзя удалить свою учётную запись" });
    }
    const { rows: admins } = await pool.query("SELECT COUNT(*)::int AS c FROM users WHERE role = $1", ["admin"]);
    const { rows: target } = await pool.query("SELECT role FROM users WHERE id = $1", [id]);
    if (!target.length) {
      return res.status(404).json({ error: "Не найден" });
    }
    if (target[0].role === "admin" && admins[0].c <= 1) {
      return res.status(400).json({ error: "Нельзя удалить последнего администратора" });
    }
    await pool.query("DELETE FROM users WHERE id = $1", [id]);
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сервера" });
  }
});

app.use(express.static(rootDir, { index: false }));

app.get("*", (_req, res) => {
  res.sendFile(path.join(rootDir, "index.html"));
});

const PORT = Number(process.env.PORT) || 3000;

async function seedAdminFromEnv() {
  const { rows } = await pool.query("SELECT COUNT(*)::int AS c FROM users");
  if (rows[0].c > 0) {
    return;
  }
  const phone = normalizePhone(ADMIN_PHONE);
  const password = String(ADMIN_PASSWORD || "");
  if (!phone || password.length < 4) {
    console.warn(
      "В таблице users нет записей: задайте ADMIN_PHONE и ADMIN_PASSWORD в окружении — при следующем старте будет создан первый администратор."
    );
    return;
  }
  const hash = await bcrypt.hash(password, 11);
  const displayName = String(process.env.ADMIN_DISPLAY_NAME || "Администратор").trim() || "Администратор";
  await pool.query(
    `INSERT INTO users (phone, password_hash, display_name, role) VALUES ($1, $2, $3, 'admin')`,
    [phone, hash, displayName]
  );
  console.log("Создан первый администратор из ADMIN_PHONE / ADMIN_PASSWORD.");
}

async function main() {
  if (!process.env.DATABASE_URL) {
    console.error("Ошибка: задайте переменную DATABASE_URL (PostgreSQL в Railway → Connect).");
    process.exit(1);
  }
  await runMigrations(pool);
  await seedAdminFromEnv();
  startGoogleSheetsAutoSync(pool);
  console.log("Миграции и сид пользователей выполнены.");

  if (JWT_SECRET === "change-me-in-production" && NODE_ENV === "production") {
    console.error("Задайте JWT_SECRET в production.");
    process.exit(1);
  }
  if (JWT_SECRET === "change-me-in-production") {
    console.warn("Задайте JWT_SECRET в переменных окружения перед production.");
  }
  await ensureMediaStorageDir();
  console.log(`Папка медиа: ${MEDIA_STORAGE_PATH}`);

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Сервер слушает порт ${PORT} (${NODE_ENV})`);
  });
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

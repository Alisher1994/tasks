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
const TASK_DUE_DATE_COL = 14;
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
const OVERDUE_DIGEST_INTERVAL_MS = 60 * 1000;
const SMS_INVITE_LOG_LIMIT = 500;
const SMS_DEFAULT_TIMEOUT_MS = 15000;
const SMS_DEFAULT_PROVIDER = "generic";
const SMS_GATE_PROVIDER_ID = "sms-gate.app";
const SMS_DEFAULT_PHONE_FIELD = "phone";
const SMS_DEFAULT_MESSAGE_FIELD = "message";
const SMS_DEFAULT_SENDER_FIELD = "sender";
const SMS_DEFAULT_API_KEY_HEADER = "Authorization";
const SMS_DEFAULT_AUTH_TYPE = "header";
const SMS_DEFAULT_GATE_URL = "https://api.sms-gate.app/3rdparty/v1/messages";
const SMS_DEFAULT_INVITE_TEMPLATE =
  "Здравствуйте, [ФИО]. Пожалуйста, пройдите регистрацию в Telegram-боте [Бот] по ссылке: [Ссылка_бота]. После регистрации вы будете получать задачи от руководителей.";
let overdueDigestTickInFlight = false;
let overdueDigestTimer = null;

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

async function tgSendMessageWithMarkup(token, chatId, text, replyMarkup = null) {
  const body = { chat_id: String(chatId || "").trim(), text: String(text || "") };
  if (replyMarkup && typeof replyMarkup === "object") body.reply_markup = replyMarkup;
  const r = await fetch(buildBotApiUrl(token, "sendMessage"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const j = await r.json().catch(() => ({}));
  return {
    ok: r.ok && j?.ok === true,
    description: String(j?.description || ""),
    messageId: Number(j?.result?.message_id) || 0
  };
}

function truncateForLog(raw, max = 1200) {
  const text = String(raw || "").trim();
  if (!text) return "";
  if (text.length <= max) return text;
  return `${text.slice(0, max - 1)}…`;
}

function normalizeSmsGatewayMethod(raw) {
  const value = String(raw || "").trim().toLowerCase();
  return value === "get-query" ? "get-query" : "post-json";
}

function normalizeSmsGatewayProvider(raw) {
  const value = String(raw || "").trim().toLowerCase();
  if (value === SMS_GATE_PROVIDER_ID) return SMS_GATE_PROVIDER_ID;
  return SMS_DEFAULT_PROVIDER;
}

function normalizeSmsGatewayAuthType(raw) {
  const value = String(raw || "").trim().toLowerCase();
  if (value === "basic") return "basic";
  if (value === "bearer") return "bearer";
  if (value === "none") return "none";
  return SMS_DEFAULT_AUTH_TYPE;
}

function normalizeSmsFieldKey(raw, fallback) {
  const value = String(raw || "").trim();
  if (!value) return fallback;
  if (!/^[A-Za-z_][A-Za-z0-9_.-]{0,63}$/.test(value)) return fallback;
  return value;
}

function normalizeSmsHeaderName(raw, fallback = SMS_DEFAULT_API_KEY_HEADER) {
  const value = String(raw || "").trim();
  if (!value) return fallback;
  if (!/^[A-Za-z0-9-]{1,64}$/.test(value)) return fallback;
  return value;
}

function normalizeSmsInviteBoolean(raw, fallback = false) {
  if (typeof raw === "boolean") return raw;
  const value = String(raw || "").trim().toLowerCase();
  if (!value) return fallback;
  if (["1", "true", "yes", "on", "да"].includes(value)) return true;
  if (["0", "false", "no", "off", "нет"].includes(value)) return false;
  return fallback;
}

function ensureSmsInviteLogArray(payload) {
  if (!Array.isArray(payload.smsInviteLog)) payload.smsInviteLog = [];
  return payload.smsInviteLog;
}

function buildSmsInviteBotLink(displaySettings, employeeId) {
  const botUsername = String(displaySettings?.telegramBotUsername || "").trim().replace(/^@+/, "");
  if (!botUsername) return "";
  const id = String(employeeId || "").trim();
  if (!id) return `https://t.me/${botUsername}`;
  return `https://t.me/${botUsername}?start=e_${encodeURIComponent(id)}`;
}

function applySmsInviteTemplate(rawTemplate, context) {
  let tpl = String(rawTemplate || "").trim();
  if (!tpl) tpl = SMS_DEFAULT_INVITE_TEMPLATE;
  const replacements = new Map([
    ["[ФИО]", context.fullName || "сотрудник"],
    ["[ID]", context.employeeId || ""],
    ["[Телефон]", context.phone || ""],
    ["[Бот]", context.botLabel || "бот"],
    ["[Ссылка_бота]", context.botLink || "ссылка будет предоставлена администратором"],
    ["[full_name]", context.fullName || "employee"],
    ["[employee_id]", context.employeeId || ""],
    ["[phone]", context.phone || ""],
    ["[bot]", context.botLabel || "bot"],
    ["[bot_link]", context.botLink || ""]
  ]);
  let out = tpl;
  replacements.forEach((value, token) => {
    out = out.split(token).join(String(value || ""));
  });
  return truncateForLog(out, 2000);
}

function getSmsGatewaySettings(displaySettings) {
  const ds = displaySettings && typeof displaySettings === "object" ? displaySettings : {};
  const provider = normalizeSmsGatewayProvider(ds.smsGatewayProvider);
  const timeoutRaw = Number(ds.smsGatewayTimeoutMs);
  const simRaw = Number(ds.smsGatewaySimNumber);
  const ttlRaw = Number(ds.smsGatewayTtlSeconds);
  const prioRaw = Number(ds.smsGatewayPriority);
  const deviceWithinRaw = Number(ds.smsGatewayDeviceActiveWithinHours);
  return {
    enabled: ds.smsGatewayEnabled === true,
    provider,
    url: String(ds.smsGatewayUrl || "").trim() || SMS_DEFAULT_GATE_URL,
    method: provider === SMS_GATE_PROVIDER_ID ? "post-json" : normalizeSmsGatewayMethod(ds.smsGatewayMethod),
    authType: normalizeSmsGatewayAuthType(ds.smsGatewayAuthType),
    username: String(ds.smsGatewayUsername || "").trim(),
    password: String(ds.smsGatewayPassword || "").trim(),
    apiKey: String(ds.smsGatewayApiKey || "").trim(),
    apiKeyHeader: normalizeSmsHeaderName(ds.smsGatewayApiKeyHeader, SMS_DEFAULT_API_KEY_HEADER),
    sender: String(ds.smsGatewaySender || "").trim(),
    deviceId: String(ds.smsGatewayDeviceId || "").trim(),
    simNumber: Number.isFinite(simRaw) && simRaw >= 1 ? Math.floor(simRaw) : 0,
    ttlSeconds: Number.isFinite(ttlRaw) && ttlRaw > 0 ? Math.floor(ttlRaw) : 0,
    priority: Number.isFinite(prioRaw) ? Math.max(0, Math.min(100, Math.floor(prioRaw))) : 0,
    skipPhoneValidation: normalizeSmsInviteBoolean(ds.smsGatewaySkipPhoneValidation, true),
    deviceActiveWithinHours: Number.isFinite(deviceWithinRaw) && deviceWithinRaw > 0
      ? Math.min(720, Math.max(1, Math.floor(deviceWithinRaw)))
      : 0,
    phoneField: normalizeSmsFieldKey(ds.smsGatewayPhoneField, SMS_DEFAULT_PHONE_FIELD),
    messageField: normalizeSmsFieldKey(ds.smsGatewayMessageField, SMS_DEFAULT_MESSAGE_FIELD),
    senderField: normalizeSmsFieldKey(ds.smsGatewaySenderField, SMS_DEFAULT_SENDER_FIELD),
    timeoutMs: Number.isFinite(timeoutRaw)
      ? Math.min(60000, Math.max(3000, Math.floor(timeoutRaw)))
      : SMS_DEFAULT_TIMEOUT_MS,
    inviteTemplate: String(ds.smsInviteTemplate || "").trim()
  };
}

function buildSmsInviteMessage(payload, employeeRow) {
  const displaySettings = payload?.displaySettings || {};
  const settings = getSmsGatewaySettings(displaySettings);
  const employeeId = String(employeeRow?.[0] || "").trim();
  const fullName = String(employeeRow?.[1] || "").trim() || "Сотрудник";
  const phone = normalizePhone(employeeRow?.[4] || "");
  const botUsername = String(displaySettings?.telegramBotUsername || "").trim().replace(/^@+/, "");
  const botLabel = botUsername ? `@${botUsername}` : "Telegram-бот";
  const botLink = buildSmsInviteBotLink(displaySettings, employeeId);
  const text = applySmsInviteTemplate(settings.inviteTemplate, {
    employeeId,
    fullName,
    phone,
    botLabel,
    botLink
  });
  return { text, botLink, botLabel };
}

async function sendSmsViaGateway(settings, { phone, text }) {
  const url = String(settings?.url || "").trim();
  if (!url) {
    return { ok: false, reason: "Не указан URL SMS Gateway." };
  }
  const provider = normalizeSmsGatewayProvider(settings?.provider);
  const method = normalizeSmsGatewayMethod(settings?.method);
  const phoneField = normalizeSmsFieldKey(settings?.phoneField, SMS_DEFAULT_PHONE_FIELD);
  const messageField = normalizeSmsFieldKey(settings?.messageField, SMS_DEFAULT_MESSAGE_FIELD);
  const senderField = normalizeSmsFieldKey(settings?.senderField, SMS_DEFAULT_SENDER_FIELD);
  const sender = String(settings?.sender || "").trim();
  const timeoutMs = Number.isFinite(Number(settings?.timeoutMs))
    ? Math.min(60000, Math.max(3000, Math.floor(Number(settings.timeoutMs))))
    : SMS_DEFAULT_TIMEOUT_MS;

  const headers = {};
  const authType = normalizeSmsGatewayAuthType(settings?.authType);
  const username = String(settings?.username || "").trim();
  const password = String(settings?.password || "").trim();
  const apiKey = String(settings?.apiKey || "").trim();
  if (provider === SMS_GATE_PROVIDER_ID) {
    if (authType === "basic") {
      if (!username || !password) {
        return { ok: false, reason: "Для sms-gate.app (Basic) укажите username и password." };
      }
      const credentials = Buffer.from(`${username}:${password}`, "utf8").toString("base64");
      headers.Authorization = `Basic ${credentials}`;
    } else if (authType === "bearer") {
      if (!apiKey) {
        return { ok: false, reason: "Для sms-gate.app (Bearer) укажите токен." };
      }
      headers.Authorization = apiKey.toLowerCase().startsWith("bearer ")
        ? apiKey
        : `Bearer ${apiKey}`;
    } else if (authType === "header" && apiKey) {
      const keyHeader = normalizeSmsHeaderName(settings?.apiKeyHeader, SMS_DEFAULT_API_KEY_HEADER);
      headers[keyHeader] = apiKey;
    }
  } else if (authType === "basic") {
    if (!username || !password) {
      return { ok: false, reason: "Для Basic авторизации укажите username и password." };
    }
    const credentials = Buffer.from(`${username}:${password}`, "utf8").toString("base64");
    headers.Authorization = `Basic ${credentials}`;
  } else if (authType === "bearer") {
    if (!apiKey) {
      return { ok: false, reason: "Для Bearer авторизации укажите токен." };
    }
    headers.Authorization = apiKey.toLowerCase().startsWith("bearer ")
      ? apiKey
      : `Bearer ${apiKey}`;
  } else if (apiKey) {
    const keyHeader = normalizeSmsHeaderName(settings?.apiKeyHeader, SMS_DEFAULT_API_KEY_HEADER);
    headers[keyHeader] = apiKey;
  }

  const phoneNormalized = String(phone || "").trim();
  const messageText = String(text || "");
  let payload = {
    [phoneField]: phoneNormalized,
    [messageField]: messageText
  };
  if (sender) payload[senderField] = sender;

  const params = new URLSearchParams();
  if (provider === SMS_GATE_PROVIDER_ID) {
    payload = {
      textMessage: { text: messageText },
      phoneNumbers: [phoneNormalized]
    };
    const deviceId = String(settings?.deviceId || "").trim();
    const simNumber = Number(settings?.simNumber) || 0;
    const ttlSeconds = Number(settings?.ttlSeconds) || 0;
    const priority = Number(settings?.priority) || 0;
    const deviceActiveWithinHours = Number(settings?.deviceActiveWithinHours) || 0;
    const skipPhoneValidation = normalizeSmsInviteBoolean(settings?.skipPhoneValidation, true);
    if (deviceId) payload.deviceId = deviceId;
    if (simNumber > 0) payload.simNumber = simNumber;
    if (ttlSeconds > 0) payload.ttl = ttlSeconds;
    if (priority > 0) payload.priority = priority;
    params.set("skipPhoneValidation", skipPhoneValidation ? "true" : "false");
    if (deviceActiveWithinHours > 0) {
      params.set("deviceActiveWithin", String(deviceActiveWithinHours));
    }
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  let response;
  try {
    const requestUrl = params.size > 0
      ? `${url}${url.includes("?") ? "&" : "?"}${params.toString()}`
      : url;
    if (method === "get-query") {
      const u = new URL(requestUrl);
      Object.entries(payload).forEach(([k, v]) => {
        u.searchParams.set(k, String(v || ""));
      });
      response = await fetch(u.toString(), {
        method: "GET",
        headers,
        signal: controller.signal
      });
    } else {
      headers["Content-Type"] = "application/json";
      response = await fetch(requestUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
        signal: controller.signal
      });
    }
  } catch (e) {
    clearTimeout(timer);
    if (String(e?.name || "") === "AbortError") {
      return { ok: false, reason: `Таймаут SMS Gateway (${timeoutMs} мс).` };
    }
    return { ok: false, reason: String(e?.message || "Ошибка сети при отправке SMS.") };
  }
  clearTimeout(timer);

  const responseText = await response.text().catch(() => "");
  let parsed = null;
  try {
    parsed = responseText ? JSON.parse(responseText) : null;
  } catch (_) {
    parsed = null;
  }
  const providerMessage = truncateForLog(
    parsed?.message
      || parsed?.description
      || parsed?.status
      || parsed?.error
      || responseText
      || "",
    800
  );

  if (!response.ok) {
    return {
      ok: false,
      httpStatus: response.status,
      reason: providerMessage || `HTTP ${response.status}`,
      responsePreview: truncateForLog(responseText, 1200)
    };
  }

  return {
    ok: true,
    httpStatus: response.status,
    providerMessage: providerMessage || (provider === SMS_GATE_PROVIDER_ID ? "Запрос принят SMS Gate." : ""),
    responsePreview: truncateForLog(responseText, 1200)
  };
}

function normalizePersonName(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function findEmployeeByFullNameInPayload(payload, fullName) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  const want = normalizePersonName(fullName).toLowerCase();
  if (!want) return null;
  return rows.find((row) => normalizePersonName(row?.[EMPLOYEE_FULL_NAME_COL] || "").toLowerCase() === want) || null;
}

function getDuplicateRecipientChatIds(payload) {
  const ds = payload?.displaySettings || {};
  const allowIds = new Set(
    Array.isArray(ds.telegramGlobalDuplicateRecipientIds)
      ? ds.telegramGlobalDuplicateRecipientIds.map((x) => String(x || "").trim()).filter(Boolean)
      : []
  );
  if (!allowIds.size) return [];
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  const out = new Set();
  for (const row of rows) {
    const id = String(row?.[0] || "").trim();
    if (!id || !allowIds.has(id)) continue;
    if (String(row?.[EMPLOYEE_TELEGRAM_COL] || "").trim() !== "Подключен") continue;
    const chatId = String(row?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
    if (chatId) out.add(chatId);
  }
  return Array.from(out);
}

function collectTaskTelegramRecipientNames(row) {
  const out = new Set();
  parseTaskAssigneeNames(row?.[TASK_ASSIGNED_COL] || "").forEach((x) => out.add(x));
  const responsible = normalizePersonName(row?.[TASK_RESPONSIBLE_COL] || "");
  if (responsible) out.add(responsible);
  return Array.from(out);
}

function resolveTaskRecipientChatIds(payload, row) {
  const chats = new Set();
  const recipientNames = collectTaskTelegramRecipientNames(row);
  recipientNames.forEach((name) => {
    const emp = findEmployeeByFullNameInPayload(payload, name);
    if (!emp) return;
    if (String(emp?.[EMPLOYEE_TELEGRAM_COL] || "").trim() !== "Подключен") return;
    const chatId = String(emp?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
    if (chatId) chats.add(chatId);
  });
  getDuplicateRecipientChatIds(payload).forEach((chatId) => chats.add(chatId));
  return Array.from(chats);
}

function resolveServerTimeZone(payload) {
  const tz = String(payload?.displaySettings?.serverTimezone || "").trim();
  return tz || "UTC";
}

function formatYmdInTimeZone(date, timeZone) {
  const dtf = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  });
  return dtf.format(date);
}

function getMinutesInTimeZone(date, timeZone) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).formatToParts(date);
  const hh = Number(parts.find((p) => p.type === "hour")?.value || 0);
  const mm = Number(parts.find((p) => p.type === "minute")?.value || 0);
  return hh * 60 + mm;
}

function parseRuDateToYmd(rawValue) {
  const s = String(rawValue || "").trim();
  const m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!m) return null;
  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);
  if (!Number.isFinite(day) || !Number.isFinite(month) || !Number.isFinite(year)) return null;
  if (month < 1 || month > 12) return null;
  if (day < 1 || day > 31) return null;
  return { year, month, day };
}

function compareYmd(a, b) {
  if (a.year !== b.year) return a.year - b.year;
  if (a.month !== b.month) return a.month - b.month;
  return a.day - b.day;
}

function normalizeOverdueNotifyTimeValue(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}):(\d{1,2})$/);
  if (!match) return "09:00";
  const h = Math.max(0, Math.min(23, Number(match[1]) || 0));
  const m = Math.max(0, Math.min(59, Number(match[2]) || 0));
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function buildOverdueDigestText(totalCount) {
  const count = Math.max(0, Number(totalCount) || 0);
  return [
    "⚠️ Уведомление",
    "",
    `Количество просроченных задач: ${count} шт.`,
    "Нажмите «Посмотреть», чтобы открыть список."
  ].join("\n");
}

function buildOverdueDigestInlineKeyboard() {
  return {
    inline_keyboard: [[{ text: "📋 Посмотреть", callback_data: "ov|ls" }]]
  };
}

async function runOverdueDigestSchedulerTick() {
  if (overdueDigestTickInFlight) return;
  overdueDigestTickInFlight = true;
  try {
    const payload = await loadAppPayload();
    const ds = payload?.displaySettings || {};
    if (ds.overdueNotificationsEnabled !== true) return;
    const token = String(ds.telegramBotToken || "").trim();
    if (!token) return;

    const tz = resolveServerTimeZone(payload);
    const now = new Date();
    const nowMinutes = getMinutesInTimeZone(now, tz);
    const [hh, mm] = normalizeOverdueNotifyTimeValue(ds.overdueNotificationsTime).split(":").map(Number);
    const targetMinutes = hh * 60 + mm;
    if (nowMinutes < targetMinutes) return;

    const runtimeStore = ensureObjectStore(payload, "serverSchedulers");
    const overdueRuntime = runtimeStore.overdueDigest && typeof runtimeStore.overdueDigest === "object"
      ? runtimeStore.overdueDigest
      : {};
    const today = formatYmdInTimeZone(now, tz);
    if (String(overdueRuntime.lastRunDate || "").trim() === today) return;

    const tasks = getTaskRows(payload);
    const todayRu = new Intl.DateTimeFormat("ru-RU", {
      timeZone: tz,
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    }).format(now);
    const todayParts = parseRuDateToYmd(todayRu);
    if (!todayParts) return;

    const byChat = new Map();
    for (const row of tasks) {
      if (!Array.isArray(row)) continue;
      const status = String(row[TASK_STATUS_COL] || "").trim();
      if (status === "Закрыт") continue;
      const due = parseRuDateToYmd(String(row[TASK_DUE_DATE_COL] || "").trim());
      if (!due) continue;
      const diff = compareYmd(todayParts, due);
      if (diff <= 0) continue;
      const taskId = String(row[TASK_NUMBER_COL] || "").trim();
      if (!taskId) continue;
      const chatIds = resolveTaskRecipientChatIds(payload, row);
      chatIds.forEach((chatId) => {
        if (!byChat.has(chatId)) byChat.set(chatId, new Set());
        byChat.get(chatId).add(taskId);
      });
    }

    for (const [chatId, taskIds] of byChat.entries()) {
      const text = buildOverdueDigestText(taskIds.size);
      await tgSendMessageWithMarkup(token, chatId, text, buildOverdueDigestInlineKeyboard());
    }

    runtimeStore.overdueDigest = {
      lastRunDate: today,
      lastRunAt: now.toISOString(),
      timeZone: tz,
      sentChats: byChat.size
    };
    await saveAppPayload(payload);
  } catch (e) {
    console.error("overdue digest scheduler", e);
  } finally {
    overdueDigestTickInFlight = false;
  }
}

function startOverdueDigestScheduler() {
  clearInterval(overdueDigestTimer);
  overdueDigestTimer = setInterval(() => {
    runOverdueDigestSchedulerTick().catch(() => {});
  }, OVERDUE_DIGEST_INTERVAL_MS);
  runOverdueDigestSchedulerTick().catch(() => {});
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

function latestHistoryTs(store, taskId) {
  const list = Array.isArray(store?.[String(taskId || "").trim()]) ? store[String(taskId || "").trim()] : [];
  let maxTs = 0;
  for (const item of list) {
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
    const curAnyTs = latestHistoryTs(currentPayload?.taskHistory, taskId);
    const inAnyTs = latestHistoryTs(next?.taskHistory, taskId);
    const incomingHasNewerManualEdits = inAnyTs > curAnyTs;
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
      if (!incomingHasNewerManualEdits) {
        row[TASK_MEDIA_AFTER_COL] = currentRow[TASK_MEDIA_AFTER_COL];
      }
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
    const token = String(payload?.displaySettings?.telegramBotToken || "").trim();
    const logStore = ensureObjectStore(payload, "taskReassignLog");
    if (!Array.isArray(logStore[taskId])) logStore[taskId] = [];

    if (decision === "approve") {
      const fromName = String(reqEntry.fromEmployeeName || "").trim();
      const toName = String(reqEntry.toEmployeeName || "").trim();
      const reassignCode = String(reqEntry.code || `${taskId}/1`).trim();
      row[TASK_STATUS_COL] = "Передано";
      row[TASK_REASSIGN_REASON_COL] = String(reqEntry.reasonText || "").trim();
      reqEntry.status = "approved";
      reqEntry.decidedAt = nowIso;
      reqEntry.decidedBy = actorName;
      appendTaskHistory(payload, taskId, actorName, `Переназначение подтверждено: ${fromName || "—"} → ${toName || "—"}`);
      if (token) {
        const requester = findEmployeeByFullNameInPayload(payload, String(reqEntry.requesterName || "").trim());
        const target = findEmployeeByFullNameInPayload(payload, toName);
        const requesterChat = String(requester?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
        const targetChat = String(target?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
        const notifyTextRequester = [
          `✅ Переназначение подтверждено`,
          `Задача: ${reassignCode}`,
          `Кому: ${toName || "—"}`,
          `Подтвердил: ${actorName}`,
          `Дата/время: ${nowIso}`
        ].join("\n");
        const notifyTextTarget = [
          `📌 Вам назначена задача (переназначение)`,
          `Задача: ${reassignCode}`,
          `От кого: ${fromName || "—"}`,
          `Причина: ${String(reqEntry.reasonText || "").trim() || "—"}`,
          `Подтвердил: ${actorName}`
        ].join("\n");
        if (requesterChat) {
          await tgSendMessage(token, requesterChat, notifyTextRequester);
        }
        if (targetChat) {
          await tgSendMessage(token, targetChat, notifyTextTarget);
        }
      }
    } else {
      reqEntry.status = "rejected";
      reqEntry.decidedAt = nowIso;
      reqEntry.decidedBy = actorName;
      appendTaskHistory(payload, taskId, actorName, "Переназначение отклонено");
      if (token) {
        const requester = findEmployeeByFullNameInPayload(payload, String(reqEntry.requesterName || "").trim());
        const requesterChat = String(requester?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
        const reassignCode = String(reqEntry.code || `${taskId}/1`).trim();
        if (requesterChat) {
          const rejectText = [
            `❌ Переназначение отклонено`,
            `Задача: ${reassignCode}`,
            `Отклонил: ${actorName}`,
            `Дата/время: ${nowIso}`
          ].join("\n");
          await tgSendMessage(token, requesterChat, rejectText);
        }
      }
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

app.post("/api/sms/invite/send", authMiddleware, requireAdmin, async (req, res) => {
  const nowMs = Date.now();
  let payload = null;
  let logEntry = null;
  try {
    const employeeId = String(req.body?.employeeId || "").trim();
    if (!employeeId) {
      return res.status(400).json({ ok: false, error: "employeeId обязателен." });
    }
    payload = await loadAppPayload();
    const employeeRow = findEmployeeByIdInPayload(payload, employeeId);
    if (!employeeRow) {
      return res.status(404).json({ ok: false, error: "Сотрудник не найден." });
    }

    const fullName = String(employeeRow?.[EMPLOYEE_FULL_NAME_COL] || "").trim() || "Сотрудник";
    const phone = normalizePhone(employeeRow?.[4] || "");
    const chatId = String(employeeRow?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
    const smsSettings = getSmsGatewaySettings(payload?.displaySettings || {});
    const inviteMessage = buildSmsInviteMessage(payload, employeeRow);
    const actorName = String(req.user?.name || "").trim() || "Администратор";

    logEntry = {
      id: randomBytes(10).toString("hex"),
      atMs: nowMs,
      atIso: new Date(nowMs).toISOString(),
      employeeId,
      employeeName: fullName,
      phone,
      chatId,
      text: truncateForLog(inviteMessage.text, 2000),
      actor: actorName,
      status: "failed",
      ok: false,
      gatewayProvider: smsSettings.provider,
      gatewayMethod: smsSettings.method,
      gatewayUrl: truncateForLog(smsSettings.url, 240),
      httpStatus: 0,
      resultMessage: "",
      responsePreview: ""
    };

    if (chatId) {
      logEntry.resultMessage = "Сотрудник уже активирован (есть Chat ID), SMS не требуется.";
    } else if (!phone) {
      logEntry.resultMessage = "У сотрудника не указан корректный телефон.";
    } else if (!smsSettings.enabled) {
      logEntry.resultMessage = "SMS Gateway выключен в настройках.";
    } else if (!smsSettings.url) {
      logEntry.resultMessage = "Не указан URL SMS Gateway в настройках.";
    } else if (!inviteMessage.text) {
      logEntry.resultMessage = "Текст SMS пустой. Проверьте шаблон.";
    } else {
      const sendResult = await sendSmsViaGateway(smsSettings, { phone, text: inviteMessage.text });
      logEntry.ok = sendResult.ok === true;
      logEntry.status = sendResult.ok ? "sent" : "failed";
      logEntry.httpStatus = Number(sendResult.httpStatus) || 0;
      logEntry.resultMessage = truncateForLog(
        sendResult.providerMessage || sendResult.reason || (sendResult.ok ? "SMS отправлено." : "Неизвестная ошибка."),
        800
      );
      logEntry.responsePreview = truncateForLog(sendResult.responsePreview || "", 1200);
    }

    const logStore = ensureSmsInviteLogArray(payload);
    logStore.unshift(logEntry);
    if (logStore.length > SMS_INVITE_LOG_LIMIT) {
      logStore.length = SMS_INVITE_LOG_LIMIT;
    }
    await saveAppPayload(payload);

    if (logEntry.ok) {
      return res.json({ ok: true, entry: logEntry });
    }
    return res.status(400).json({ ok: false, error: logEntry.resultMessage || "Не удалось отправить SMS.", entry: logEntry });
  } catch (e) {
    console.error(e);
    try {
      if (payload && logEntry) {
        logEntry.resultMessage = truncateForLog(`Ошибка сервера: ${String(e?.message || e)}`, 800);
        logEntry.status = "failed";
        logEntry.ok = false;
        const logStore = ensureSmsInviteLogArray(payload);
        logStore.unshift(logEntry);
        if (logStore.length > SMS_INVITE_LOG_LIMIT) logStore.length = SMS_INVITE_LOG_LIMIT;
        await saveAppPayload(payload);
      }
    } catch (_) {
      /* noop */
    }
    return res.status(500).json({ ok: false, error: "Ошибка отправки SMS-приглашения." });
  }
});

app.get("/api/sms/invite/history", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const payload = await loadAppPayload();
    const list = Array.isArray(payload?.smsInviteLog)
      ? payload.smsInviteLog.filter((item) => item && typeof item === "object")
      : [];
    return res.json({ ok: true, entries: list });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Не удалось загрузить историю SMS." });
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
    // Служебные поля Telegram и журналы отправок (SMS) живут только на сервере.
    // На клиентском PUT /api/data мы всегда сохраняем их из текущего payload в БД
    // (без доверия клиентскому слепку), чтобы гонки между webhook/endpoint и авто-sync
    // не ломали сценарии комментариев/фото/истории отправок.
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
            '{serverSchedulers}',
            COALESCE(app_state.payload->'serverSchedulers', '{}'::jsonb),
            true
          ),
          '{smsInviteLog}',
          COALESCE(app_state.payload->'smsInviteLog', '[]'::jsonb),
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
  startOverdueDigestScheduler();
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

/**
 * API + раздача статики для Railway.
 * Переменные: DATABASE_URL, JWT_SECRET, ADMIN_PHONE, ADMIN_PASSWORD (первый админ), PORT
 */
import express from "express";
import { promises as fsp } from "fs";
import os from "os";
import path from "path";
import { fileURLToPath } from "url";
import jwt from "jsonwebtoken";
import pg from "pg";
import bcrypt from "bcryptjs";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import { randomBytes } from "crypto";
import { createServer as createHttpServer } from "http";
import { configureTelegramWebhook, handleTelegramWebhook, clearCloseConfirmPrompts } from "./telegramWebhook.js";
import { runGoogleSheetsSync, startGoogleSheetsAutoSync } from "./googleSheetsSync.js";
import { runMigrations } from "./migrate.js";
import { validateAppPayload } from "./validatePayload.js";
import { attachRealtimeHub, broadcastStateChanged, getRealtimeConnectionCount } from "./realtimeHub.js";
import { ensureMediaStorageDir, validateMediaUpload } from "./mediaUpload.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.join(__dirname, "..");
const distDir = path.join(rootDir, "dist");
const { existsSync } = await import("fs");
const staticRoot = existsSync(distDir) ? distDir : rootDir;

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
let previousCpuSnapshot = null;
let previousNetworkSnapshot = null;
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
const TASK_REASSIGN_TYPE_COL = 24;
const EMPLOYEE_FULL_NAME_COL = 1;
const EMPLOYEE_PHONE_COL = 4;
const EMPLOYEE_TELEGRAM_COL = 5;
const EMPLOYEE_CHAT_ID_COL = 6;
const EMPLOYEE_LAST_ACTIVITY_COL = 7;
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
const SMS_DEFAULT_TASK_TEMPLATE =
  "У вас есть задача №[ID_задачи]: [Название_задачи]. Для подробностей перейдите в Telegram-бот [Бот]: [Ссылка_бота].";
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

function findEmployeeForAuthUserInPayload(payload, user) {
  if (!user) return null;
  const byPhone = findEmployeeByPhoneInPayload(payload, normalizePhone(user.phone || ""));
  if (byPhone) return byPhone;
  const sub = String(user.sub || "").trim();
  const employeeId = sub.startsWith("emp:") ? "" : sub;
  const byId = findEmployeeByIdInPayload(payload, employeeId);
  if (byId) return byId;
  return findEmployeeByDisplayNameInPayload(payload, user.name || "");
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
    "доступкнастройкам",
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

function resolveEffectiveRoleFromPayload(payload, user, fallbackRole = "user") {
  const normalizedFallbackRole = String(fallbackRole || "").trim().toLowerCase() === "admin" ? "admin" : "user";
  if (normalizedFallbackRole === "admin") return "admin";
  const employee = findEmployeeForAuthUserInPayload(payload, user);
  return isEmployeeAdminAccessEnabled(payload, employee) ? "admin" : normalizedFallbackRole;
}

function touchEmployeeLastLoginInPayload(payload, employeeRow, at = new Date()) {
  if (!Array.isArray(employeeRow)) return false;
  employeeRow[EMPLOYEE_LAST_ACTIVITY_COL] = at.toISOString();
  return true;
}

async function recordEmployeeLastLogin({ phone = "", displayName = "" } = {}) {
  try {
    const payload = await loadAppPayload();
    const normalizedPhone = normalizePhone(phone);
    const row = normalizedPhone
      ? findEmployeeByPhoneInPayload(payload, normalizedPhone)
      : findEmployeeByDisplayNameInPayload(payload, displayName);
    if (!row) return false;
    touchEmployeeLastLoginInPayload(payload, row);
    await saveAppPayload(payload);
    return true;
  } catch (e) {
    console.warn("Не удалось записать последнюю активность сотрудника:", e?.message || e);
    return false;
  }
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

function getPayloadSection(payload, sectionId) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  return sections.find((section) => section && String(section.id || "") === sectionId) || null;
}

function sectionRowsAsObjects(section) {
  const columns = Array.isArray(section?.columns) ? section.columns.map((x) => String(x || "")) : [];
  const rows = Array.isArray(section?.rows) ? section.rows : [];
  return rows.map((row) => {
    const out = {};
    columns.forEach((column, index) => {
      out[column || `col_${index}`] = Array.isArray(row) ? row[index] ?? "" : "";
    });
    return out;
  });
}

function buildExportSectionPayload(payload, sectionId) {
  const section = getPayloadSection(payload, sectionId);
  if (!section) return null;
  return {
    id: String(section.id || sectionId),
    title: String(section.title || sectionId),
    columns: Array.isArray(section.columns) ? section.columns : [],
    rows: Array.isArray(section.rows) ? section.rows : [],
    items: sectionRowsAsObjects(section)
  };
}

function buildCatalogsExportPayload(payload) {
  const ids = ["data", "phases", "phaseSections", "phaseSubsections", "delayReasons", "roles", "departments"];
  return Object.fromEntries(
    ids
      .map((id) => [id, buildExportSectionPayload(payload, id)])
      .filter(([, value]) => Boolean(value))
  );
}

function sanitizeDisplaySettingsForExport(settings) {
  const out = settings && typeof settings === "object" && !Array.isArray(settings)
    ? JSON.parse(JSON.stringify(settings))
    : {};
  [
    "telegramBotToken",
    "smsGatewayPassword",
    "smsGatewayApiKey",
    "googleSheetsSpreadsheetId"
  ].forEach((key) => {
    if (key in out) out[key] = "";
  });
  return out;
}

function takeCpuSnapshot() {
  const cpus = os.cpus();
  let idle = 0;
  let total = 0;
  cpus.forEach((cpu) => {
    idle += cpu.times.idle;
    total += Object.values(cpu.times).reduce((sum, value) => sum + value, 0);
  });
  return { idle, total };
}

function getCpuUsagePercent() {
  const current = takeCpuSnapshot();
  const previous = previousCpuSnapshot;
  previousCpuSnapshot = current;
  if (!previous) return 0;
  const idleDelta = current.idle - previous.idle;
  const totalDelta = current.total - previous.total;
  if (totalDelta <= 0) return 0;
  return Math.max(0, Math.min(100, ((totalDelta - idleDelta) / totalDelta) * 100));
}

async function getDiskStats(targetPath) {
  try {
    const stats = await fsp.statfs(targetPath);
    const blockSize = Number(stats.bsize) || 0;
    const total = Number(stats.blocks) * blockSize;
    const free = Number(stats.bfree) * blockSize;
    const available = Number(stats.bavail) * blockSize;
    const used = Math.max(0, total - free);
    const usedPercent = total > 0 ? (used / total) * 100 : 0;
    return { path: targetPath, total, used, free, available, usedPercent };
  } catch (_) {
    return { path: targetPath, total: 0, used: 0, free: 0, available: 0, usedPercent: 0 };
  }
}

function getNetworkInterfaceSummary() {
  return Object.entries(os.networkInterfaces())
    .map(([name, items]) => {
      const active = (items || []).filter((item) => !item.internal);
      if (!active.length) return null;
      return {
        name,
        addresses: active.map((item) => item.address).filter(Boolean),
        families: [...new Set(active.map((item) => item.family).filter(Boolean))]
      };
    })
    .filter(Boolean);
}

async function getNetworkStats() {
  const interfaces = getNetworkInterfaceSummary();
  try {
    const raw = await fsp.readFile("/proc/net/dev", "utf8");
    let rxBytes = 0;
    let txBytes = 0;
    raw.split(/\r?\n/).forEach((line) => {
      const parts = line.trim().split(/\s+/);
      if (parts.length < 17 || !parts[0].endsWith(":")) return;
      const name = parts[0].slice(0, -1);
      if (name === "lo") return;
      rxBytes += Number(parts[1]) || 0;
      txBytes += Number(parts[9]) || 0;
    });
    const now = Date.now();
    let rxPerSec = 0;
    let txPerSec = 0;
    if (previousNetworkSnapshot && now > previousNetworkSnapshot.at) {
      const seconds = (now - previousNetworkSnapshot.at) / 1000;
      rxPerSec = Math.max(0, (rxBytes - previousNetworkSnapshot.rxBytes) / seconds);
      txPerSec = Math.max(0, (txBytes - previousNetworkSnapshot.txBytes) / seconds);
    }
    previousNetworkSnapshot = { at: now, rxBytes, txBytes };
    return { interfaces, rxBytes, txBytes, rxPerSec, txPerSec, supported: true };
  } catch (_) {
    return { interfaces, rxBytes: 0, txBytes: 0, rxPerSec: 0, txPerSec: 0, supported: false };
  }
}

async function buildServerMetricsPayload() {
  const memoryTotal = os.totalmem();
  const memoryFree = os.freemem();
  const memoryUsed = Math.max(0, memoryTotal - memoryFree);
  const processMemory = process.memoryUsage();
  const disk = await getDiskStats(rootDir);
  const mediaDisk = MEDIA_STORAGE_PATH === rootDir ? disk : await getDiskStats(MEDIA_STORAGE_PATH);
  const network = await getNetworkStats();
  return {
    ok: true,
    at: Date.now(),
    server: {
      nodeEnv: NODE_ENV,
      platform: `${os.type()} ${os.release()}`,
      arch: os.arch(),
      hostname: os.hostname(),
      nodeVersion: process.version,
      uptimeSec: os.uptime(),
      processUptimeSec: process.uptime()
    },
    cpu: {
      usagePercent: getCpuUsagePercent(),
      cores: os.cpus().length,
      model: os.cpus()[0]?.model || ""
    },
    memory: {
      total: memoryTotal,
      used: memoryUsed,
      free: memoryFree,
      usedPercent: memoryTotal > 0 ? (memoryUsed / memoryTotal) * 100 : 0,
      processRss: processMemory.rss,
      processHeapUsed: processMemory.heapUsed,
      processHeapTotal: processMemory.heapTotal
    },
    disk,
    mediaDisk,
    network,
    process: {
      pid: process.pid,
      pool: {
        total: pool.totalCount,
        idle: pool.idleCount,
        waiting: pool.waitingCount
      }
    },
    realtime: {
      connections: getRealtimeConnectionCount()
    }
  };
}

function buildOpenApiSpec(req) {
  const baseUrl = `${req.protocol}://${req.get("host")}`;
  const jsonResponse = {
    type: "object",
    properties: {
      ok: { type: "boolean", example: true }
    }
  };
  const sectionResponse = {
    type: "object",
    properties: {
      ok: { type: "boolean", example: true },
      section: {
        type: "object",
        properties: {
          id: { type: "string", example: "tasks" },
          title: { type: "string", example: "Задачи" },
          columns: { type: "array", items: { type: "string" } },
          rows: { type: "array", items: { type: "array", items: {} } },
          items: { type: "array", items: { type: "object" } }
        }
      }
    }
  };
  const makeGet = (summary, description, schema = sectionResponse) => ({
    get: {
      tags: ["Методы API"],
      summary,
      description,
      security: [{ bearerAuth: [] }],
      responses: {
        200: {
          description: "Успешный ответ",
          content: {
            "application/json": {
              schema
            }
          }
        },
        401: { description: "Нет или неверный Bearer-токен" },
        403: { description: "Недостаточно прав: нужен администратор" }
      }
    }
  });
  return {
    openapi: "3.0.3",
    info: {
      title: "MBC Task Management API",
      version: "1.0.0",
      description: [
        "Read-only API для безопасной передачи данных из системы.",
        "Все методы требуют права администратора и заголовок Authorization: Bearer <JWT>.",
        "JWT можно взять из авторизованной сессии приложения или получить через POST /api/auth/login."
      ].join("\n")
    },
    servers: [{ url: baseUrl }],
    tags: [
      {
        name: "Методы API",
        description: "Доступные методы API"
      }
    ],
    components: {
      securitySchemes: {
        bearerAuth: {
          type: "http",
          scheme: "bearer",
          bearerFormat: "JWT"
        }
      }
    },
    paths: {
      "/api/auth/login": {
        post: {
          tags: ["Методы API"],
          summary: "Вход и получение JWT",
          description: "Передайте phone и password. В ответе вернётся token, который нужно вставить в Authorize как Bearer token.",
          requestBody: {
            required: true,
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  required: ["phone", "password"],
                  properties: {
                    phone: { type: "string", example: "+998991234567" },
                    password: { type: "string", example: "123456" }
                  }
                }
              }
            }
          },
          responses: {
            200: {
              description: "JWT токен",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: {
                      token: { type: "string" },
                      displayName: { type: "string" },
                      role: { type: "string", example: "admin" }
                    }
                  }
                }
              }
            },
            401: { description: "Неверный логин или пароль" }
          }
        }
      },
      "/api/export/tasks": makeGet("Экспорт задач", "Возвращает таблицу задач: columns, исходные rows и удобные items-объекты."),
      "/api/export/employees": makeGet("Экспорт сотрудников", "Возвращает сотрудников, включая отдел, должность, Telegram-статус и Chat ID."),
      "/api/export/objects": makeGet("Экспорт объектов", "Возвращает объекты, адреса, РП/ЗРП и связанные поля."),
      "/api/export/catalogs": makeGet("Экспорт справочников", "Возвращает справочники фаз, разделов, подразделов, причин, ролей, отделов и ответственных.", {
        ...jsonResponse,
        properties: {
          ok: { type: "boolean", example: true },
          catalogs: { type: "object" }
        }
      }),
      "/api/export/all": makeGet("Экспорт всех данных", "Возвращает все sections в исходном виде и основные настройки без изменения данных.", {
        ...jsonResponse,
        properties: {
          ok: { type: "boolean", example: true },
          sections: { type: "array", items: { type: "object" } },
          displaySettings: { type: "object" }
        }
      }),
      "/api/admin/server/metrics": makeGet("Метрики сервера", "Возвращает текущие параметры сервера: CPU, память, диск, сеть, процесс Node.js и пул PostgreSQL.", {
        ...jsonResponse,
        properties: {
          ok: { type: "boolean", example: true },
          at: { type: "number", example: 1715200000000 },
          server: { type: "object" },
          cpu: { type: "object" },
          memory: { type: "object" },
          disk: { type: "object" },
          mediaDisk: { type: "object" },
          network: { type: "object" },
          process: { type: "object" }
        }
      })
    }
  };
}

function renderSwaggerDocsHtml() {
  return `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>MBC API Docs</title>
  <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist@5/swagger-ui.css" />
  <style>
    :root {
      --mbc-blue: #3e4095;
      --mbc-gold: #d1ae6c;
      --mbc-ink: #17213a;
      --mbc-muted: #60708b;
      --mbc-line: #dce4ee;
      --mbc-soft: #f4f7fb;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: linear-gradient(180deg, #eef3f9 0, #f8fafc 360px);
      color: var(--mbc-ink);
    }
    body::before {
      content: "";
      display: block;
      height: 6px;
      background: linear-gradient(90deg, var(--mbc-blue) 0 72%, var(--mbc-gold) 72% 100%);
    }
    .docs-gate { max-width: 760px; margin: 48px auto; padding: 20px; font: 14px/1.5 Inter, Arial, sans-serif; color: #26364e; background: #fff; border: 1px solid #dce4ee; border-radius: 12px; }
    .docs-gate code { background: #eef2f7; padding: 2px 5px; border-radius: 5px; }
    .swagger-ui { font-family: Inter, Arial, sans-serif; color: var(--mbc-ink); }
    .swagger-ui .information-container {
      max-width: none;
      margin: 0;
      padding: 0;
      background:
        linear-gradient(135deg, #ffffff 0%, #f4f7fb 100%);
      border-bottom: 1px solid var(--mbc-line);
    }
    .swagger-ui .info {
      position: relative;
      display: grid;
      grid-template-columns: minmax(0, 1fr) 178px;
      gap: 24px;
      max-width: 1380px;
      min-height: 150px;
      margin: 0 auto;
      padding: 24px 32px 26px;
      align-items: center;
    }
    .swagger-ui .info::before {
      content: "";
      display: block;
      grid-column: 2;
      grid-row: 1 / span 3;
      width: 172px;
      height: 64px;
      justify-self: end;
      background: url("/horizontal-v1.svg") right center / contain no-repeat;
    }
    .swagger-ui .info > * { grid-column: 1; }
    .swagger-ui .info .title {
      display: flex;
      flex-wrap: wrap;
      gap: 9px;
      align-items: center;
      margin: 0 0 6px;
      color: var(--mbc-ink);
      font-size: 32px;
      line-height: 1.12;
      letter-spacing: 0;
    }
    .swagger-ui .info .title::before {
      content: none;
    }
    .swagger-ui .info .title small {
      top: 0;
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      padding: 3px 8px;
      border-radius: 7px;
      background: #7b8798;
      color: #fff;
      font-size: 12px;
      font-weight: 800;
    }
    .swagger-ui .info .title small.version-stamp {
      background: var(--mbc-blue);
    }
    .swagger-ui .info .title small pre {
      color: inherit;
      font-family: inherit;
      font-weight: inherit;
    }
    .swagger-ui .info .base-url,
    .swagger-ui .info .link {
      display: none;
    }
    .swagger-ui .info .description {
      max-width: 980px;
      margin-top: 12px;
      padding-left: 16px;
      border-left: 3px solid var(--mbc-gold);
      color: #26364e;
    }
    .swagger-ui .info .description p {
      margin: 0;
      font-size: 15px;
      line-height: 1.45;
    }
    .swagger-ui .scheme-container {
      max-width: none;
      margin: 0 0 26px;
      padding: 24px 32px;
      background: rgba(255, 255, 255, 0.86);
      border-bottom: 1px solid var(--mbc-line);
      box-shadow: none;
    }
    .swagger-ui .scheme-container .schemes {
      max-width: 1380px;
      margin: 0 auto;
      align-items: end;
    }
    .swagger-ui .btn.authorize {
      min-width: 150px;
      border-color: var(--mbc-blue);
      color: var(--mbc-blue);
      border-radius: 7px;
      box-shadow: 0 10px 22px rgba(62, 64, 149, 0.12);
    }
    .swagger-ui .btn.authorize svg { fill: var(--mbc-blue); }
    .swagger-ui .opblock svg,
    .swagger-ui .opblock-summary svg,
    .swagger-ui .opblock-control-arrow,
    .swagger-ui .copy-to-clipboard button svg,
    .swagger-ui .authorization__btn svg,
    .swagger-ui .expand-operation svg,
    .swagger-ui .models-control svg {
      color: var(--mbc-blue);
      fill: var(--mbc-blue);
      stroke: var(--mbc-blue);
      transition: none;
      animation: none;
      transform: none;
      filter: none;
    }
    .swagger-ui .opblock button,
    .swagger-ui .opblock-summary button,
    .swagger-ui .copy-to-clipboard button,
    .swagger-ui .authorization__btn,
    .swagger-ui .expand-operation,
    .swagger-ui .models-control {
      transition: none;
      animation: none;
      transform: none;
      filter: none;
    }
    .swagger-ui .copy-to-clipboard {
      opacity: 1;
      visibility: visible;
    }
    .swagger-ui .copy-to-clipboard button {
      width: 28px;
      height: 28px;
      padding: 0;
      background: transparent;
      border: 0;
      box-shadow: none;
    }
    .swagger-ui .copy-to-clipboard button:hover,
    .swagger-ui .copy-to-clipboard button:focus,
    .swagger-ui .copy-to-clipboard button:active {
      background: transparent;
      border: 0;
      box-shadow: none;
    }
    .swagger-ui .opblock button:hover svg,
    .swagger-ui .opblock-summary button:hover svg,
    .swagger-ui .copy-to-clipboard button:hover svg,
    .swagger-ui .authorization__btn:hover svg,
    .swagger-ui .expand-operation:hover svg,
    .swagger-ui .models-control:hover svg {
      color: var(--mbc-blue);
      fill: var(--mbc-blue);
      stroke: var(--mbc-blue);
      transform: none;
      filter: none;
    }
    .swagger-ui select {
      border-color: #b8c3d4;
      border-radius: 6px;
      color: var(--mbc-ink);
    }
    .swagger-ui .wrapper {
      max-width: 1380px;
      padding: 0 32px;
    }
    @media (max-width: 760px) {
      .swagger-ui .info {
        grid-template-columns: 1fr;
        gap: 18px;
        padding: 22px 18px 26px;
      }
      .swagger-ui .info::before {
        grid-column: 1;
        grid-row: auto;
        justify-self: start;
        width: min(172px, 100%);
        height: 64px;
      }
      .swagger-ui .info > * { grid-column: 1; }
      .swagger-ui .info .title { font-size: 28px; }
      .swagger-ui .scheme-container { padding: 18px; }
      .swagger-ui .wrapper { padding: 0 18px; }
    }
  </style>
</head>
<body>
  <div id="swagger-ui"></div>
  <div id="docsGate" class="docs-gate">Проверяем права администратора...</div>
  <script src="https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
  <script>
    const AUTH_TOKEN_KEY = "mbc_jwt";
    function readTokenFromWindowName() {
      try {
        const parsed = JSON.parse(window.name || "{}");
        return parsed && typeof parsed.mbcSwaggerToken === "string" ? parsed.mbcSwaggerToken : "";
      } catch (_) {
        return "";
      }
    }
    function writeTokenToWindowName(token) {
      try {
        const parsed = JSON.parse(window.name || "{}");
        parsed.mbcSwaggerToken = token || "";
        window.name = JSON.stringify(parsed);
      } catch (_) {
        window.name = JSON.stringify({ mbcSwaggerToken: token || "" });
      }
    }
    function readStoredToken() {
      return readTokenFromWindowName() || sessionStorage.getItem(AUTH_TOKEN_KEY) || localStorage.getItem(AUTH_TOKEN_KEY) || "";
    }
    function saveStoredToken(token) {
      if (!token) return;
      writeTokenToWindowName(token);
      try { sessionStorage.setItem(AUTH_TOKEN_KEY, token); } catch (_) {}
    }
    function isAuthLoginRequest(url) {
      try {
        return new URL(url, window.location.origin).pathname === "/api/auth/login";
      } catch (_) {
        return String(url || "").includes("/api/auth/login");
      }
    }
    function saveTokenFromLoginResponse(res) {
      try {
        const url = res?.url || "";
        if (!isAuthLoginRequest(url) || Number(res?.status) !== 200) return res;
        const raw = typeof res?.text === "string" ? res.text : "";
        const body = raw ? JSON.parse(raw) : res?.body;
        if (body && typeof body.token === "string") saveStoredToken(body.token);
      } catch (_) {}
      return res;
    }
    const token = readStoredToken();
    const gate = document.getElementById("docsGate");
    async function boot() {
      if (!token) {
        gate.innerHTML = "<h2>Нужен вход администратора</h2><p>Откройте основное приложение, войдите администратором, затем вернитесь на <code>/api/docs</code>.</p>";
        return;
      }
      const me = await fetch("/api/auth/me", { headers: { Authorization: "Bearer " + token } }).then(r => r.ok ? r.json() : null).catch(() => null);
      if (!me || me.role !== "admin") {
        gate.innerHTML = "<h2>Недостаточно прав</h2><p>Swagger доступен только пользователю с правами администратора.</p>";
        return;
      }
      gate.remove();
      SwaggerUIBundle({
        url: "/api/openapi.json",
        dom_id: "#swagger-ui",
        persistAuthorization: true,
        requestInterceptor: (req) => {
          if (!isAuthLoginRequest(req.url)) {
            const freshToken = readStoredToken();
            if (freshToken) req.headers.Authorization = "Bearer " + freshToken;
          }
          return req;
        },
        responseInterceptor: saveTokenFromLoginResponse
      });
    }
    boot();
  </script>
</body>
</html>`;
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
  const { rows } = await pool.query(
    `INSERT INTO app_state (id, payload, updated_at, revision)
     VALUES (1, $1::jsonb, NOW(), 1)
     ON CONFLICT (id) DO UPDATE SET
       payload = EXCLUDED.payload,
       updated_at = NOW(),
       revision = app_state.revision + 1
     RETURNING revision`,
    [JSON.stringify(nextPayload)]
  );
  const newRev = Number(rows[0]?.revision) || 0;
  // Real-time: уведомляем всех подключённых клиентов о новой ревизии,
  // чтобы они подтянули свежие данные без ожидания auto-pull.
  try { broadcastStateChanged(newRev, { source: "server" }); } catch {}
  return newRev;
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

function applySmsTaskTemplate(rawTemplate, context) {
  let tpl = String(rawTemplate || "").trim();
  if (!tpl) tpl = SMS_DEFAULT_TASK_TEMPLATE;
  const replacements = new Map([
    ["[ID_задачи]", context.taskId || ""],
    ["[Название_задачи]", context.taskTitle || "задача"],
    ["[Объект]", context.objectName || ""],
    ["[Срок]", context.dueDate || ""],
    ["[ФИО]", context.fullName || "сотрудник"],
    ["[Бот]", context.botLabel || "бот"],
    ["[Ссылка_бота]", context.botLink || "ссылка будет предоставлена администратором"],
    ["[task_id]", context.taskId || ""],
    ["[task_title]", context.taskTitle || ""],
    ["[object]", context.objectName || ""],
    ["[due_date]", context.dueDate || ""],
    ["[full_name]", context.fullName || ""],
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
    inviteTemplate: String(ds.smsInviteTemplate || "").trim(),
    taskTemplate: String(ds.smsTaskTemplate || "").trim()
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

function buildSmsTaskMessage(payload, taskRow, employeeRow) {
  const displaySettings = payload?.displaySettings || {};
  const settings = getSmsGatewaySettings(displaySettings);
  const employeeId = String(employeeRow?.[0] || "").trim();
  const fullName = String(employeeRow?.[EMPLOYEE_FULL_NAME_COL] || "").trim() || "Сотрудник";
  const botUsername = String(displaySettings?.telegramBotUsername || "").trim().replace(/^@+/, "");
  const botLabel = botUsername ? `@${botUsername}` : "Telegram-бот";
  const botLink = buildSmsInviteBotLink(displaySettings, employeeId);
  const text = applySmsTaskTemplate(settings.taskTemplate, {
    taskId: String(taskRow?.[TASK_NUMBER_COL] || "").trim(),
    taskTitle: String(taskRow?.[TASK_TITLE_COL] || "").trim(),
    objectName: String(taskRow?.[1] || "").trim(),
    dueDate: String(taskRow?.[TASK_DUE_DATE_COL] || "").trim(),
    fullName,
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
  let providerMessage = truncateForLog(
    parsed?.message
      || parsed?.description
      || parsed?.status
      || parsed?.error
      || responseText
      || "",
    800
  );

  if (provider === SMS_GATE_PROVIDER_ID && parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
    const messageId = String(parsed?.id || "").trim();
    const state = String(parsed?.state || "").trim() || "Pending";
    const recipients = Array.isArray(parsed?.recipients) ? parsed.recipients.length : 0;
    providerMessage = truncateForLog(
      `Принято SMS Gate: state=${state}${messageId ? `, id=${messageId}` : ""}${recipients > 0 ? `, recipients=${recipients}` : ""}.`,
      300
    );
  }

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

function employeeLastActivityMs(value) {
  const raw = String(value || "").trim();
  if (!raw || raw === "—") return 0;
  const ms = Date.parse(raw);
  return Number.isFinite(ms) ? ms : 0;
}

function mergeEmployeeLastActivityFields(currentPayload, nextPayload) {
  const currentRows = Array.isArray(getEmployeesSection(currentPayload)?.rows)
    ? getEmployeesSection(currentPayload).rows
    : [];
  const nextRows = Array.isArray(getEmployeesSection(nextPayload)?.rows)
    ? getEmployeesSection(nextPayload).rows
    : [];
  if (!currentRows.length || !nextRows.length) return;
  const currentByKey = new Map();
  currentRows.forEach((row) => {
    const id = String(row?.[0] || "").trim();
    const phone = normalizePhone(row?.[EMPLOYEE_PHONE_COL] || "");
    const name = String(row?.[EMPLOYEE_FULL_NAME_COL] || "").trim().replace(/\s+/g, " ").toLowerCase();
    const keys = [id ? `id:${id}` : "", phone ? `phone:${phone}` : "", name ? `name:${name}` : ""].filter(Boolean);
    keys.forEach((key) => {
      if (!currentByKey.has(key)) currentByKey.set(key, row);
    });
  });
  nextRows.forEach((row) => {
    const id = String(row?.[0] || "").trim();
    const phone = normalizePhone(row?.[EMPLOYEE_PHONE_COL] || "");
    const name = String(row?.[EMPLOYEE_FULL_NAME_COL] || "").trim().replace(/\s+/g, " ").toLowerCase();
    const current = (id && currentByKey.get(`id:${id}`))
      || (phone && currentByKey.get(`phone:${phone}`))
      || (name && currentByKey.get(`name:${name}`));
    if (!current) return;
    const currentMs = employeeLastActivityMs(current[EMPLOYEE_LAST_ACTIVITY_COL]);
    const nextMs = employeeLastActivityMs(row[EMPLOYEE_LAST_ACTIVITY_COL]);
    if (currentMs > nextMs) {
      row[EMPLOYEE_LAST_ACTIVITY_COL] = current[EMPLOYEE_LAST_ACTIVITY_COL];
    }
  });
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
  mergeEmployeeLastActivityFields(currentPayload, next);
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

async function issueSingleSessionVersion(subject) {
  const key = String(subject || "").trim();
  if (!key) return 0;
  const { rows } = await pool.query(
    `INSERT INTO auth_sessions (subject, session_version, updated_at)
     VALUES ($1, 1, NOW())
     ON CONFLICT (subject)
     DO UPDATE SET session_version = auth_sessions.session_version + 1, updated_at = NOW()
     RETURNING session_version`,
    [key]
  );
  return Number(rows[0]?.session_version) || 1;
}

async function isSessionVersionCurrent(subject, sessionVersion) {
  const key = String(subject || "").trim();
  if (!key) return false;
  const { rows } = await pool.query(
    "SELECT session_version FROM auth_sessions WHERE subject = $1",
    [key]
  );
  if (!rows.length) return sessionVersion == null;
  const current = Number(rows[0]?.session_version) || 0;
  const tokenVersion = Number(sessionVersion);
  return Number.isFinite(tokenVersion) && tokenVersion === current;
}

async function authMiddleware(req, res, next) {
  const h = req.headers.authorization || "";
  const m = h.match(/^Bearer\s+(.+)$/i);
  if (!m) {
    return res.status(401).json({ error: "Требуется авторизация" });
  }
  try {
    const payload = jwt.verify(m[1], JWT_SECRET);
    const subject = payload?.sub != null ? String(payload.sub) : "";
    const hasCurrentSession = await isSessionVersionCurrent(subject, payload?.sv);
    if (!hasCurrentSession) {
      return res.status(401).json({ error: "Сессия завершена на другом устройстве" });
    }
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

async function resolveEffectiveAuthRole(req) {
  if (isAdminUser(req)) return "admin";
  try {
    const payload = await loadAppPayload();
    const employee = findEmployeeForAuthUserInPayload(payload, req.user);
    if (isEmployeeAdminAccessEnabled(payload, employee)) return "admin";
  } catch (e) {
    console.warn("Не удалось проверить админ-доступ сотрудника:", e?.message || e);
  }
  return "user";
}

async function requireAdmin(req, res, next) {
  const role = await resolveEffectiveAuthRole(req);
  if (role !== "admin") {
    return res.status(403).json({
      error: "Недостаточно прав",
      hint: "Токен действителен, но сервер не видит у пользователя права администратора. Проверьте галочку Админ у сотрудника и войдите заново."
    });
  }
  req.user.role = "admin";
  next();
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
      const subject = String(u.id);
      const authUser = {
        sub: subject,
        role: u.role,
        name: u.display_name,
        phone: u.phone || phone
      };
      let effectiveRole = String(u.role || "").trim().toLowerCase() === "admin" ? "admin" : "user";
      try {
        const { rows: stateRows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
        const payload = stateRows[0]?.payload;
        effectiveRole = resolveEffectiveRoleFromPayload(payload, authUser, effectiveRole);
      } catch (_) {
        // Если app_state временно недоступен — используем роль из users.
      }
      const sessionVersion = await issueSingleSessionVersion(subject);
      const token = jwt.sign(
        { sub: subject, role: effectiveRole, name: u.display_name, phone: u.phone || phone, sv: sessionVersion },
        JWT_SECRET,
        { expiresIn: "30d" }
      );
      await recordEmployeeLastLogin({ phone: u.phone || phone, displayName: u.display_name });
      return res.json({ token, displayName: u.display_name, role: effectiveRole });
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
        const subject = `emp:${phone}`;
        const sessionVersion = await issueSingleSessionVersion(subject);
        const token = jwt.sign(
          { sub: subject, role, name: displayName, phone, sv: sessionVersion },
          JWT_SECRET,
          { expiresIn: "30d" }
        );
        await recordEmployeeLastLogin({ phone, displayName });
        return res.json({ token, displayName, role });
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
      const sessionVersion = await issueSingleSessionVersion("admin");
      const token = jwt.sign({ sub: "admin", role: "admin", name: "Пользователь", phone: envPhone, sv: sessionVersion }, JWT_SECRET, { expiresIn: "30d" });
      return res.json({ token, displayName: "Пользователь", role: "admin" });
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сервера" });
  }

  return res.status(401).json({ error: "Неверный логин или пароль" });
});

app.get("/api/auth/me", authMiddleware, async (req, res) => {
  const name = String(req.user?.name || "").trim() || "Пользователь";
  const role = await resolveEffectiveAuthRole(req);
  const id = req.user?.sub != null ? String(req.user.sub) : "";
  return res.json({ id, displayName: name, role });
});

app.get("/api/docs", (_req, res) => {
  res.type("html").send(renderSwaggerDocsHtml());
});

app.get("/api/openapi.json", authMiddleware, requireAdmin, (req, res) => {
  res.json(buildOpenApiSpec(req));
});

app.get("/api/export/tasks", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const payload = await loadAppPayload();
    const section = buildExportSectionPayload(payload, "tasks");
    return res.json({ ok: true, section });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка экспорта задач" });
  }
});

app.get("/api/export/employees", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const payload = await loadAppPayload();
    const section = buildExportSectionPayload(payload, "employees");
    return res.json({ ok: true, section });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка экспорта сотрудников" });
  }
});

app.get("/api/export/objects", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const payload = await loadAppPayload();
    const section = buildExportSectionPayload(payload, "objects");
    return res.json({ ok: true, section });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка экспорта объектов" });
  }
});

app.get("/api/export/catalogs", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const payload = await loadAppPayload();
    return res.json({ ok: true, catalogs: buildCatalogsExportPayload(payload) });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка экспорта справочников" });
  }
});

app.get("/api/export/all", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    const payload = await loadAppPayload();
    return res.json({
      ok: true,
      sections: Array.isArray(payload?.sections) ? payload.sections : [],
      displaySettings: sanitizeDisplaySettingsForExport(payload?.displaySettings)
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка экспорта данных" });
  }
});

app.get("/api/admin/server/metrics", authMiddleware, requireAdmin, async (_req, res) => {
  try {
    return res.json(await buildServerMetricsPayload());
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка чтения метрик сервера" });
  }
});

app.get("/api/data", authMiddleware, async (_req, res) => {
  try {
    const { rows } = await pool.query("SELECT payload, revision FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload ?? null;
    const rev = Number(rows[0]?.revision) || 0;
    return res.json({ data: payload, rev });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка чтения данных" });
  }
});

app.post("/api/telegram/webhook", express.json(), async (req, res) => {
  await handleTelegramWebhook(req, res, pool);
});

/** После сохранения токена в приложении: зарегистрировать webhook на этом домене. */
app.post("/api/telegram/set-webhook", authMiddleware, requireAdmin, async (req, res) => {
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

app.post("/api/employees/refresh-chat-ids", authMiddleware, requireAdmin, async (_req, res) => {
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

app.post("/api/telegram/send-photo-proxy", authMiddleware, requireAdmin, async (req, res) => {
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

app.post("/api/telegram/send-media-group-proxy", authMiddleware, requireAdmin, async (req, res) => {
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

function computeNextReassignCodeInPayload(payload, taskId) {
  const task = String(taskId || "").trim();
  if (!task) return "1";
  let max = 0;
  const parseSeq = (code) => {
    const m = String(code || "").trim().match(/^(.+?)\/(\d+)$/);
    if (!m || m[1] !== task) return 0;
    const n = Number(m[2]);
    return Number.isFinite(n) && n > 0 ? n : 0;
  };
  const logArr = Array.isArray(payload?.taskReassignLog?.[task]) ? payload.taskReassignLog[task] : [];
  for (const item of logArr) {
    const n = parseSeq(item?.code);
    if (n > max) max = n;
  }
  const reqStore = payload?.telegramReassignRequests && typeof payload.telegramReassignRequests === "object"
    ? payload.telegramReassignRequests
    : {};
  for (const key of Object.keys(reqStore)) {
    const item = reqStore[key];
    if (String(item?.taskId || "").trim() !== task) continue;
    const n = parseSeq(item?.code);
    if (n > max) max = n;
  }
  return `${task}/${max + 1}`;
}

app.post("/api/tasks/reassign/request", authMiddleware, async (req, res) => {
  try {
    const taskId = String(req.body?.taskId || "").trim();
    const toEmployeeName = String(req.body?.toEmployeeName || "").trim();
    const reasonText = String(req.body?.reasonText || "").trim();
    const reasonTypeRaw = String(req.body?.reasonType || "mistake").trim().toLowerCase();
    // Принимаем как новые имена (mistake/delegation), так и старые (objective/subjective)
    // — последние мапим к новым: objective→mistake, subjective→delegation.
    const normalizedType = reasonTypeRaw === "objective"
      ? "mistake"
      : reasonTypeRaw === "subjective"
        ? "delegation"
        : reasonTypeRaw;
    const reasonType = ["mistake", "delegation"].includes(normalizedType) ? normalizedType : "mistake";
    const departmentNameInput = String(req.body?.departmentName || "").trim();

    if (!taskId) return res.status(400).json({ ok: false, error: "taskId обязателен" });
    if (!toEmployeeName) return res.status(400).json({ ok: false, error: "Не указан получатель" });
    // reasonText теперь опционален: оба новых типа (Ошибочная/Делегирование)
    // идут одним маршрутом без свободного текста причины.
    if (reasonText.length > 4000) return res.status(400).json({ ok: false, error: "Причина слишком длинная" });

    const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload && typeof rows[0].payload === "object"
      ? JSON.parse(JSON.stringify(rows[0].payload))
      : {};

    const tasks = getTaskRows(payload);
    const taskRow = tasks.find((r) => String(r?.[TASK_NUMBER_COL] || "").trim() === taskId);
    if (!taskRow) return res.status(404).json({ ok: false, error: "Задача не найдена" });

    const targetEmployee = findEmployeeByFullNameInPayload(payload, toEmployeeName);
    if (!targetEmployee) {
      return res.status(404).json({ ok: false, error: "Получатель не найден среди сотрудников" });
    }
    const targetDepartment = departmentNameInput
      || String(targetEmployee?.[2] || "").trim();

    const reqStore = ensureObjectStore(payload, "telegramReassignRequests");
    const duplicate = Object.values(reqStore).some((item) =>
      item && typeof item === "object"
      && String(item.taskId || "").trim() === taskId
      && String(item.status || "").trim() === "pending"
      && String(item.toEmployeeName || "").trim().toLowerCase() === toEmployeeName.toLowerCase()
    );
    if (duplicate) {
      return res.status(409).json({ ok: false, error: "По этой задаче уже есть pending-заявка на этого сотрудника." });
    }

    const requesterName = String(req.user?.name || "").trim() || "Пользователь";
    const currentAssignees = parseTaskAssigneeNames(taskRow[TASK_ASSIGNED_COL]);
    const fromName = currentAssignees.find((n) => n.toLowerCase() === requesterName.toLowerCase())
      || currentAssignees[0]
      || "";

    const requestId = `rr_${Date.now().toString(36)}${randomBytes(3).toString("hex")}`;
    const code = computeNextReassignCodeInPayload(payload, taskId);
    const nowIso = new Date().toISOString();

    reqStore[requestId] = {
      id: requestId,
      code,
      taskId,
      status: "pending",
      reasonType,
      reasonText,
      departmentName: targetDepartment,
      fromEmployeeName: fromName,
      toEmployeeName,
      requesterChatId: "",
      requesterName,
      requesterAssigneeName: fromName || requesterName,
      createdAt: nowIso,
      sourceMessageId: null,
      source: "web",
      allowedConfirmChatIds: []
    };

    appendTaskHistory(
      payload,
      taskId,
      requesterName,
      `Веб: запрошено переназначение (${fromName || "—"} → ${toEmployeeName})`
    );

    await saveAppPayload(payload);
    return res.json({ ok: true, requestId, code });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка создания заявки на переназначение" });
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
      // Колонка "Тип переназначения" — Ошибочная задача / Делегирование задачи.
      const _typeRaw = String(reqEntry.reasonType || "").trim().toLowerCase();
      const _typeLabel = (_typeRaw === "mistake" || _typeRaw === "objective")
        ? "Ошибочная задача"
        : (_typeRaw === "delegation" || _typeRaw === "subjective")
          ? "Делегирование задачи"
          : "";
      if (_typeLabel) row[TASK_REASSIGN_TYPE_COL] = _typeLabel;
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

/**
 * Подтверждение/отклонение запроса на закрытие задачи администратором из веба.
 * Полностью повторяет логику бот-handler-а ad|y / ad|n: на approve — статус → Закрыт,
 * на reject — статус возвращается на previousStatus (или В процессе), telegramCloseRequests
 * entry удаляется. Если есть подключённый бот — исполнитель получает уведомление в Telegram.
 */
app.post("/api/tasks/close/decision", authMiddleware, requireAdmin, async (req, res) => {
  try {
    const taskId = String(req.body?.taskId || "").trim();
    const decision = String(req.body?.decision || "").trim().toLowerCase();
    if (!taskId) return res.status(400).json({ ok: false, error: "taskId обязателен" });
    if (!["approve", "reject"].includes(decision)) {
      return res.status(400).json({ ok: false, error: "decision должен быть approve/reject" });
    }
    const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload && typeof rows[0].payload === "object"
      ? JSON.parse(JSON.stringify(rows[0].payload))
      : {};
    const tasks = getTaskRows(payload);
    const taskRow = tasks.find((r) => String(r?.[TASK_NUMBER_COL] || "").trim() === taskId);
    if (!taskRow) return res.status(404).json({ ok: false, error: "Задача не найдена" });
    const closeRequests = payload.telegramCloseRequests || {};
    const closeReq = closeRequests[taskId];
    if (!closeReq) return res.status(404).json({ ok: false, error: "Запрос на закрытие не найден" });

    const actorName = String(req.user?.name || "").trim() || "Администратор";
    const nowIso = new Date().toISOString();
    const token = String(payload?.displaySettings?.telegramBotToken || "").trim();
    const requesterChat = String(closeReq.chatId || "").trim();

    if (decision === "approve") {
      taskRow[TASK_STATUS_COL] = "Закрыт";
      const tz = String(payload?.displaySettings?.serverTimezone || "").trim() || "UTC";
      try {
        taskRow[TASK_CLOSED_DATE_COL] = new Intl.DateTimeFormat("ru-RU", {
          timeZone: tz,
          day: "2-digit",
          month: "2-digit",
          year: "numeric"
        }).format(new Date());
      } catch (_) {
        taskRow[TASK_CLOSED_DATE_COL] = new Date().toISOString().split("T")[0];
      }
      appendTaskHistory(payload, taskId, actorName, `Веб: подтверждено закрытие задачи (${actorName}, ${nowIso})`);
      delete payload.telegramCloseRequests[taskId];
      if (token) {
        if (requesterChat) {
          await tgSendMessage(token, requesterChat, `✅ Задача №${taskId}: закрытие подтверждено (${actorName}).`);
        }
        // Гасим висящие "Подтвердить/Отклонить" в админских чатах.
        try { await clearCloseConfirmPrompts(token, closeReq, taskId, `закрытие подтверждено (${actorName})`); } catch {}
      }
    } else {
      const prevStatus = String(closeReq.previousStatus || "В процессе").trim() || "В процессе";
      if (String(taskRow[TASK_STATUS_COL] || "").trim() === "Проверка") {
        taskRow[TASK_STATUS_COL] = prevStatus;
      }
      appendTaskHistory(payload, taskId, actorName, `Веб: отклонён запрос на закрытие, статус → «${prevStatus}»`);
      delete payload.telegramCloseRequests[taskId];
      if (token) {
        if (requesterChat) {
          await tgSendMessage(token, requesterChat, `❌ Задача №${taskId}: запрос на закрытие отклонён администратором (${actorName}). Статус возвращён в «${prevStatus}».`);
        }
        try { await clearCloseConfirmPrompts(token, closeReq, taskId, `закрытие отклонено (${actorName})`); } catch {}
      }
    }

    // ВАЖНО: не используем saveAppPayload — она перетирает секцию задач из БД
    // и сбрасывает наши изменения status/closedDate. Пишем payload напрямую с
    // бампом revision (чтобы веб-клиенты с устаревшим baseRev получили 409 и
    // прошли через 3-way merge) и шлём broadcast по WebSocket.
    const { rows: savedRows } = await pool.query(
      `INSERT INTO app_state (id, payload, updated_at, revision) VALUES (1, $1::jsonb, NOW(), 1)
       ON CONFLICT (id) DO UPDATE SET
         payload = EXCLUDED.payload,
         updated_at = NOW(),
         revision = app_state.revision + 1
       RETURNING revision`,
      [JSON.stringify(payload)]
    );
    const newRev = Number(savedRows[0]?.revision) || 0;
    try { broadcastStateChanged(newRev, { source: "close-decision" }); } catch {}
    return res.json({ ok: true });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка обработки решения по закрытию" });
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

    if (!phone) {
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

app.post("/api/sms/task/send", authMiddleware, requireAdmin, async (req, res) => {
  const nowMs = Date.now();
  try {
    const taskId = String(req.body?.taskId || "").trim();
    if (!taskId) {
      return res.status(400).json({ ok: false, error: "taskId обязателен." });
    }

    const payload = await loadAppPayload();
    const smsSettings = getSmsGatewaySettings(payload?.displaySettings || {});
    if (!smsSettings.enabled) {
      return res.status(400).json({ ok: false, error: "SMS Gateway выключен в настройках." });
    }
    if (!smsSettings.url) {
      return res.status(400).json({ ok: false, error: "Не указан URL SMS Gateway в настройках." });
    }

    const tasksRows = getTaskRows(payload);
    const taskRow = tasksRows.find((row) => String(row?.[TASK_NUMBER_COL] || "").trim() === taskId);
    if (!taskRow) {
      return res.status(404).json({ ok: false, error: "Задача не найдена." });
    }
    const taskTitle = String(taskRow?.[TASK_TITLE_COL] || "").trim() || `Задача ${taskId}`;
    const actorName = String(req.user?.name || "").trim() || "Администратор";

    const recipientNames = collectTaskTelegramRecipientNames(taskRow);
    if (!recipientNames.length) {
      return res.status(400).json({ ok: false, error: "В задаче не указан исполнитель/ответственный для SMS." });
    }

    const recipientsByPhone = new Map();
    for (const name of recipientNames) {
      const employeeRow = findEmployeeByFullNameInPayload(payload, name);
      if (!employeeRow) continue;
      const phone = normalizePhone(employeeRow?.[4] || "");
      if (!phone) continue;
      if (!recipientsByPhone.has(phone)) {
        recipientsByPhone.set(phone, employeeRow);
      }
    }

    if (!recipientsByPhone.size) {
      return res.status(400).json({ ok: false, error: "Не найдены сотрудники с корректным телефоном для отправки SMS." });
    }

    const logStore = ensureSmsInviteLogArray(payload);
    let sentCount = 0;
    let failCount = 0;
    const errors = [];
    const entries = [];

    for (const [phone, employeeRow] of recipientsByPhone.entries()) {
      const employeeId = String(employeeRow?.[0] || "").trim();
      const employeeName = String(employeeRow?.[EMPLOYEE_FULL_NAME_COL] || "").trim() || "Сотрудник";
      const chatId = String(employeeRow?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
      const taskMessage = buildSmsTaskMessage(payload, taskRow, employeeRow);
      const logEntry = {
        id: randomBytes(10).toString("hex"),
        atMs: nowMs,
        atIso: new Date(nowMs).toISOString(),
        employeeId,
        employeeName,
        phone,
        chatId,
        text: truncateForLog(taskMessage.text, 2000),
        actor: actorName,
        status: "failed",
        ok: false,
        gatewayProvider: smsSettings.provider,
        gatewayMethod: smsSettings.method,
        gatewayUrl: truncateForLog(smsSettings.url, 240),
        httpStatus: 0,
        resultMessage: "",
        responsePreview: "",
        smsKind: "task",
        taskId,
        taskTitle: truncateForLog(taskTitle, 240)
      };

      if (!taskMessage.text) {
        logEntry.resultMessage = "Текст SMS по задаче пустой. Проверьте шаблон.";
      } else {
        const sendResult = await sendSmsViaGateway(smsSettings, { phone, text: taskMessage.text });
        logEntry.ok = sendResult.ok === true;
        logEntry.status = sendResult.ok ? "sent" : "failed";
        logEntry.httpStatus = Number(sendResult.httpStatus) || 0;
        logEntry.resultMessage = truncateForLog(
          sendResult.providerMessage || sendResult.reason || (sendResult.ok ? "SMS отправлено." : "Неизвестная ошибка."),
          800
        );
        logEntry.responsePreview = truncateForLog(sendResult.responsePreview || "", 1200);
      }

      if (logEntry.ok) {
        sentCount += 1;
      } else {
        failCount += 1;
        errors.push(`${employeeName}: ${String(logEntry.resultMessage || "ошибка")}`);
      }
      entries.push(logEntry);
      logStore.unshift(logEntry);
    }

    if (logStore.length > SMS_INVITE_LOG_LIMIT) {
      logStore.length = SMS_INVITE_LOG_LIMIT;
    }
    await saveAppPayload(payload);

    if (sentCount > 0) {
      return res.json({
        ok: true,
        taskId,
        taskTitle,
        totalRecipients: recipientsByPhone.size,
        sentCount,
        failCount,
        errors,
        entries
      });
    }
    return res.status(400).json({
      ok: false,
      error: errors[0] || "Не удалось отправить SMS по задаче.",
      taskId,
      taskTitle,
      totalRecipients: recipientsByPhone.size,
      sentCount,
      failCount,
      errors,
      entries
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "Ошибка отправки SMS по задаче." });
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

app.post("/api/google-sheets/sync", authMiddleware, requireAdmin, async (_req, res) => {
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
    await ensureMediaStorageDir(MEDIA_STORAGE_PATH);
    const dataUrl = String(req.body?.dataUrl || "").trim();
    const sourceFileName = String(req.body?.fileName || "").trim();
    const upload = validateMediaUpload({ dataUrl, fileName: sourceFileName });
    if (!upload.ok) {
      return res.status(upload.status || 400).json({ error: upload.error || "Некорректный файл." });
    }
    const fileName = `${Date.now()}-${randomBytes(6).toString("hex")}.${upload.ext}`;
    const absPath = path.join(MEDIA_STORAGE_PATH, fileName);
    await fsp.writeFile(absPath, upload.buf);
    const base = getPublicBaseUrl(req);
    const url = `${base}/media/${encodeURIComponent(fileName)}`;
    return res.json({
      ok: true,
      url,
      fileName,
      mime: upload.mime
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
    const { rows: currentRows } = await pool.query("SELECT payload, revision FROM app_state WHERE id = 1");
    const currentPayload = currentRows[0]?.payload && typeof currentRows[0].payload === "object"
      ? currentRows[0].payload
      : {};
    const currentRev = Number(currentRows[0]?.revision) || 0;
    if (baseRev > 0 && currentRev > 0 && baseRev < currentRev) {
      return res.status(409).json({
        error: "stale_data",
        message: "Данные устарели. Сервер вернул свежий снимок — наложите свои правки и повторите.",
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
         updated_at = NOW(),
         revision = app_state.revision + 1
       WHERE app_state.id = 1
         AND ($2::BIGINT = 0 OR app_state.revision = $2::BIGINT)
       RETURNING revision`,
      [JSON.stringify(mergedData), baseRev]
    );
    if (!savedRows.length) {
      const { rows: latestRows } = await pool.query("SELECT payload, revision FROM app_state WHERE id = 1");
      const latestPayload = latestRows[0]?.payload && typeof latestRows[0].payload === "object"
        ? latestRows[0].payload
        : {};
      const latestRev = Number(latestRows[0]?.revision) || 0;
      return res.status(409).json({
        error: "stale_data",
        message: "Данные изменились между чтением и записью — повторите с актуальной базой.",
        data: latestPayload,
        rev: latestRev
      });
    }
    const savedRev = Number(savedRows[0]?.revision) || 0;
    const broadcasterId = req.user?.sub != null ? String(req.user.sub) : "";
    try { broadcastStateChanged(savedRev, { source: "put", by: broadcasterId }); } catch {}
    return res.json({ ok: true, rev: savedRev });
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

app.use(express.static(staticRoot, { index: false }));

app.get("*", (_req, res) => {
  res.sendFile(path.join(staticRoot, "index.html"));
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
  await ensureMediaStorageDir(MEDIA_STORAGE_PATH);
  console.log(`Папка медиа: ${MEDIA_STORAGE_PATH}`);

  const httpServer = createHttpServer(app);
  attachRealtimeHub({
    httpServer,
    jwtSecret: JWT_SECRET,
    isSessionVersionCurrent,
    getCurrentRevision: async () => {
      try {
        const { rows } = await pool.query("SELECT revision FROM app_state WHERE id = 1");
        return Number(rows[0]?.revision) || 0;
      } catch {
        return 0;
      }
    }
  });
  // Exposing broadcast so server/telegramWebhook.js#savePayload может позвать его
  // без циклического импорта: бот пишет в БД → бампит revision → broadcast →
  // web-клиенты пуллят свежий снимок до того как успеют запушить старый.
  globalThis.__broadcastStateChanged = broadcastStateChanged;
  httpServer.listen(PORT, "0.0.0.0", () => {
    console.log(`Сервер слушает порт ${PORT} (${NODE_ENV}) — WebSocket на /ws/state`);
  });
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

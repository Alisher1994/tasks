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
const TASK_PLAN_COL = 12;
const TASK_CLOSED_DATE_COL = 15;
const TASK_MEDIA_AFTER_COL = 17;
const TASK_READ_STATE_COL = 18;
const TASK_LAST_SENT_AT_COL = 19;
const TASK_DELAY_REASON_COL = 20;

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

function findEmployeeByPhoneInPayload(payload, phone) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  for (const row of rows) {
    const rowPhone = normalizePhone(row?.[4] || "");
    if (rowPhone === phone) return row;
  }
  return null;
}

function getTaskRows(payload) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  const tasks = sections.find((s) => s && s.id === "tasks");
  return Array.isArray(tasks?.rows) ? tasks.rows : [];
}

function isReadStateValue(value) {
  const firstLine = String(value || "").split(/\r?\n/)[0].trim().toLowerCase();
  return firstLine.startsWith("прочитано");
}

function hasLastSentValue(value) {
  const s = String(value || "").trim();
  return Boolean(s && s !== "—");
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

function mergeTaskSyncSafeFields(currentPayload, incomingPayload) {
  const next = incomingPayload && typeof incomingPayload === "object"
    ? JSON.parse(JSON.stringify(incomingPayload))
    : incomingPayload;
  const incomingRows = getTaskRows(next);
  const currentRows = getTaskRows(currentPayload);
  if (!incomingRows.length || !currentRows.length) return next;
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
      "SELECT id, password_hash, display_name, role FROM users WHERE phone = $1",
      [phone]
    );
    if (rows.length > 0) {
      const u = rows[0];
      const ok = await bcrypt.compare(password, u.password_hash);
      if (!ok) {
        return res.status(401).json({ error: "Неверный логин или пароль" });
      }
      const token = jwt.sign(
        { sub: String(u.id), role: u.role, name: u.display_name },
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
        const token = jwt.sign(
          { sub: `emp:${phone}`, role: "user", name: displayName },
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
    const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const payload = rows[0]?.payload ?? null;
    return res.json({ data: payload });
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
    const { rows: currentRows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
    const currentPayload = currentRows[0]?.payload && typeof currentRows[0].payload === "object"
      ? currentRows[0].payload
      : {};
    const mergedData = mergeTaskSyncSafeFields(currentPayload, data);
    // Служебные поля Telegram живут только на сервере. На клиентском PUT /api/data
    // мы всегда сохраняем их из текущего payload в БД (без доверия клиентскому слепку),
    // чтобы гонки между webhook и авто-sync не ломали сценарии комментариев/фото.
    await pool.query(
      `INSERT INTO app_state (id, payload, updated_at)
       VALUES (1, $1::jsonb, NOW())
       ON CONFLICT (id) DO UPDATE SET
         payload = jsonb_set(
           jsonb_set(
             jsonb_set(
               EXCLUDED.payload,
               '{telegramSessions}',
               COALESCE(app_state.payload->'telegramSessions', '{}'::jsonb),
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
         updated_at = NOW()`,
      [JSON.stringify(mergedData)]
    );
    return res.json({ ok: true });
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

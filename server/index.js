/**
 * API + раздача статики для Railway.
 * Переменные: DATABASE_URL, JWT_SECRET, ADMIN_PHONE, ADMIN_PASSWORD (первый админ), PORT
 */
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import jwt from "jsonwebtoken";
import pg from "pg";
import bcrypt from "bcryptjs";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import { randomBytes } from "crypto";
import { configureTelegramWebhook, handleTelegramWebhook } from "./telegramWebhook.js";
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

function normalizePhone(raw) {
  const digits = String(raw || "").replace(/\D/g, "");
  const c = "998";
  let local = digits;
  if (local.startsWith(c)) local = local.slice(3);
  else if (local.startsWith("8")) local = local.slice(1);
  local = local.slice(0, 9);
  return `+998${local}`;
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
    return res.json({ ok: true, webhookUrl: result.webhookUrl, botUsername: result.botUsername || "" });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Ошибка сервера" });
  }
});

app.put("/api/data", authMiddleware, async (req, res) => {
  try {
    const data = req.body?.data;
    if (data === undefined) {
      return res.status(400).json({ error: "Поле data обязательно" });
    }
    const v = validateAppPayload(data);
    if (!v.ok) {
      return res.status(400).json({ error: v.error || "Некорректные данные" });
    }
    await pool.query(
      `INSERT INTO app_state (id, payload, updated_at)
       VALUES (1, $1::jsonb, NOW())
       ON CONFLICT (id) DO UPDATE SET payload = EXCLUDED.payload, updated_at = NOW()`,
      [JSON.stringify(data)]
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
  console.log("Миграции и сид пользователей выполнены.");

  if (JWT_SECRET === "change-me-in-production" && NODE_ENV === "production") {
    console.error("Задайте JWT_SECRET в production.");
    process.exit(1);
  }
  if (JWT_SECRET === "change-me-in-production") {
    console.warn("Задайте JWT_SECRET в переменных окружения перед production.");
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Сервер слушает порт ${PORT} (${NODE_ENV})`);
  });
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

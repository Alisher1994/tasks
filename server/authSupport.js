import jwt from "jsonwebtoken";

export function createAuthSupport({
  pool,
  jwtSecret,
  loadAppPayload,
  findEmployeeForAuthUserInPayload,
  isEmployeeAdminAccessEnabled
}) {
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
      const payload = jwt.verify(m[1], jwtSecret);
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

  return {
    issueSingleSessionVersion,
    isSessionVersionCurrent,
    authMiddleware,
    resolveEffectiveAuthRole,
    requireAdmin
  };
}

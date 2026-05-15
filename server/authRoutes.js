import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";

export function registerAuthRoutes(app, {
  pool,
  loginLimiter,
  authMiddleware,
  jwtSecret,
  adminPhone,
  adminPassword,
  normalizePhone,
  last4DigitsPasswordFromPhone,
  findEmployeeByPhoneInPayload,
  isEmployeeAdminAccessEnabled,
  resolveEffectiveRoleFromPayload,
  issueSingleSessionVersion,
  recordEmployeeLastLogin,
  resolveEffectiveAuthRole
}) {
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
          jwtSecret,
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
            jwtSecret,
            { expiresIn: "30d" }
          );
          await recordEmployeeLastLogin({ phone, displayName });
          return res.json({ token, displayName, role });
        }
      }

      const { rows: cntRows } = await pool.query("SELECT COUNT(*)::int AS c FROM users");
      const userCount = cntRows[0]?.c ?? 0;
      const envPhone = normalizePhone(adminPhone);
      if (
        userCount === 0 &&
        envPhone &&
        adminPassword &&
        phone === envPhone &&
        password === adminPassword
      ) {
        const sessionVersion = await issueSingleSessionVersion("admin");
        const token = jwt.sign(
          { sub: "admin", role: "admin", name: "Пользователь", phone: envPhone, sv: sessionVersion },
          jwtSecret,
          { expiresIn: "30d" }
        );
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
}

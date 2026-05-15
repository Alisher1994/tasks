import bcrypt from "bcryptjs";

export function registerAdminUserRoutes(app, { pool, authMiddleware, requireAdmin, normalizePhone }) {
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
}

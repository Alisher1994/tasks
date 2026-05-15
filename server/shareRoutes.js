import { randomBytes } from "crypto";

export function registerShareRoutes(app, { pool, authMiddleware }) {
  app.post("/api/share", authMiddleware, async (req, res) => {
    try {
      const pin = String(req.body?.pin || "");
      const expiresAt = Number(req.body?.expiresAt);
      const rows = req.body?.rows;
      if (!/^\d{4}$/.test(pin) || !Number.isFinite(expiresAt) || !Array.isArray(rows)) {
        return res.status(400).json({ error: "Некорректные данные" });
      }
      const id = randomBytes(16).toString("hex");
      await pool.query("INSERT INTO report_shares (id, pin, expires_at, rows) VALUES ($1, $2, $3, $4::jsonb)", [
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
}

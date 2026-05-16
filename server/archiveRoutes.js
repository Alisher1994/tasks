/**
 * Роуты архива закрытых задач.
 *
 * Архив хранится отдельно от app_state.payload, поэтому закрытые задачи
 * больше не возят при каждой синхронизации. Данные не удаляются:
 *  - POST /api/archive            — переложить пачку задач в архив (upsert)
 *  - GET  /api/archive            — список архива (пагинация + поиск)
 *  - GET  /api/archive/:taskId    — полная карточка из архива
 *  - POST /api/archive/:taskId/restore — вернуть задачу из архива (отдаёт data и удаляет запись)
 *
 * Все ручки за authMiddleware — как /api/data (любой авторизованный пользователь).
 */

const MAX_LIMIT = 200;
const DEFAULT_LIMIT = 50;
const MAX_BATCH = 500;

function clampInt(value, fallback, min, max) {
  const n = Number.parseInt(String(value), 10);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

export function registerArchiveRoutes(app, { pool, authMiddleware }) {
  // Список архива: пагинация + поиск по task_id / объекту / названию.
  app.get("/api/archive", authMiddleware, async (req, res) => {
    try {
      const limit = clampInt(req.query?.limit, DEFAULT_LIMIT, 1, MAX_LIMIT);
      const offset = clampInt(req.query?.offset, 0, 0, Number.MAX_SAFE_INTEGER);
      const q = String(req.query?.q || "").trim();

      const where = q ? "WHERE task_id ILIKE $1 OR object ILIKE $1 OR title ILIKE $1" : "";
      const params = q ? [`%${q}%`] : [];

      const countSql = `SELECT COUNT(*)::int AS n FROM task_archive ${where}`;
      const { rows: countRows } = await pool.query(countSql, params);
      const total = countRows[0]?.n || 0;

      const listSql = `
        SELECT task_id, object, status, closed_date, title, archived_at
        FROM task_archive
        ${where}
        ORDER BY archived_at DESC
        LIMIT $${params.length + 1} OFFSET $${params.length + 2}
      `;
      const { rows } = await pool.query(listSql, [...params, limit, offset]);

      return res.json({
        items: rows.map((r) => ({
          taskId: r.task_id,
          object: r.object,
          status: r.status,
          closedDate: r.closed_date,
          title: r.title,
          archivedAt: r.archived_at
        })),
        total,
        limit,
        offset
      });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ error: "Не удалось загрузить архив" });
    }
  });

  // Полная карточка из архива.
  app.get("/api/archive/:taskId", authMiddleware, async (req, res) => {
    try {
      const taskId = String(req.params.taskId || "").trim();
      if (!taskId) return res.status(400).json({ error: "Не указан taskId" });
      const { rows } = await pool.query(
        "SELECT data, archived_at FROM task_archive WHERE task_id = $1",
        [taskId]
      );
      if (!rows.length) return res.status(404).json({ error: "not_found" });
      return res.json({ data: rows[0].data, archivedAt: rows[0].archived_at });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ error: "Не удалось открыть карточку из архива" });
    }
  });

  // Переложить пачку задач в архив (idempotent upsert по task_id).
  app.post("/api/archive", authMiddleware, async (req, res) => {
    const tasks = Array.isArray(req.body?.tasks) ? req.body.tasks : null;
    if (!tasks) {
      return res.status(400).json({ error: "Ожидался массив tasks" });
    }
    if (tasks.length === 0) {
      return res.json({ archived: [] });
    }
    if (tasks.length > MAX_BATCH) {
      return res.status(400).json({ error: `Слишком большая пачка (макс. ${MAX_BATCH})` });
    }

    const client = await pool.connect();
    try {
      await client.query("BEGIN");
      const archived = [];
      for (const t of tasks) {
        const taskId = String(t?.taskId || "").trim();
        if (!taskId || t?.data === undefined || t?.data === null) continue;
        await client.query(
          `INSERT INTO task_archive (task_id, object, status, closed_date, title, data, archived_at)
           VALUES ($1, $2, $3, $4, $5, $6::jsonb, NOW())
           ON CONFLICT (task_id)
           DO UPDATE SET object = EXCLUDED.object,
                         status = EXCLUDED.status,
                         closed_date = EXCLUDED.closed_date,
                         title = EXCLUDED.title,
                         data = EXCLUDED.data,
                         archived_at = NOW()`,
          [
            taskId,
            String(t?.object || ""),
            String(t?.status || ""),
            String(t?.closedDate || ""),
            String(t?.title || ""),
            JSON.stringify(t.data)
          ]
        );
        archived.push(taskId);
      }
      await client.query("COMMIT");
      return res.json({ archived });
    } catch (e) {
      await client.query("ROLLBACK");
      console.error(e);
      return res.status(500).json({ error: "Не удалось заархивировать задачи" });
    } finally {
      client.release();
    }
  });

  // Вернуть задачу из архива: отдаём её данные и удаляем запись из архива.
  // Клиент после успешного ответа вставляет задачу обратно в живое состояние.
  app.post("/api/archive/:taskId/restore", authMiddleware, async (req, res) => {
    const taskId = String(req.params.taskId || "").trim();
    if (!taskId) return res.status(400).json({ error: "Не указан taskId" });
    const client = await pool.connect();
    try {
      await client.query("BEGIN");
      const { rows } = await client.query(
        "SELECT data FROM task_archive WHERE task_id = $1 FOR UPDATE",
        [taskId]
      );
      if (!rows.length) {
        await client.query("ROLLBACK");
        return res.status(404).json({ error: "not_found" });
      }
      const data = rows[0].data;
      await client.query("DELETE FROM task_archive WHERE task_id = $1", [taskId]);
      await client.query("COMMIT");
      return res.json({ data });
    } catch (e) {
      await client.query("ROLLBACK");
      console.error(e);
      return res.status(500).json({ error: "Не удалось вернуть задачу из архива" });
    } finally {
      client.release();
    }
  });
}

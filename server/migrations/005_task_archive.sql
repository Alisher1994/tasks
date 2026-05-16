-- Архив закрытых задач.
-- Отдельная таблица, чтобы закрытые задачи НЕ возили в едином app_state.payload
-- при каждой синхронизации (главный драйвер роста payload).
-- Данные не удаляются, а перекладываются сюда; возврат — через /api/archive/:id/restore.
--
-- Денормализованные колонки (object/status/closed_date/title) нужны для
-- пагинации и поиска в списке архива без разбора JSONB.
-- В data лежит самодостаточный снимок задачи:
--   { row, columns, history, attachments, closeMeta, multiState }
-- чтобы карточку из архива можно было открыть, не имея её в живом состоянии.
CREATE TABLE IF NOT EXISTS task_archive (
  task_id     TEXT PRIMARY KEY,
  archived_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  object      TEXT NOT NULL DEFAULT '',
  status      TEXT NOT NULL DEFAULT '',
  closed_date TEXT NOT NULL DEFAULT '',
  title       TEXT NOT NULL DEFAULT '',
  data        JSONB NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_task_archive_archived_at ON task_archive (archived_at DESC);
CREATE INDEX IF NOT EXISTS idx_task_archive_object ON task_archive (object);

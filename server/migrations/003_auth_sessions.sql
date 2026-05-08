-- Одна активная веб-сессия на пользователя/сотрудника.
CREATE TABLE IF NOT EXISTS auth_sessions (
  subject TEXT PRIMARY KEY,
  session_version INTEGER NOT NULL DEFAULT 0,
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

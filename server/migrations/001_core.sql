-- Основные таблицы приложения (единый JSON + шаринг отчётов)
CREATE TABLE IF NOT EXISTS app_state (
  id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS report_shares (
  id TEXT PRIMARY KEY,
  pin TEXT NOT NULL,
  expires_at BIGINT NOT NULL,
  rows JSONB NOT NULL
);

-- Монотонный счётчик ревизий app_state для optimistic locking.
-- Целочисленное сравнение точнее, чем Date.parse(updated_at), и
-- избавляет от теоретических race conditions в пределах одной миллисекунды.
ALTER TABLE app_state
  ADD COLUMN IF NOT EXISTS revision BIGINT NOT NULL DEFAULT 0;

-- Если строка уже существует (id=1) и revision=0, оставляем 0 —
-- клиенты увидят rev=0 при первом pull, и последующие PUT начнут с baseRev=0
-- (на этом этапе любой PUT успешен, дальше счётчик растёт монотонно).

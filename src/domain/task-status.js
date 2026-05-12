export const STATUS_DECISION_OLD = "Треб. реш. рук.";
export const STATUS_DECISION = "Требует решение руководителя";
// "Проверка" — задача переходит в этот статус после того, как исполнитель
// запросил закрытие. Админ должен подтвердить → "Закрыт", или отклонить →
// возврат на "В процессе". В клавиатуре боту/вебе исполнитель НЕ выбирает
// "Проверка" напрямую — система выставляет её сама.
export const STATUS_OPTIONS = ["Новый", "В процессе", "Проверка", "Закрыт"];

export function normalizeTaskStatusValue(raw) {
  const value = String(raw || "").trim();
  if (value === STATUS_DECISION_OLD || value === STATUS_DECISION) return "В процессе";
  return value;
}

export function normalizeTaskPriorityValue(raw) {
  const value = String(raw || "").trim();
  if (value === "Низкий") return "Средний";
  return value;
}

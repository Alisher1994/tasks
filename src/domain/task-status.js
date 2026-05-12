export const STATUS_DECISION_OLD = "Треб. реш. рук.";
export const STATUS_DECISION = "Требует решение руководителя";
export const STATUS_OPTIONS = ["Новый", "В процессе", "Закрыт"];

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

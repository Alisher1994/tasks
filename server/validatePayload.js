/**
 * Проверка структуры тела PUT /api/data перед записью в JSONB.
 */
const MAX_JSON_CHARS = 48 * 1024 * 1024;
const MAX_SECTIONS = 500;
const MAX_ROWS_PER_SECTION = 50000;

export function validateAppPayload(data) {
  if (data === null || typeof data !== "object" || Array.isArray(data)) {
    return { ok: false, error: "data должен быть объектом" };
  }

  let serialized;
  try {
    serialized = JSON.stringify(data);
  } catch {
    return { ok: false, error: "Данные не сериализуются в JSON" };
  }

  if (serialized.length > MAX_JSON_CHARS) {
    return { ok: false, error: "Слишком большой объём данных" };
  }

  if (data.sections !== undefined) {
    if (!Array.isArray(data.sections)) {
      return { ok: false, error: "sections должен быть массивом" };
    }
    if (data.sections.length > MAX_SECTIONS) {
      return { ok: false, error: "Слишком много секций" };
    }
    for (const sec of data.sections) {
      if (!sec || typeof sec !== "object" || Array.isArray(sec)) {
        return { ok: false, error: "Некорректный элемент sections" };
      }
      if (typeof sec.id !== "string" || !sec.id.trim()) {
        return { ok: false, error: "У секции должен быть непустой id (строка)" };
      }
      if (typeof sec.title !== "string") {
        return { ok: false, error: `Секция ${sec.id}: title` };
      }
      if (!Array.isArray(sec.columns)) {
        return { ok: false, error: `Секция ${sec.id}: columns должен быть массивом` };
      }
      if (!Array.isArray(sec.rows)) {
        return { ok: false, error: `Секция ${sec.id}: rows должен быть массивом` };
      }
      if (sec.rows.length > MAX_ROWS_PER_SECTION) {
        return { ok: false, error: `Секция ${sec.id}: слишком много строк` };
      }
    }
  }

  const optionalObjects = [
    "displaySettings",
    "trashBySection",
    "taskHistory",
    "taskMultiState",
    "telegramReassignRequests",
    "taskReassignLog",
    "taskAttachments",
    "reportShares"
  ];
  for (const key of optionalObjects) {
    if (data[key] !== undefined && (data[key] === null || typeof data[key] !== "object" || Array.isArray(data[key]))) {
      return { ok: false, error: `Поле ${key} должно быть объектом` };
    }
  }

  if (data.reportChartOrder !== undefined && !Array.isArray(data.reportChartOrder)) {
    return { ok: false, error: "reportChartOrder должен быть массивом" };
  }

  if (data.reportPhaseLayout !== undefined && data.reportPhaseLayout !== null) {
    const t = typeof data.reportPhaseLayout;
    if (t !== "string" && t !== "object") {
      return { ok: false, error: "reportPhaseLayout: строка или объект" };
    }
  }

  return { ok: true };
}

/**
 * 3-way merge для разрешения конфликтов optimistic locking.
 *
 * Когда сервер отвечает 409 на наш PUT /api/data, он возвращает свежий снимок
 * payload. Чтобы не потерять ни наши локальные правки, ни чужие, делаем 3-way
 * merge: baseline (что мы видели при последнем pull) vs local (наше состояние)
 * vs server (актуальное состояние сервера).
 *
 * Правила для каждой записи (поля, строки, ключа словаря):
 *  - local изменилось, server не изменилось → берём local
 *  - local не изменилось, server изменилось → берём server
 *  - оба изменились → конфликт, выигрывает local (наша правка свежее в UI), но
 *    счётчик `conflicts` увеличивается — UI покажет тост
 *  - ничего не изменилось → берём server (на всякий случай — это canonical state)
 *
 * Сравнение через JSON.stringify — достаточно для наших row-of-strings структур.
 */

function jsonEq(a, b) {
  if (a === b) return true;
  if (a === undefined && b === undefined) return true;
  return JSON.stringify(a) === JSON.stringify(b);
}

function clone(value) {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value));
}

export function cloneSnapshot(payload) {
  if (!payload || typeof payload !== "object") return {};
  return clone(payload);
}

/**
 * Считаем "ID" строки секции — первая колонка. Для секции tasks это task_id.
 * Если первая колонка пустая, fallback на индекс — но для корректного merge
 * лучше, чтобы у всех строк был непустой id.
 */
function rowKey(row, fallbackIndex) {
  if (!Array.isArray(row)) return `__idx_${fallbackIndex}`;
  const id = String(row[0] ?? "").trim();
  return id || `__idx_${fallbackIndex}`;
}

function mergeRows(baselineRows, localRows, serverRows) {
  const baseline = Array.isArray(baselineRows) ? baselineRows : [];
  const local = Array.isArray(localRows) ? localRows : [];
  const server = Array.isArray(serverRows) ? serverRows : [];

  const baselineByKey = new Map();
  baseline.forEach((row, idx) => baselineByKey.set(rowKey(row, idx), row));
  const localByKey = new Map();
  local.forEach((row, idx) => localByKey.set(rowKey(row, idx), row));
  const serverByKey = new Map();
  server.forEach((row, idx) => serverByKey.set(rowKey(row, idx), row));

  // Порядок строк определяется local — пользователь видит свой порядок;
  // в конец добавляем строки, появившиеся только на сервере.
  const orderedKeys = [];
  const seen = new Set();
  local.forEach((row, idx) => {
    const key = rowKey(row, idx);
    if (!seen.has(key)) {
      orderedKeys.push(key);
      seen.add(key);
    }
  });
  server.forEach((row, idx) => {
    const key = rowKey(row, idx);
    if (!seen.has(key)) {
      orderedKeys.push(key);
      seen.add(key);
    }
  });

  let conflicts = 0;
  const merged = [];
  for (const key of orderedKeys) {
    const baseRow = baselineByKey.get(key);
    const localRow = localByKey.get(key);
    const serverRow = serverByKey.get(key);

    // Удаление: было в baseline, удалено локально → удалить, даже если есть на сервере
    if (baseRow !== undefined && localRow === undefined) {
      continue;
    }
    // Удалено на сервере, не было локальных правок → удалить
    if (baseRow !== undefined && serverRow === undefined && jsonEq(localRow, baseRow)) {
      continue;
    }

    if (localRow === undefined && serverRow !== undefined) {
      // Новая строка с сервера
      merged.push(clone(serverRow));
      continue;
    }
    if (localRow !== undefined && serverRow === undefined) {
      // Новая локальная строка (либо удалена на сервере, но нами правлена — сохраняем)
      merged.push(clone(localRow));
      continue;
    }

    const localChanged = !jsonEq(localRow, baseRow);
    const serverChanged = !jsonEq(serverRow, baseRow);

    if (localChanged && !serverChanged) {
      merged.push(clone(localRow));
    } else if (!localChanged && serverChanged) {
      merged.push(clone(serverRow));
    } else if (localChanged && serverChanged) {
      // Конфликт на одной и той же строке: ячейка-в-ячейку merge.
      const mergedRow = clone(localRow);
      if (Array.isArray(baseRow) && Array.isArray(serverRow) && Array.isArray(localRow)) {
        const maxLen = Math.max(baseRow.length, serverRow.length, localRow.length);
        for (let i = 0; i < maxLen; i += 1) {
          const b = baseRow[i];
          const l = localRow[i];
          const s = serverRow[i];
          const lChanged = !jsonEq(l, b);
          const sChanged = !jsonEq(s, b);
          if (!lChanged && sChanged) {
            mergedRow[i] = clone(s);
          } else if (lChanged && sChanged && !jsonEq(l, s)) {
            // Реальный конфликт ячейка-в-ячейку: предпочитаем local, считаем конфликт
            conflicts += 1;
          }
        }
      } else {
        conflicts += 1;
      }
      merged.push(mergedRow);
    } else {
      // Ничего не менялось — берём server как canonical
      merged.push(clone(serverRow !== undefined ? serverRow : localRow));
    }
  }

  return { merged, conflicts };
}

function mergeSections(baselineSections, localSections, serverSections) {
  const baseline = Array.isArray(baselineSections) ? baselineSections : [];
  const local = Array.isArray(localSections) ? localSections : [];
  const server = Array.isArray(serverSections) ? serverSections : [];

  const baselineById = new Map(baseline.map((s) => [String(s?.id || ""), s]));
  const localById = new Map(local.map((s) => [String(s?.id || ""), s]));
  const serverById = new Map(server.map((s) => [String(s?.id || ""), s]));

  // Порядок секций — как в local
  const orderedIds = [];
  const seen = new Set();
  local.forEach((s) => {
    const id = String(s?.id || "");
    if (id && !seen.has(id)) {
      orderedIds.push(id);
      seen.add(id);
    }
  });
  server.forEach((s) => {
    const id = String(s?.id || "");
    if (id && !seen.has(id)) {
      orderedIds.push(id);
      seen.add(id);
    }
  });

  let totalConflicts = 0;
  const mergedSections = [];
  for (const id of orderedIds) {
    const baseSec = baselineById.get(id);
    const localSec = localById.get(id);
    const serverSec = serverById.get(id);

    if (!localSec && serverSec) {
      mergedSections.push(clone(serverSec));
      continue;
    }
    if (localSec && !serverSec) {
      mergedSections.push(clone(localSec));
      continue;
    }
    if (!localSec && !serverSec) continue;

    // Метаполя секции (title, columns) — если поменялись в local но не в server, берём local;
    // если в server, берём server; конфликт → local.
    const mergedSection = clone(localSec);
    ["title", "columns"].forEach((field) => {
      const b = baseSec?.[field];
      const l = localSec?.[field];
      const s = serverSec?.[field];
      const lChanged = !jsonEq(l, b);
      const sChanged = !jsonEq(s, b);
      if (!lChanged && sChanged) mergedSection[field] = clone(s);
    });

    // Сами строки — full 3-way merge
    const rowsRes = mergeRows(baseSec?.rows, localSec?.rows, serverSec?.rows);
    mergedSection.rows = rowsRes.merged;
    totalConflicts += rowsRes.conflicts;

    mergedSections.push(mergedSection);
  }

  return { sections: mergedSections, conflicts: totalConflicts };
}

/**
 * Главная функция: 3-way merge всего payload.
 * Возвращает { merged, conflicts } — merged готов к повторной отправке на сервер.
 */
export function mergePayloadOnConflict({ baseline, local, server }) {
  const b = baseline && typeof baseline === "object" ? baseline : {};
  const l = local && typeof local === "object" ? local : {};
  const s = server && typeof server === "object" ? server : {};

  const merged = clone(s);
  let totalConflicts = 0;

  const allKeys = new Set([...Object.keys(b), ...Object.keys(l), ...Object.keys(s)]);

  for (const key of allKeys) {
    const baseVal = b[key];
    const localVal = l[key];
    const serverVal = s[key];

    if (key === "sections") {
      const res = mergeSections(baseVal, localVal, serverVal);
      merged.sections = res.sections;
      totalConflicts += res.conflicts;
      continue;
    }

    if (localVal === undefined && serverVal !== undefined) {
      merged[key] = clone(serverVal);
      continue;
    }
    if (localVal !== undefined && serverVal === undefined) {
      merged[key] = clone(localVal);
      continue;
    }
    if (localVal === undefined && serverVal === undefined) {
      delete merged[key];
      continue;
    }

    const localChanged = !jsonEq(localVal, baseVal);
    const serverChanged = !jsonEq(serverVal, baseVal);

    if (localChanged && !serverChanged) {
      merged[key] = clone(localVal);
    } else if (!localChanged && serverChanged) {
      merged[key] = clone(serverVal);
    } else if (localChanged && serverChanged) {
      // Конфликт на top-level ключе (например, displaySettings). Берём local —
      // пользователь только что осознанно правил. Конфликт фиксируем для UI.
      merged[key] = clone(localVal);
      totalConflicts += 1;
    } else {
      merged[key] = clone(serverVal);
    }
  }

  return { merged, conflicts: totalConflicts };
}

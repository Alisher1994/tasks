/**
 * Обработка обновлений Telegram: /start → привязка chat_id к сотруднику, inline-кнопки у задачи,
 * комментарии, согласование закрытия, запись в taskHistory в payload.
 * Токен: TELEGRAM_BOT_TOKEN в окружении или displaySettings.telegramBotToken в JSON приложения (после синхронизации).
 */

import { promises as fsp } from "fs";
import path from "path";
import { randomBytes } from "crypto";

const STATUS_DECISION_OLD = "Треб. реш. рук.";
const STATUS_DECISION = "Требует решение руководителя";
const EMPLOYEE_STATUS_OPTIONS = ["В процессе", "Закрыт"];
const STATUS_EMOJI = {
  Новый: "🟣",
  "В процессе": "🟡",
  [STATUS_DECISION]: "🔴",
  [STATUS_DECISION_OLD]: "🔴",
  Закрыт: "🟢"
};
const TASK_COLUMNS = {
  number: 0,
  object: 1,
  status: 2,
  priority: 3,
  addedDate: 4,
  phase: 5,
  phaseSection: 6,
  phaseSubsection: 7,
  task: 8,
  responsible: 9,
  assignedResponsible: 10,
  note: 11,
  plan: 12,
  fact: 13,
  dueDate: 14,
  closedDate: 15,
  mediaBefore: 16,
  mediaAfter: 17,
  readState: 18,
  lastSentAt: 19,
  delayReason: 20,
  createdBy: 21,
  createdAt: 22,
  reassignReason: 23,
  reassignType: 24
};
const TASK_ROW_LENGTH = TASK_COLUMNS.reassignType + 1;

/** Маппинг типа переназначения в человекочитаемую метку. */
function getReassignTypeLabel(reasonType) {
  const t = String(reasonType || "").trim().toLowerCase();
  if (t === "mistake" || t === "objective") return "Ошибочная задача";
  if (t === "delegation" || t === "subjective") return "Делегирование задачи";
  return "";
}
const EMPLOYEE_COLUMNS = {
  id: 0,
  fullName: 1,
  department: 2,
  position: 3,
  phone: 4,
  telegram: 5,
  chatId: 6,
  activity: 7
};
const CONFIRM_ALWAYS_POSITIONS = new Set(["Администратор", "Генеральный директор"]);

const TASK_HISTORY_MAX = 300;

const PLACEHOLDERS = [
  ["[ид_задачи]", "number"],
  ["[Ид]", "number"],
  ["[название_задачи]", "task"],
  ["[статус]", "status"],
  ["[объект]", "object"],
  ["[причина_отставания]", "delayReason"],
  ["[причина_отстования]", "delayReason"],
  ["[причина_отставаний]", "delayReason"]
];

function normalizeTaskColumnLabel(raw) {
  return String(raw || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[^a-zа-я0-9]+/giu, "");
}

function normalizeTaskStatusValue(raw) {
  const value = String(raw || "").trim();
  if (value === STATUS_DECISION_OLD || value === STATUS_DECISION) return "В процессе";
  return value;
}

function normalizeTaskPriorityValue(raw) {
  const value = String(raw || "").trim();
  if (value === "Низкий") return "Средний";
  return value;
}

function remapTaskRowToCurrentOrder(sourceRow, sourceColumns) {
  const row = Array.isArray(sourceRow) ? sourceRow : [];
  const cols = Array.isArray(sourceColumns) ? sourceColumns : [];
  const byNorm = new Map();
  cols.forEach((label, index) => {
    const key = normalizeTaskColumnLabel(label);
    if (!key || byNorm.has(key)) return;
    byNorm.set(key, index);
  });
  const pick = (variants) => {
    for (const v of variants) {
      const i = byNorm.get(normalizeTaskColumnLabel(v));
      if (Number.isInteger(i) && i >= 0) return i;
    }
    return -1;
  };
  const has = (label) => byNorm.has(normalizeTaskColumnLabel(label));
  const out = new Array(TASK_ROW_LENGTH).fill("");
  const setByIndex = (targetIndex, sourceIndex, fallbackIndex = -1) => {
    let idx = sourceIndex;
    if (!Number.isInteger(idx) || idx < 0 || idx >= row.length) idx = fallbackIndex;
    out[targetIndex] = Number.isInteger(idx) && idx >= 0 && idx < row.length ? row[idx] : "";
  };

  const idxAssignedFromLabels = pick(["Исполнитель"]);
  const idxResponsibleFromLabels = pick(["Постановщик задачи", "Контролирующий ответственный"]);
  const idxAnswer = pick(["Ответственный"]);
  const idxTaskFromLabels = pick(["Задача", "Название задачи"]);
  const idxAssigned = idxAssignedFromLabels >= 0 ? idxAssignedFromLabels : has("Постановщик задачи") ? idxAnswer : -1;
  const idxResponsible = idxResponsibleFromLabels >= 0 ? idxResponsibleFromLabels : idxAssignedFromLabels === -1 ? idxAnswer : -1;

  setByIndex(TASK_COLUMNS.number, pick(["№", "ID", "Ид", "Ид задачи", "Номер"]), 0);
  setByIndex(TASK_COLUMNS.object, pick(["Название объекта", "Объект"]), 1);
  setByIndex(TASK_COLUMNS.status, pick(["Статус"]), 2);
  setByIndex(TASK_COLUMNS.priority, pick(["Приоритет"]), 3);
  setByIndex(TASK_COLUMNS.addedDate, pick(["Дата постановки задачи", "Дата добавления"]), 4);
  setByIndex(TASK_COLUMNS.phase, pick(["Фаза"]), 5);
  setByIndex(TASK_COLUMNS.phaseSection, pick(["Раздел"]), 6);
  setByIndex(TASK_COLUMNS.phaseSubsection, pick(["Подраздел"]), 7);
  setByIndex(TASK_COLUMNS.task, idxTaskFromLabels, 9);
  setByIndex(TASK_COLUMNS.responsible, idxResponsible, 10);
  setByIndex(TASK_COLUMNS.assignedResponsible, idxAssigned, 8);
  setByIndex(TASK_COLUMNS.note, pick(["Коментарии к задаче", "Комментарии к задаче", "Примичание", "Примечание"]), 11);
  setByIndex(TASK_COLUMNS.plan, pick(["Комментарии сотрудника (Результат)", "План решения (коммент сотрудника)", "План"]), 12);
  setByIndex(TASK_COLUMNS.fact, pick(["Коментарии администратора", "Комментарии администратора", "Факт исполнения", "Факт"]), 13);
  setByIndex(TASK_COLUMNS.dueDate, pick(["Плановый срок устранения", "Срок устранения", "Срок"]), 14);
  setByIndex(TASK_COLUMNS.closedDate, pick(["Факт даты устранения", "Дата устранения", "Дата закрытия"]), 15);
  setByIndex(TASK_COLUMNS.mediaBefore, pick(["Медиа до (5)", "Медиа до"]), 16);
  setByIndex(TASK_COLUMNS.mediaAfter, pick(["Медиа после (5)", "Медиа после"]), 17);
  setByIndex(TASK_COLUMNS.readState, pick(["Ознакомление"]), 18);
  setByIndex(TASK_COLUMNS.lastSentAt, pick(["Дата последней отправки"]), 19);
  setByIndex(TASK_COLUMNS.delayReason, pick(["Причина отставания"]), 20);
  setByIndex(TASK_COLUMNS.createdBy, pick(["Кем добавлена задача", "Добавил", "Автор добавления"]), 21);
  setByIndex(TASK_COLUMNS.createdAt, pick(["Время занесения в систему", "Дата и время занесения", "Дата занесения"]), 22);
  setByIndex(TASK_COLUMNS.reassignReason, pick(["Причина переназначения"]), 23);
  setByIndex(TASK_COLUMNS.reassignType, pick(["Тип переназначения"]), 24);
  out[TASK_COLUMNS.status] = normalizeTaskStatusValue(out[TASK_COLUMNS.status]);
  out[TASK_COLUMNS.priority] = normalizeTaskPriorityValue(out[TASK_COLUMNS.priority]);
  return out;
}

function normalizeTasksSectionByColumns(payload) {
  const tasks = payload?.sections?.find?.((s) => s.id === "tasks");
  if (!tasks || !Array.isArray(tasks.rows) || !Array.isArray(tasks.columns) || !tasks.columns.length) return;
  tasks.rows = tasks.rows.map((row) => {
    const next = remapTaskRowToCurrentOrder(row, tasks.columns);
    next[TASK_COLUMNS.status] = normalizeTaskStatusValue(next[TASK_COLUMNS.status]);
    next[TASK_COLUMNS.priority] = normalizeTaskPriorityValue(next[TASK_COLUMNS.priority]);
    return next;
  });
}

function getTasksSection(payload) {
  const sections = payload?.sections || [];
  return sections.find((s) => s.id === "tasks");
}

function getEmployeesSection(payload) {
  const sections = payload?.sections || [];
  return sections.find((s) => s.id === "employees");
}

function getDelayReasonsSection(payload) {
  const sections = payload?.sections || [];
  return sections.find((s) => s.id === "delayReasons");
}

function getDelayReasonOptions(payload) {
  const section = getDelayReasonsSection(payload);
  const rows = Array.isArray(section?.rows) ? section.rows : [];
  const seen = new Set();
  const out = [];
  for (const row of rows) {
    const reason = String(row?.[1] || "").trim();
    if (!reason) continue;
    const key = reason.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(reason);
  }
  return out;
}

function findTaskRow(tasksSection, taskNumber) {
  const rows = tasksSection?.rows || [];
  const want = String(taskNumber ?? "").trim();
  return rows.find((row) => String(row[TASK_COLUMNS.number] ?? "").trim() === want);
}

function findEmployeeByChatId(empSection, chatId) {
  const want = String(chatId ?? "").trim();
  const rows = empSection?.rows || [];
  return rows.find((row) => String(row[EMPLOYEE_COLUMNS.chatId] ?? "").trim() === want);
}

function appendTaskHistory(payload, taskId, who, actionText) {
  const id = String(taskId ?? "").trim() || "—";
  const action = String(actionText ?? "").trim();
  if (!action) return;
  if (!payload.taskHistory || typeof payload.taskHistory !== "object") {
    payload.taskHistory = {};
  }
  const store = payload.taskHistory;
  if (!store[id]) store[id] = [];
  store[id].unshift({ t: Date.now(), who: String(who || "Telegram"), action });
  if (store[id].length > TASK_HISTORY_MAX) {
    store[id].length = TASK_HISTORY_MAX;
  }
}

function applySimpleTemplate(template, row) {
  let out = String(template || "");
  for (const [token, colKey] of PLACEHOLDERS) {
    const col = TASK_COLUMNS[colKey];
    if (col === undefined) continue;
    const val = String(row[col] ?? "");
    out = out.split(token).join(val);
  }
  return out.replace(/\r\n?/g, "\n");
}

function parseRuDateToYmd(raw) {
  const src = String(raw || "").trim();
  const m = src.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})$/);
  if (!m) return null;
  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);
  if (!Number.isInteger(day) || !Number.isInteger(month) || !Number.isInteger(year)) return null;
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (dt.getUTCFullYear() !== year || dt.getUTCMonth() !== month - 1 || dt.getUTCDate() !== day) return null;
  return { year, month, day };
}

function getTodayYmdInTimeZone(timeZone = "UTC") {
  try {
    const parts = new Intl.DateTimeFormat("ru-RU", {
      timeZone,
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    }).formatToParts(new Date());
    const pick = (type) => Number(parts.find((p) => p.type === type)?.value || 0);
    const day = pick("day");
    const month = pick("month");
    const year = pick("year");
    if (!day || !month || !year) return null;
    return { year, month, day };
  } catch (_) {
    return null;
  }
}

function compareYmd(a, b) {
  if (a.year !== b.year) return a.year - b.year;
  if (a.month !== b.month) return a.month - b.month;
  return a.day - b.day;
}

function isDelayReasonAllowedNow(row, appTimeZone = "UTC") {
  const dueRaw = String(row?.[TASK_COLUMNS.dueDate] || "").trim();
  if (!dueRaw) return false;
  const due = parseRuDateToYmd(dueRaw);
  if (!due) return false;
  const today = getTodayYmdInTimeZone(appTimeZone) || parseRuDateToYmd(formatRuDate(new Date(), "UTC"));
  if (!today) return false;
  // Причина отставания нужна только после наступления просрочки (сегодня > планового срока).
  return compareYmd(today, due) > 0;
}

function isTaskOverdueWithoutDelayReason(row, appTimeZone = "UTC") {
  const status = normalizeTaskStatusValue(String(row?.[TASK_COLUMNS.status] || "").trim());
  if (status === "Закрыт") return false;
  return isDelayReasonAllowedNow(row, appTimeZone)
    && !String(row?.[TASK_COLUMNS.delayReason] || "").trim();
}

function mainKeyboard(taskNumber, row, appTimeZone = "UTC") {
  const status = normalizeTaskStatusValue(String(row?.[TASK_COLUMNS.status] || "").trim());
  if (status === "Закрыт") return [];
  const actualTaskNumber = String(row?.[TASK_COLUMNS.number] || taskNumber || "").trim() || String(taskNumber || "").trim();
  const n = encodeTaskNum(actualTaskNumber);
  const overdue = isDelayReasonAllowedNow(row, appTimeZone);
  const keyboard = [
    [
      { text: "⌛️ Сменить статус", callback_data: cb(n, "sm") },
      { text: "🗣 Комментарий", callback_data: cb(n, "cm") }
    ]
  ];
  if (overdue) {
    // Просроченная задача: переназначение скрыто, ветка реассайна идёт через
    // «Причина отставания → Внутренний фактор → выбор отдела/сотрудника».
    keyboard.push([{ text: "🚧 Причина отставания", callback_data: cb(n, "dr") }]);
  } else {
    keyboard.push([{ text: "🔁 Переназначить задачу", callback_data: cb(n, "ra") }]);
  }
  return keyboard;
}

function quickUserKeyboard() {
  return {
    keyboard: [[{ text: "Мои задачи" }, { text: "Просроченные задачи" }]],
    is_persistent: true,
    resize_keyboard: true
  };
}

function parseMyTasksCallbackData(data) {
  const raw = String(data || "").trim();
  const parts = raw.split("|");
  if (parts[0] !== "mt" || parts.length < 3) return null;
  return {
    action: String(parts[1] || "").trim(),
    args: parts.slice(2).map((x) => String(x || "").trim())
  };
}

function myTasksCb(action, ...args) {
  return `mt|${String(action || "").trim()}|${args.map((x) => String(x || "").trim()).join("|")}`.slice(0, 64);
}

function findEmployeeByFullName(empSection, fullName) {
  const want = String(fullName || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
  if (!want) return null;
  const rows = empSection?.rows || [];
  return (
    rows.find((row) => String(row[EMPLOYEE_COLUMNS.fullName] || "").trim().replace(/\s+/g, " ").toLowerCase() === want) || null
  );
}

function parseTaskAssigneeNames(rawValue) {
  const seen = new Set();
  return String(rawValue || "")
    .split(",")
    .map((x) => normalizePersonName(x))
    .filter(Boolean)
    .filter((name) => {
      const key = name.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function ensureTaskMultiStateStore(payload) {
  if (!payload.taskMultiState || typeof payload.taskMultiState !== "object" || Array.isArray(payload.taskMultiState)) {
    payload.taskMultiState = {};
  }
  return payload.taskMultiState;
}

function ensureTaskCloseMetaStore(payload) {
  if (!payload.taskCloseMeta || typeof payload.taskCloseMeta !== "object" || Array.isArray(payload.taskCloseMeta)) {
    payload.taskCloseMeta = {};
  }
  return payload.taskCloseMeta;
}

function setSingleTaskClosedAt(payload, row, closedAtText) {
  const taskId = String(row?.[TASK_COLUMNS.number] ?? "").trim();
  if (!taskId) return;
  const store = ensureTaskCloseMetaStore(payload);
  const value = String(closedAtText || "").trim();
  if (!value) {
    delete store[taskId];
    return;
  }
  store[taskId] = { closedAt: value };
}

function getTaskMultiStateForRow(payload, row, { create = false } = {}) {
  const taskId = String(row?.[TASK_COLUMNS.number] ?? "").trim();
  if (!taskId) return null;
  const store = ensureTaskMultiStateStore(payload);
  if (!store[taskId] || typeof store[taskId] !== "object" || Array.isArray(store[taskId])) {
    if (!create) return null;
    store[taskId] = {};
  }
  return store[taskId];
}

function syncTaskMultiStateForRow(payload, row) {
  const assignees = parseTaskAssigneeNames(row?.[TASK_COLUMNS.assignedResponsible]);
  const taskId = String(row?.[TASK_COLUMNS.number] ?? "").trim();
  if (!taskId) return;
  const store = ensureTaskMultiStateStore(payload);
  const prevState = store[taskId] && typeof store[taskId] === "object" && !Array.isArray(store[taskId]) ? store[taskId] : {};
  const existingFiles = Array.isArray(prevState.__files) ? prevState.__files : [];
  if (assignees.length <= 1) {
    if (existingFiles.length > 0) {
      store[taskId] = { __files: existingFiles };
    } else {
      delete store[taskId];
    }
    return;
  }
  const state = getTaskMultiStateForRow(payload, row, { create: true });
  if (existingFiles.length > 0) {
    state.__files = existingFiles;
  }
  const known = new Set(assignees.map((x) => x.toLowerCase()));
  assignees.forEach((name) => {
    if (!state[name] || typeof state[name] !== "object" || Array.isArray(state[name])) {
      state[name] = {};
    }
    if (!String(state[name].status || "").trim()) {
      state[name].status = String(row[TASK_COLUMNS.status] || "").trim() || "Новый";
    }
  });
  Object.keys(state).forEach((name) => {
    if (String(name).startsWith("__")) return;
    if (!known.has(String(name).toLowerCase())) delete state[name];
  });
}

function getEmployeeNameByChatId(empSection, chatId) {
  const emp = findEmployeeByChatId(empSection, chatId);
  return normalizePersonName(emp?.[EMPLOYEE_COLUMNS.fullName] || "");
}

function getTaskAssigneeNameByChat(payload, row, chatId) {
  const employees = getEmployeesSection(payload);
  const clickName = getEmployeeNameByChatId(employees, chatId);
  if (!clickName) return "";
  const assignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  const wanted = new Set(assignees.map((x) => x.toLowerCase()));
  if (!wanted.has(clickName.toLowerCase())) return "";
  return assignees.find((x) => x.toLowerCase() === clickName.toLowerCase()) || "";
}

function getTaskAssigneeStateEntry(stateMap, assigneeName) {
  if (!stateMap || typeof stateMap !== "object" || Array.isArray(stateMap)) return null;
  const exact = stateMap[assigneeName];
  if (exact && typeof exact === "object" && !Array.isArray(exact)) {
    return { key: assigneeName, state: exact };
  }
  const wanted = String(assigneeName || "").trim().toLowerCase();
  if (!wanted) return null;
  for (const key of Object.keys(stateMap)) {
    if (String(key).startsWith("__")) continue;
    if (String(key).toLowerCase() !== wanted) continue;
    const state = stateMap[key];
    if (state && typeof state === "object" && !Array.isArray(state)) {
      return { key, state };
    }
  }
  return null;
}

function buildTaskRowForAssignee(payload, row, assigneeName) {
  const assignee = normalizePersonName(assigneeName);
  if (!assignee) return row;
  const stateMap = getTaskMultiStateForRow(payload, row);
  const found = getTaskAssigneeStateEntry(stateMap, assignee);
  if (!found?.state) return row;
  const state = found.state;
  const scoped = Array.isArray(row) ? row.slice() : [];
  const status = String(state.status || "").trim();
  const comment = String(state.comment || "").trim();
  const due = String(state.dueDate || "").trim();
  if (status) scoped[TASK_COLUMNS.status] = status;
  if (comment) scoped[TASK_COLUMNS.plan] = comment;
  if (due) scoped[TASK_COLUMNS.dueDate] = due;
  scoped[TASK_COLUMNS.assignedResponsible] = assignee;
  return scoped;
}

function buildTaskRowForChat(payload, row, chatId, taskIdentity = "") {
  const baseTaskId = String(row?.[TASK_COLUMNS.number] || "").trim();
  const wantedTaskId = String(taskIdentity || baseTaskId).trim();
  const parsedChild = parseReassignChildTaskId(wantedTaskId);
  if (parsedChild) {
    const exactReassign = getActiveReassignForTask(payload, parsedChild.code);
    if (exactReassign) {
      return buildTaskRowForReassign(row, exactReassign, parsedChild.baseTaskId || baseTaskId);
    }
  }
  const activeReassign = getActiveReassignForChat(payload, row, chatId, wantedTaskId);
  if (activeReassign) {
    return buildTaskRowForReassign(row, activeReassign, baseTaskId);
  }
  const assigneeName = getTaskAssigneeNameByChat(payload, row, chatId);
  if (!assigneeName) return row;
  return buildTaskRowForAssignee(payload, row, assigneeName);
}

function buildTaskRowForReassign(row, reassignEntry, fallbackTaskId = "") {
  const scoped = Array.isArray(row) ? row.slice() : [];
  const baseTaskId = String(row?.[TASK_COLUMNS.number] || fallbackTaskId || "").trim();
  const scopedTaskId = String(reassignEntry?.code || buildReassignChildTaskId(baseTaskId, 1)).trim();
  const activeTo = normalizePersonName(reassignEntry?.to || reassignEntry?.toEmployeeName || "");
  scoped[TASK_COLUMNS.number] = scopedTaskId;
  scoped[TASK_COLUMNS.status] = String(reassignEntry?.currentStatus || "В процессе").trim() || "В процессе";
  if (activeTo) scoped[TASK_COLUMNS.assignedResponsible] = activeTo;
  scoped[TASK_COLUMNS.reassignReason] = String(reassignEntry?.reasonText || "").trim();
  scoped[TASK_COLUMNS.plan] = String(reassignEntry?.comment || "").trim();
  const readAt = String(reassignEntry?.readAt || "").trim();
  scoped[TASK_COLUMNS.readState] = readAt ? composeReadStateValue(true, readAt) : composeReadStateValue(false, "—");
  return scoped;
}

function updateTaskAggregateStatusFromMulti(payload, row, appTimeZone) {
  const assignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  if (assignees.length <= 1) return;
  syncTaskMultiStateForRow(payload, row);
  const state = getTaskMultiStateForRow(payload, row) || {};
  const statuses = assignees.map((name) => String(state?.[name]?.status || "").trim() || "Новый");
  if (!statuses.length) return;
  if (statuses.every((s) => s === "Закрыт")) {
    row[TASK_COLUMNS.status] = "Закрыт";
    if (!String(row[TASK_COLUMNS.closedDate] || "").trim()) {
      row[TASK_COLUMNS.closedDate] = formatRuDate(new Date(), appTimeZone || "UTC");
    }
    return;
  }
  row[TASK_COLUMNS.closedDate] = "";
  if (statuses.some((s) => s === "В процессе")) {
    row[TASK_COLUMNS.status] = "В процессе";
    return;
  }
  row[TASK_COLUMNS.status] = "Новый";
}

function findDepartmentRowByName(payload, departmentName) {
  const want = String(departmentName || "")
    .trim()
    .replace(/\s+/g, " ");
  if (!want) return null;
  const sections = payload?.sections || [];
  const departments = sections.find((s) => s.id === "departments");
  const rows = departments?.rows || [];
  return rows.find((row) => String(row[1] || "").trim().replace(/\s+/g, " ") === want) || null;
}

function isDepartmentHeadsCloseConfirmEnabled(payload) {
  return payload?.displaySettings?.telegramDepartmentHeadsCanConfirmClose !== false;
}

function getDepartmentHeadChatIdsForTask(payload, row) {
  if (!isDepartmentHeadsCloseConfirmEnabled(payload)) return [];
  const employees = getEmployeesSection(payload);
  const out = new Set();
  const assignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  for (const assignedName of assignees) {
    const assignedEmp = findEmployeeByFullName(employees, assignedName);
    if (!assignedEmp) continue;
    const departmentName = String(assignedEmp[EMPLOYEE_COLUMNS.department] || "").trim();
    if (!departmentName) continue;
    const depRow = findDepartmentRowByName(payload, departmentName);
    if (!depRow) continue;
    const headName = String(depRow[2] || "").trim();
    if (!headName) continue;
    const headEmp = findEmployeeByFullName(employees, headName);
    if (!headEmp) continue;
    if (String(headEmp[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") continue;
    const chatId = String(headEmp[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (chatId) out.add(chatId);
  }
  return Array.from(out);
}

function getAlwaysConfirmChatIds(payload) {
  const employees = getEmployeesSection(payload);
  const rows = employees?.rows || [];
  const out = new Set();
  for (const r of rows) {
    if (String(r[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") continue;
    const pos = String(r[EMPLOYEE_COLUMNS.position] || "").trim();
    if (!CONFIRM_ALWAYS_POSITIONS.has(pos)) continue;
    const chatId = String(r[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (chatId) out.add(chatId);
  }
  return Array.from(out);
}

function backOnlyKeyboard(taskNumber) {
  const n = encodeTaskNum(taskNumber);
  return [[{ text: "⬅️ Назад", callback_data: cb(n, "bk") }]];
}

function encodeTaskNum(num) {
  return String(num ?? "")
    .trim()
    .replace(/\|/g, "·")
    .slice(0, 48);
}

function cb(num, code) {
  const n = encodeTaskNum(num);
  const s = `t|${n}|${code}`;
  if (s.length > 64) {
    return `t|${n.slice(0, 40)}|${code}`.slice(0, 64);
  }
  return s;
}

function parseCallbackData(data) {
  const raw = String(data || "");
  const parts = raw.split("|");
  if (parts[0] !== "t" || parts.length < 3) return null;
  return {
    taskNum: parts[1].replace(/·/g, "|"),
    action: parts[2],
    rest: parts.slice(3)
  };
}

function parseOverdueCallbackData(data) {
  const raw = String(data || "").trim();
  const parts = raw.split("|");
  if (parts[0] !== "ov" || parts.length < 2) return null;
  return {
    action: String(parts[1] || "").trim()
  };
}

function ensureReassignRequestsStore(payload) {
  if (!payload.telegramReassignRequests || typeof payload.telegramReassignRequests !== "object" || Array.isArray(payload.telegramReassignRequests)) {
    payload.telegramReassignRequests = {};
  }
  return payload.telegramReassignRequests;
}

function ensureTaskReassignLogStore(payload) {
  if (!payload.taskReassignLog || typeof payload.taskReassignLog !== "object" || Array.isArray(payload.taskReassignLog)) {
    payload.taskReassignLog = {};
  }
  return payload.taskReassignLog;
}

function hasPendingMistakeReassignRequest(payload, taskId) {
  const task = String(taskId || "").trim();
  if (!task) return false;
  const reqStore = payload?.telegramReassignRequests && typeof payload.telegramReassignRequests === "object"
    ? payload.telegramReassignRequests
    : {};
  return Object.values(reqStore).some((item) => {
    if (!item || typeof item !== "object") return false;
    const reasonType = String(item.reasonType || "").trim().toLowerCase();
    return String(item.taskId || "").trim() === task
      && String(item.status || "").trim() === "pending"
      && (reasonType === "mistake" || reasonType === "objective")
      && !String(item.toEmployeeName || "").trim();
  });
}

function buildReassignChildTaskId(baseTaskId, seq = 1) {
  const base = String(baseTaskId || "").trim();
  const n = Math.max(1, Number(seq) || 1);
  if (!base) return `${n}`;
  return `${base}/${n}`;
}

function parseReassignChildTaskId(taskId) {
  const raw = String(taskId || "").trim();
  const match = raw.match(/^(.+?)\/(\d+)$/);
  if (!match) return null;
  const seq = Number(match[2]);
  if (!Number.isFinite(seq) || seq < 1) return null;
  return {
    baseTaskId: String(match[1] || "").trim(),
    seq,
    code: raw
  };
}

function getActiveReassignForTask(payload, taskId) {
  const id = String(taskId || "").trim();
  if (!id) return null;
  const parsedChild = parseReassignChildTaskId(id);
  const baseId = parsedChild?.baseTaskId || id;
  const logArr = Array.isArray(payload?.taskReassignLog?.[baseId]) ? payload.taskReassignLog[baseId] : [];
  if (parsedChild) {
    return logArr.find((x) => String(x?.code || "").trim() === parsedChild.code) || null;
  }
  const approved = logArr.find((x) => String(x?.status || "").trim() === "approved");
  if (!approved) return null;
  return approved;
}

function getActiveReassignForChat(payload, row, chatId, taskIdentity = "") {
  const wantedTaskId = String(taskIdentity || row?.[TASK_COLUMNS.number] || "").trim();
  const active = getActiveReassignForTask(payload, wantedTaskId);
  if (!active) return null;
  const toName = normalizePersonName(active.to || active.toEmployeeName || "");
  if (!toName) return null;
  const employees = getEmployeesSection(payload);
  const clickEmp = findEmployeeByChatId(employees, String(chatId || "").trim());
  const clickName = normalizePersonName(clickEmp?.[EMPLOYEE_COLUMNS.fullName] || "");
  if (!clickName || clickName.toLowerCase() !== toName.toLowerCase()) return null;
  return active;
}

function buildDepartmentOptions(payload) {
  const sections = payload?.sections || [];
  const departments = sections.find((s) => s && s.id === "departments");
  const rows = Array.isArray(departments?.rows) ? departments.rows : [];
  const out = [];
  const seen = new Set();
  for (const row of rows) {
    const name = String(row?.[1] || "").trim();
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(name);
  }
  return out.sort((a, b) => a.localeCompare(b, "ru"));
}

function buildDepartmentEmployees(payload, departmentName) {
  const employees = getEmployeesSection(payload);
  const rows = Array.isArray(employees?.rows) ? employees.rows : [];
  const want = String(departmentName || "").trim().toLowerCase();
  const out = [];
  for (const row of rows) {
    const fio = String(row?.[EMPLOYEE_COLUMNS.fullName] || "").trim();
    if (!fio) continue;
    const dep = String(row?.[EMPLOYEE_COLUMNS.department] || "").trim().toLowerCase();
    if (!want || dep !== want) continue;
    out.push({
      id: String(row?.[EMPLOYEE_COLUMNS.id] || "").trim(),
      fullName: fio,
      chatId: String(row?.[EMPLOYEE_COLUMNS.chatId] || "").trim(),
      telegramConnected: String(row?.[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен"
    });
  }
  return out.sort((a, b) => a.fullName.localeCompare(b.fullName, "ru"));
}

function getNextReassignCode(payload, taskId) {
  const task = String(taskId || "").trim();
  if (!task) return "1";
  let max = 0;
  const logArr = Array.isArray(payload?.taskReassignLog?.[task]) ? payload.taskReassignLog[task] : [];
  for (const item of logArr) {
    const code = String(item?.code || "").trim();
    const parsed = parseReassignChildTaskId(code);
    if (!parsed || parsed.baseTaskId !== task) continue;
    const n = parsed.seq;
    if (Number.isFinite(n) && n > max) max = n;
  }
  const reqStore = payload?.telegramReassignRequests && typeof payload.telegramReassignRequests === "object"
    ? payload.telegramReassignRequests
    : {};
  for (const key of Object.keys(reqStore)) {
    const item = reqStore[key];
    if (String(item?.taskId || "").trim() !== task) continue;
    const code = String(item?.code || "").trim();
    const parsed = parseReassignChildTaskId(code);
    if (!parsed || parsed.baseTaskId !== task) continue;
    const n = parsed.seq;
    if (Number.isFinite(n) && n > max) max = n;
  }
  return buildReassignChildTaskId(task, max + 1);
}

function statusLabelWithEmoji(status) {
  const s = String(status || "").trim();
  if (!s) return "";
  const emoji = STATUS_EMOJI[s] || "•";
  const text = s === "Закрыт" ? "Закрыто" : s;
  return `${emoji} ${text}`;
}

function buildOverdueSummaryText(overdueCount) {
  const count = Number(overdueCount) || 0;
  return [
    "⚠️ Уведомление",
    "",
    `Количество просроченных задач: ${count} шт.`
  ].join("\n");
}

function overdueSummaryKeyboard(overdueCount = 0) {
  const count = Math.max(0, Number(overdueCount) || 0);
  return {
    inline_keyboard: [[{ text: `📋 Посмотреть - ${count} шт`, callback_data: "ov|ls" }]]
  };
}

function isPrivilegedViewerChat(payload, chatId) {
  const chat = String(chatId || "").trim();
  if (!chat) return false;
  const adminChat = String(payload?.displaySettings?.telegramAdminChatId || "").trim();
  if (adminChat && adminChat === chat) return true;

  const employees = getEmployeesSection(payload);
  const emp = findEmployeeByChatId(employees, chat);
  if (!emp) return false;
  if (String(emp[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") return false;

  const position = String(emp[EMPLOYEE_COLUMNS.position] || "").trim().toLowerCase();
  if (!position) return false;
  if (position === "администратор") return true;
  return position.includes("директор");
}

function getOverdueRowsForChat(payload, chatId, appTimeZone = "UTC") {
  const tasks = getTasksSection(payload);
  const rows = Array.isArray(tasks?.rows) ? tasks.rows : [];
  const today = getTodayYmdInTimeZone(appTimeZone);
  if (!today) return [];
  const chat = String(chatId || "").trim();
  const out = [];
  for (const row of rows) {
    if (!Array.isArray(row)) continue;
    if (!isTaskVisibleForChat(payload, row, chat)) continue;
    const status = normalizeTaskStatusValue(String(row[TASK_COLUMNS.status] || "").trim());
    if (status === "Закрыт") continue;
    const dueRaw = String(row[TASK_COLUMNS.dueDate] || "").trim();
    const due = parseRuDateToYmd(dueRaw);
    if (!due) continue;
    const diff = compareYmd(today, due);
    if (diff <= 0) continue;
    out.push({ row, overdueDays: diff });
  }
  out.sort((a, b) => b.overdueDays - a.overdueDays);
  return out;
}

function overdueListKeyboard(payload, overdueRows) {
  const inline = overdueRows.map(({ row }) => {
    const taskId = String(row?.[TASK_COLUMNS.number] || "").trim() || "—";
    const status = normalizeTaskStatusValue(String(row?.[TASK_COLUMNS.status] || "").trim());
    const emoji = STATUS_EMOJI[status] || "⚪";
    return [{ text: `${emoji} Задача № ${taskId}`, callback_data: cb(taskId, "bk") }];
  });
  inline.push([{ text: "⬅️ Назад", callback_data: "ov|bk" }]);
  return { inline_keyboard: inline };
}

function isTaskActionAllowedForChat(payload, row, chatId) {
  const actionChats = getTaskActionChatIds(payload, row);
  const chat = String(chatId || "").trim();
  if (!chat) return false;
  if (!actionChats.size) return false;
  return actionChats.has(chat);
}

function isTaskVisibleForChat(payload, row, chatId) {
  return isTaskActionAllowedForChat(payload, row, chatId) || isPrivilegedViewerChat(payload, chatId);
}

function getMyTasksForChat(payload, chatId) {
  const tasks = getTasksSection(payload);
  const rows = Array.isArray(tasks?.rows) ? tasks.rows : [];
  return rows.filter((row) => Array.isArray(row) && isTaskVisibleForChat(payload, row, chatId));
}

function getMyTaskObjects(rows) {
  const seen = new Set();
  const out = [];
  for (const row of rows) {
    const objectName = String(row?.[TASK_COLUMNS.object] || "").trim() || "—";
    const key = objectName.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(objectName);
  }
  return out.sort((a, b) => a.localeCompare(b, "ru"));
}

function myTaskStatusByCode(code) {
  const c = String(code || "").trim();
  if (c === "all") return "";
  if (c === "new") return "Новый";
  if (c === "prog") return "В процессе";
  if (c === "closed") return "Закрыт";
  if (c === "overdue") return "__overdue__";
  return "";
}

function myTaskStatusLabel(code) {
  const c = String(code || "").trim();
  if (c === "all") return "Все статусы";
  if (c === "new") return "Новый";
  if (c === "prog") return "В процессе";
  if (c === "closed") return "Закрыт";
  if (c === "overdue") return "Просроченные";
  return "Все статусы";
}

function getMyTaskStatusCounts(rows, objectName, appTz) {
  const all = filterMyTasksBySelection(rows, objectName, "all", appTz).length;
  const next = filterMyTasksBySelection(rows, objectName, "new", appTz).length;
  const progress = filterMyTasksBySelection(rows, objectName, "prog", appTz).length;
  const closed = filterMyTasksBySelection(rows, objectName, "closed", appTz).length;
  const overdue = filterMyTasksBySelection(rows, objectName, "overdue", appTz).length;
  return { all, next, progress, closed, overdue };
}

function getStatusPickKeyboardForObject(objectIndex, taskRows, objectName, appTz) {
  const oi = String(objectIndex || "0").trim();
  const counts = getMyTaskStatusCounts(Array.isArray(taskRows) ? taskRows : [], objectName, appTz);
  const keyboardRows = [
    [{ text: `🧾 Все статусы - ${counts.all} шт`, callback_data: myTasksCb("st", oi, "all") }],
    [{ text: `🟣 Новый - ${counts.next} шт`, callback_data: myTasksCb("st", oi, "new") }],
    [{ text: `🟡 В процессе - ${counts.progress} шт`, callback_data: myTasksCb("st", oi, "prog") }],
    [{ text: `🟢 Закрыт - ${counts.closed} шт`, callback_data: myTasksCb("st", oi, "closed") }],
    [{ text: `⚠️ Просроченные - ${counts.overdue} шт`, callback_data: myTasksCb("st", oi, "overdue") }],
    [{ text: "⬅️ Назад", callback_data: myTasksCb("bo", "0") }]
  ];
  return { inline_keyboard: keyboardRows };
}

function filterMyTasksBySelection(rows, objectName, statusCode, appTz) {
  const wantObject = String(objectName || "").trim();
  const statusNeed = myTaskStatusByCode(statusCode);
  const today = getTodayYmdInTimeZone(appTz || "UTC");
  return rows.filter((row) => {
    const objectOk = !wantObject || String(row?.[TASK_COLUMNS.object] || "").trim() === wantObject;
    if (!objectOk) return false;
    const st = normalizeTaskStatusValue(String(row?.[TASK_COLUMNS.status] || "").trim());
    if (statusNeed === "__overdue__") {
      if (st === "Закрыт") return false;
      const due = parseRuDateToYmd(String(row?.[TASK_COLUMNS.dueDate] || "").trim());
      if (!due || !today) return false;
      return compareYmd(today, due) > 0;
    }
    if (!statusNeed) return true;
    return st === statusNeed;
  });
}

function myTasksListKeyboard(rows, payload, chatId) {
  const inline = rows.map((row) => {
    const scoped = buildTaskRowForChat(payload, row, chatId);
    const taskId = String(scoped?.[TASK_COLUMNS.number] || "").trim() || "—";
    const status = normalizeTaskStatusValue(String(scoped?.[TASK_COLUMNS.status] || "").trim());
    const emoji = STATUS_EMOJI[status] || "⚪";
    return [{ text: `${emoji} Задача № ${taskId}`, callback_data: cb(taskId, "bk") }];
  });
  inline.push([{ text: "⬅️ К объектам", callback_data: myTasksCb("bo", "0") }]);
  return { inline_keyboard: inline.slice(0, 70) };
}

// Маркер карточки задачи: buildFullTaskMessage всегда начинается с этой строки.
// Если text/caption начинается с маркера — отправляем с parse_mode HTML, чтобы
// mention-ссылки (<a href="tg://user?id=...">) превратились в кликабельные ники.
const TG_TASK_MSG_HTML_MARKER = "📝 Задача №";

async function tg(token, method, body) {
  let outBody = body;
  if (outBody && typeof outBody === "object" && !outBody.parse_mode) {
    const candidate = typeof outBody.text === "string"
      ? outBody.text
      : typeof outBody.caption === "string"
        ? outBody.caption
        : "";
    if (candidate.startsWith(TG_TASK_MSG_HTML_MARKER)) {
      outBody = { ...outBody, parse_mode: "HTML" };
    }
  }
  const url = `https://api.telegram.org/bot${token}/${method}`;
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(outBody)
  });
  return r.json();
}

function splitMediaItems(value) {
  return String(value || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean)
    .slice(0, 5);
}

function joinMediaItems(items) {
  return items.filter(Boolean).slice(0, 5).join(", ");
}

function pickTelegramPhotoFileName(filePath, fileId) {
  const raw = String(filePath || "").trim();
  if (raw) {
    const parts = raw.split("/");
    const last = String(parts[parts.length - 1] || "").trim();
    if (last) return last;
  }
  return `${String(fileId || "").trim() || Date.now()}.jpg`;
}

async function addTelegramPhotoToTaskMediaAfter(row, token, fileId) {
  const fr = await tg(token, "getFile", { file_id: fileId });
  if (!fr?.ok || !fr?.result) return null;
  const filePath = String(fr.result.file_path || "").trim();
  const fileName = pickTelegramPhotoFileName(filePath, fileId);
  const storedValue = filePath || fileName;
  const items = splitMediaItems(row[TASK_COLUMNS.mediaAfter]);
  if (!items.includes(storedValue)) {
    if (items.length >= 5) items.shift();
    items.push(storedValue);
    row[TASK_COLUMNS.mediaAfter] = joinMediaItems(items);
  }
  return { fileName, filePath, storedValue };
}

function getMediaStoragePath() {
  const fromEnv = String(process.env.MEDIA_STORAGE_PATH || "").trim();
  if (fromEnv) return fromEnv;
  return path.join(process.cwd(), "media");
}

const TG_FILE_EXT_TO_MIME = {
  jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp",
  pdf: "application/pdf",
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  txt: "text/plain",
  csv: "text/csv",
  zip: "application/zip",
  rar: "application/x-rar-compressed",
  mp4: "video/mp4",
  mov: "video/quicktime",
  mp3: "audio/mpeg",
  ogg: "audio/ogg",
  wav: "audio/wav"
};

/**
 * Скачивает файл из Telegram по file_id, сохраняет в локальное медиа-хранилище
 * (то же что использует /api/media/upload) и возвращает структуру для записи в
 * payload.taskAttachments[taskId].
 *
 * Возвращает null при ошибке — в этом случае закрытие задачи всё равно можно
 * довести до конца, просто без файла.
 */
async function downloadTelegramFileAsAttachment(token, fileId, suggestedName = "", suggestedMime = "") {
  try {
    const fr = await tg(token, "getFile", { file_id: fileId });
    if (!fr?.ok || !fr?.result) return null;
    const filePath = String(fr.result.file_path || "").trim();
    if (!filePath) return null;

    const url = `https://api.telegram.org/file/bot${token}/${filePath}`;
    const resp = await fetch(url);
    if (!resp.ok) return null;
    const buf = Buffer.from(await resp.arrayBuffer());

    const rawExt = String(filePath.split(".").pop() || "").toLowerCase();
    const ext = /^[a-z0-9]{1,8}$/.test(rawExt) ? rawExt : "bin";
    const fileName = `${Date.now()}-${randomBytes(6).toString("hex")}.${ext}`;
    const dir = getMediaStoragePath();
    await fsp.mkdir(dir, { recursive: true });
    await fsp.writeFile(path.join(dir, fileName), buf);

    const baseName = suggestedName.trim() || filePath.split("/").pop() || fileName;
    const mime = suggestedMime.trim() || TG_FILE_EXT_TO_MIME[ext] || "application/octet-stream";
    return {
      stored: `/media/${encodeURIComponent(fileName)}`,
      name: baseName,
      type: mime,
      size: buf.length
    };
  } catch (e) {
    console.error("[tg-attachment] download failed:", e);
    return null;
  }
}

function appendTaskAttachmentEntry(payload, taskId, entry) {
  if (!entry || !entry.stored) return false;
  if (!payload.taskAttachments || typeof payload.taskAttachments !== "object" || Array.isArray(payload.taskAttachments)) {
    payload.taskAttachments = {};
  }
  const id = String(taskId || "").trim();
  if (!id) return false;
  if (!Array.isArray(payload.taskAttachments[id])) payload.taskAttachments[id] = [];
  payload.taskAttachments[id].push({
    id: `f_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`,
    stored: entry.stored,
    name: entry.name || "",
    type: entry.type || "",
    size: Number(entry.size) || 0,
    addedAt: new Date().toISOString()
  });
  return true;
}

function shortTaskCaption(row) {
  const num = String(row[TASK_COLUMNS.number] ?? "").trim();
  const title = String(row[TASK_COLUMNS.task] ?? "").trim();
  const st = String(row[TASK_COLUMNS.status] ?? "").trim();
  const stLine = st ? `${STATUS_EMOJI[st] || "⚪"} ${st}` : "—";
  return `№ ${num}${title ? `: ${title}` : ""}\nСтатус: ${stLine}`;
}

function resolveAppTimeZone(payload) {
  const tz = String(payload?.displaySettings?.serverTimezone || "").trim();
  return tz || "UTC";
}

function formatRuDateTime(dt, timeZone = "UTC") {
  try {
    return new Intl.DateTimeFormat("ru-RU", {
      timeZone,
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    }).format(dt);
  } catch (_) {
    return String(dt || "");
  }
}

function formatRuDate(dt, timeZone = "UTC") {
  try {
    return new Intl.DateTimeFormat("ru-RU", {
      timeZone,
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    }).format(dt);
  } catch (_) {
    return "";
  }
}

function composeReadStateValue(isRead, whenText = "—") {
  return `${isRead ? "Прочитано" : "Не прочитано"}\n${String(whenText || "—").trim() || "—"}`;
}

function getTaskReadStateParts(value) {
  const raw = String(value || "").trim();
  const lines = raw.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  const first = String(lines[0] || "").trim();
  const second = String(lines[1] || "—").trim() || "—";
  const isRead = first.toLowerCase().startsWith("прочитано");
  return {
    isRead,
    statusText: isRead ? "Прочитано" : "Не прочитано",
    whenText: second
  };
}

function isAssignedEmployeeReader(payload, row, chatId) {
  const employees = getEmployeesSection(payload);
  const clickChat = String(chatId || "").trim();
  const clickEmp = findEmployeeByChatId(employees, clickChat);
  if (!clickEmp) return false;
  if (String(clickEmp[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") return false;

  const clickName = String(clickEmp[EMPLOYEE_COLUMNS.fullName] || "").trim();
  const assignedNames = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  const responsibleName = String(row[TASK_COLUMNS.responsible] || "").trim();

  // Строгий путь: исполнитель или ответственный по задаче.
  if (assignedNames.some((name) => findEmployeeByFullName(employees, name) === clickEmp)) return true;
  if (responsibleName && findEmployeeByFullName(employees, responsibleName) === clickEmp) return true;

  // Фолбэк: если исполнитель не сопоставился со справочником (например, старые/шаблонные ФИО),
  // считаем прочтение по факту нажатия у подключённого сотрудника, чтобы не блокировать процесс.
  const assignedEmp = assignedNames.map((name) => findEmployeeByFullName(employees, name)).find(Boolean) || null;
  if (!assignedEmp) return true;

  // Если у исполнителя не заполнен/неверен chat_id, также разрешаем отметку текущему подключённому сотруднику.
  const assignedChat = String(assignedEmp[EMPLOYEE_COLUMNS.chatId] || "").trim();
  if (!assignedChat) return true;

  // Иначе только сам исполнитель.
  return assignedChat === clickChat;
}

/**
 * Экранирует строку для Telegram parse_mode=HTML.
 * Telegram HTML — упрощённый, достаточно &, < и >.
 */
function escapeTgHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function getTelegramChatIdFromPhoneBinding(payload, employeeRow) {
  const phoneKey = normalizePhoneBindingKey(employeeRow?.[EMPLOYEE_COLUMNS.phone]);
  if (!phoneKey) return "";
  const bindings = payload?.telegramPhoneChatBindings && typeof payload.telegramPhoneChatBindings === "object"
    ? payload.telegramPhoneChatBindings
    : null;
  return String(bindings?.[phoneKey] || "").trim();
}

function getTelegramUsernameLink(employeeRow) {
  const raw = String(employeeRow?.telegramUsername || employeeRow?.username || "").trim().replace(/^@+/, "");
  if (!/^[a-zA-Z0-9_]{5,32}$/.test(raw)) return "";
  return `https://t.me/${escapeTgHtml(raw)}`;
}

function getTelegramPhoneLink(employeeRow) {
  const digits = normalizePhoneForMatch(employeeRow?.[EMPLOYEE_COLUMNS.phone]);
  return digits ? `https://t.me/+${escapeTgHtml(digits)}` : "";
}

/**
 * Превращает одно ФИО в HTML-mention <a href="tg://user?id=CHAT_ID">Name</a>,
 * если для этого сотрудника известен chat_id. Если chat_id нет, пробуем username
 * и телефонную ссылку. Иначе — просто экранированный текст.
 *
 * Telegram-клиенты (iOS/Android/Desktop/Web) корректно показывают такие mention'ы
 * как «синий ник» и при клике открывают чат с пользователем.
 */
function buildEmployeeMentionHtml(name, payload) {
  const fio = String(name || "").trim();
  if (!fio) return "";
  const safe = escapeTgHtml(fio);
  if (!payload) return safe;
  const employees = getEmployeesSection(payload);
  const empRows = Array.isArray(employees?.rows) ? employees.rows : [];
  const want = fio.toLowerCase();
  const er = empRows.find((r) => String(r?.[EMPLOYEE_COLUMNS.fullName] || "").trim().toLowerCase() === want);
  if (!er) return safe;
  const chatId = String(er[EMPLOYEE_COLUMNS.chatId] || "").trim() || getTelegramChatIdFromPhoneBinding(payload, er);
  if (chatId) {
    // Для личных чатов Telegram chat_id совпадает с user id, поэтому ссылка открывает профиль сотрудника.
    return `<a href="tg://user?id=${escapeTgHtml(chatId)}">${safe}</a>`;
  }
  const usernameLink = getTelegramUsernameLink(er);
  if (usernameLink) return `<a href="${usernameLink}">${safe}</a>`;
  const phoneLink = getTelegramPhoneLink(er);
  if (phoneLink) return `<a href="${phoneLink}">${safe}</a>`;
  return safe;
}

/**
 * Multi-name поле через запятую — каждое имя независимо превращаем в mention.
 */
function buildEmployeeMentionsListHtml(rawValue, payload) {
  const raw = String(rawValue ?? "").trim();
  if (!raw) return "—";
  const names = raw.split(",").map((s) => s.trim()).filter(Boolean);
  if (!names.length) return "—";
  return names.map((n) => buildEmployeeMentionHtml(n, payload)).join(", ");
}

/**
 * Текст карточки задачи в HTML с mention-ссылками на ФИО.
 * Параметр payload опциональный — если не задан, ФИО рендерится обычным
 * (HTML-экранированным) текстом без ссылок (для тестов / fallback).
 *
 * ВАЖНО: вывод этой функции отправляется с parse_mode: "HTML" в Telegram API.
 */
function buildFullTaskMessage(row, payload = null) {
  const E = escapeTgHtml;
  const lines = [];
  const st = String(row[TASK_COLUMNS.status] || "").trim();
  const statusLine = st ? `${STATUS_EMOJI[st] || "⚪"} ${E(st)}` : "—";
  lines.push(`📝 Задача №${E(String(row[TASK_COLUMNS.number] || "").trim() || "—")}`);
  lines.push(`📄 Описание: ${E(String(row[TASK_COLUMNS.task] || "").trim() || "—")}`);
  lines.push(`🏢 Объект: ${E(String(row[TASK_COLUMNS.object] || "").trim() || "—")}`);
  lines.push(`📌 Статус: ${statusLine}`);
  lines.push(`⚡ Приоритет: ${E(String(row[TASK_COLUMNS.priority] || "").trim() || "—")}`);
  lines.push(`📅 Дата: ${E(String(row[TASK_COLUMNS.addedDate] || "").trim() || "—")}`);
  lines.push(`🏗 Фаза: ${E(String(row[TASK_COLUMNS.phase] || "").trim() || "—")}`);
  lines.push(`📂 Раздел: ${E(String(row[TASK_COLUMNS.phaseSection] || "").trim() || "—")}`);
  lines.push(`🗂 Подраздел: ${E(String(row[TASK_COLUMNS.phaseSubsection] || "").trim() || "—")}`);
  lines.push(`👤 Ответственный: ${buildEmployeeMentionsListHtml(row[TASK_COLUMNS.assignedResponsible], payload)}`);
  lines.push(`👤 Постановщик задачи: ${buildEmployeeMentionsListHtml(row[TASK_COLUMNS.responsible], payload)}`);
  lines.push(`⏳ Срок: ${E(String(row[TASK_COLUMNS.dueDate] || "").trim() || "—")}`);
  const delayReason = String(row[TASK_COLUMNS.delayReason] || "").trim();
  if (delayReason) lines.push(`🚧 Причина отставания: ${E(delayReason)}`);
  const reassignReason = String(row[TASK_COLUMNS.reassignReason] || "").trim();
  if (reassignReason) lines.push(`🔁 Причина переназначения: ${E(reassignReason)}`);
  const note = String(row[TASK_COLUMNS.note] || "").trim();
  if (note) lines.push(`💬 Комментарий: ${E(note)}`);
  return lines.join("\n");
}

function taskCaptionWithPlan(row, payload = null) {
  const base = buildFullTaskMessage(row, payload);
  const plan = String(row[TASK_COLUMNS.plan] || "").trim();
  if (!plan) return base;
  const compactPlan = plan.length > 800 ? `${plan.slice(0, 797)}...` : plan;
  return `${base}\n🧩 Комментарии сотрудника (Результат):\n${escapeTgHtml(compactPlan.replace(/\r\n?/g, "\n"))}`;
}

function ensureLastTaskStore(payload) {
  if (!payload.telegramLastTaskByChat || typeof payload.telegramLastTaskByChat !== "object") {
    payload.telegramLastTaskByChat = {};
  }
  return payload.telegramLastTaskByChat;
}

function setLastTaskContext(payload, chatId, taskId, promptMessageId = null) {
  const store = ensureLastTaskStore(payload);
  store[String(chatId)] = {
    taskId: String(taskId ?? "").trim(),
    promptMessageId: Number(promptMessageId) || null,
    at: Date.now()
  };
}

async function safeDeleteMessage(token, chatId, messageId) {
  const mid = Number(messageId) || 0;
  if (!mid) return;
  try {
    await tg(token, "deleteMessage", { chat_id: chatId, message_id: mid });
  } catch (_) {
    /* noop */
  }
}

function resolveTaskUpdateRecipientChatIds(payload, row, excludeChatId = "") {
  const employees = getEmployeesSection(payload);
  const empRows = Array.isArray(employees?.rows) ? employees.rows : [];
  const out = new Set();
  const exclude = String(excludeChatId || "").trim();
  const addChat = (chat) => {
    const c = String(chat || "").trim();
    if (!c) return;
    if (exclude && c === exclude) return;
    out.add(c);
  };

  const names = new Set();
  const assigned = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  const responsible = String(row[TASK_COLUMNS.responsible] || "").trim();
  assigned.forEach((name) => names.add(name));
  if (responsible) names.add(responsible);
  for (const er of empRows) {
    const fio = String(er[EMPLOYEE_COLUMNS.fullName] || "").trim();
    if (!fio || !names.has(fio)) continue;
    if (String(er[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") continue;
    addChat(er[EMPLOYEE_COLUMNS.chatId]);
  }

  const ds = payload.displaySettings || {};
  const dupIds = new Set(
    Array.isArray(ds.telegramGlobalDuplicateRecipientIds)
      ? ds.telegramGlobalDuplicateRecipientIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  for (const er of empRows) {
    const eid = String(er[EMPLOYEE_COLUMNS.id] ?? "").trim();
    if (!eid || !dupIds.has(eid)) continue;
    if (String(er[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") continue;
    addChat(er[EMPLOYEE_COLUMNS.chatId]);
  }
  return Array.from(out);
}

function getTaskActionChatIds(payload, row) {
  const employees = getEmployeesSection(payload);
  const out = new Set();
  const taskId = String(row?.[TASK_COLUMNS.number] || "").trim();
  const activeReassign = getActiveReassignForTask(payload, taskId);
  if (activeReassign) {
    const toName = normalizePersonName(activeReassign.to || activeReassign.toEmployeeName || "");
    const toEmp = toName ? findEmployeeByFullName(employees, toName) : null;
    const toChat = String(toEmp?.[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (toChat && String(toEmp?.[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен") {
      out.add(toChat);
      return out;
    }
  }
  const assignedNames = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  for (const assignedName of assignedNames) {
    const assignedEmp = assignedName ? findEmployeeByFullName(employees, assignedName) : null;
    const assignedChat = String(assignedEmp?.[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (assignedChat && String(assignedEmp?.[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен") {
      out.add(assignedChat);
    }
  }
  if (!out.size) {
    const responsibleName = String(row[TASK_COLUMNS.responsible] || "").trim();
    const responsibleEmp = responsibleName ? findEmployeeByFullName(employees, responsibleName) : null;
    const responsibleChat = String(responsibleEmp?.[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (responsibleChat && String(responsibleEmp?.[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен") {
      out.add(responsibleChat);
    }
  }
  return out;
}

function canChatUseTaskActions(payload, row, chatId) {
  return isTaskActionAllowedForChat(payload, row, chatId);
}

function mainKeyboardForChat(payload, taskNumber, row, chatId, appTimeZone = "UTC") {
  if (!canChatUseTaskActions(payload, row, chatId)) return [];
  return mainKeyboard(taskNumber, row, appTimeZone);
}

async function broadcastTaskCardUpdate(payload, token, row, reasonText, excludeChatId = "") {
  const taskId = String(row[TASK_COLUMNS.number] ?? "").trim();
  const chatIds = resolveTaskUpdateRecipientChatIds(payload, row, excludeChatId);
  const actionChats = getTaskActionChatIds(payload, row);
  const appTz = resolveAppTimeZone(payload);
  for (const cid of chatIds) {
    const scopedRow = buildTaskRowForChat(payload, row, cid);
    const body = {
      chat_id: cid,
      text: `${taskCaptionWithPlan(scopedRow, payload)}\n\n${String(reasonText || "").trim() || "Обновление по задаче."}`
    };
    if (actionChats.has(String(cid || "").trim())) {
      body.reply_markup = { inline_keyboard: mainKeyboard(taskId, scopedRow, appTz) };
    }
    await tg(token, "sendMessage", body);
  }
}

export async function handleTelegramWebhook(req, res, pool) {
  const secret = String(process.env.TELEGRAM_WEBHOOK_SECRET || "").trim();
  if (secret) {
    const got = String(req.headers["x-telegram-bot-api-secret-token"] || "");
    if (got !== secret) {
      return res.status(403).json({ ok: false });
    }
  }

  const update = req.body;
  if (!update || typeof update !== "object") {
    return res.status(400).json({ ok: false });
  }

  res.status(200).json({ ok: true });

  try {
    const payload = await loadPayload(pool);
    const token = resolveBotToken(payload);
    if (!token) {
      console.error("telegram webhook: нет токена (TELEGRAM_BOT_TOKEN или токен в настройках приложения после синхронизации)");
      return;
    }
    await processUpdate(update, pool, token);
  } catch (e) {
    console.error("telegram webhook", e);
  }
}

async function loadPayload(pool) {
  const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
  const raw = rows[0]?.payload;
  const payload = raw && typeof raw === "object" ? JSON.parse(JSON.stringify(raw)) : {};
  normalizeTasksSectionByColumns(payload);
  return payload;
}

function isBlankTaskCell(value) {
  const s = String(value ?? "").trim();
  return s === "" || s === "-" || s === "—" || s.toLowerCase() === "null" || s.toLowerCase() === "undefined";
}

function mergeTaskRowsWithCurrent(currentPayload, nextPayload) {
  const currentTasks = getTasksSection(currentPayload);
  const nextTasks = getTasksSection(nextPayload);
  if (!Array.isArray(currentTasks?.rows) || !Array.isArray(nextTasks?.rows)) return;

  const nextById = new Map();
  for (const row of nextTasks.rows) {
    const id = String(row?.[TASK_COLUMNS.number] || "").trim();
    if (!id) continue;
    nextById.set(id, row);
  }

  // Источник истины по составу задач — текущая БД.
  // Webhook обновляет только bot-managed поля, чтобы устаревший снимок
  // не откатывал новые/изменённые данные таблицы.
  const BOT_MANAGED_COLS = [
    TASK_COLUMNS.status,
    TASK_COLUMNS.plan,
    TASK_COLUMNS.mediaAfter,
    TASK_COLUMNS.closedDate,
    TASK_COLUMNS.readState,
    TASK_COLUMNS.lastSentAt,
    TASK_COLUMNS.delayReason,
    TASK_COLUMNS.reassignReason,
    TASK_COLUMNS.reassignType
  ];

  const mergedRows = currentTasks.rows.map((currentRow) => {
    const id = String(currentRow?.[TASK_COLUMNS.number] || "").trim();
    const nextRow = id ? nextById.get(id) : null;
    if (!Array.isArray(nextRow)) {
      return Array.isArray(currentRow) ? [...currentRow] : currentRow;
    }
    const merged = Array.isArray(currentRow) ? [...currentRow] : new Array(TASK_ROW_LENGTH).fill("");
    for (const col of BOT_MANAGED_COLS) {
      const nextVal = nextRow[col];
      const curVal = currentRow[col];
      if (!isBlankTaskCell(nextVal) || isBlankTaskCell(curVal)) {
        merged[col] = nextVal;
      }
    }
    return merged;
  });

  // Защита от race: если веб-клиент удалил задачу пока бот её обрабатывал
  // (подтверждение закрытия, комментарий, реассайн и т.п.) — добавляем
  // строку обратно из next. Иначе любая операция бота "вспоминает" задачу,
  // которая уже исчезла из БД, и тихо теряется при save. Так пропала задача
  // 169 после массового удаления + одновременного закрытия в боте.
  const currentIds = new Set(
    currentTasks.rows.map((r) => String(r?.[TASK_COLUMNS.number] || "").trim()).filter(Boolean)
  );
  for (const nextRow of nextTasks.rows) {
    const id = String(nextRow?.[TASK_COLUMNS.number] || "").trim();
    if (!id || currentIds.has(id)) continue;
    mergedRows.push(Array.isArray(nextRow) ? [...nextRow] : nextRow);
  }

  nextTasks.rows = mergedRows;
}

async function savePayload(pool, payload) {
  const currentPayload = await loadPayload(pool);
  const nextPayload = payload && typeof payload === "object"
    ? JSON.parse(JSON.stringify(payload))
    : payload;
  if (nextPayload && typeof nextPayload === "object") {
    normalizeTasksSectionByColumns(nextPayload);
    mergeTaskRowsWithCurrent(currentPayload, nextPayload);
  }
  // Бампим revision: web-клиент с устаревшим baseRev получит 409 на PUT /api/data
  // и пройдёт через 3-way merge, который сохранит наши изменения (например файл,
  // прикреплённый в боте) вместо того чтобы их затереть.
  const { rows: savedRows } = await pool.query(
    `INSERT INTO app_state (id, payload, updated_at, revision) VALUES (1, $1::jsonb, NOW(), 1)
     ON CONFLICT (id) DO UPDATE SET payload = EXCLUDED.payload, updated_at = NOW(), revision = app_state.revision + 1
     RETURNING revision`,
    [JSON.stringify(nextPayload)]
  );
  const newRev = Number(savedRows[0]?.revision) || 0;
  // Broadcast по WebSocket, чтобы веб-клиенты сразу подтянули свежий снимок
  // (включая taskAttachments из bot-загрузки) до того как успеют запушить
  // свой устаревший слепок.
  try {
    if (typeof globalThis.__broadcastStateChanged === "function") {
      globalThis.__broadcastStateChanged(newRev, { source: "bot" });
    }
  } catch (_) {
    /* noop */
  }
}

/** Приоритет: переменная окружения, иначе токен из настроек приложения (БД). */
export function resolveBotToken(payload) {
  const env = String(process.env.TELEGRAM_BOT_TOKEN || "").trim();
  if (env) return env;
  return String(payload?.displaySettings?.telegramBotToken || "").trim();
}

/**
 * Регистрация webhook на публичный URL приложения (вызывается из POST /api/telegram/set-webhook).
 */
export async function configureTelegramWebhook(pool, baseUrl) {
  const normalizedBase = String(baseUrl || "")
    .trim()
    .replace(/\/$/, "");
  if (!normalizedBase) {
    return {
      ok: false,
      error:
        "Не задан публичный URL приложения. Укажите PUBLIC_APP_URL в переменных сервера или откройте запрос с вашего HTTPS-домена."
    };
  }
  const payload = await loadPayload(pool);
  const token = resolveBotToken(payload);
  if (!token) {
    return {
      ok: false,
      error:
        "Нет токена бота: сохраните токен в «Прочие настройки» → Telegram и дождитесь синхронизации с сервером (или задайте TELEGRAM_BOT_TOKEN)."
    };
  }
  const hookPath = "/api/telegram/webhook";
  const url = `${normalizedBase}${hookPath}`;
  const secret = String(process.env.TELEGRAM_WEBHOOK_SECRET || "").trim();
  const body = { url };
  if (secret) {
    body.secret_token = secret;
  }
  const r = await fetch(`https://api.telegram.org/bot${token}/setWebhook`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const j = await r.json();
  if (!j.ok) {
    return { ok: false, error: j.description || "setWebhook failed", description: j.description };
  }
  let botUsername = "";
  let botDisplayName = "";
  try {
    const mr = await fetch(`https://api.telegram.org/bot${token}/getMe`);
    const mj = await mr.json();
    if (mj.ok && mj.result && typeof mj.result === "object") {
      if (mj.result.username) botUsername = String(mj.result.username);
      const fn = String(mj.result.first_name || "").trim();
      const ln = String(mj.result.last_name || "").trim();
      botDisplayName = [fn, ln].filter(Boolean).join(" ").trim();
    }
  } catch (_) {
    /* noop */
  }
  return { ok: true, webhookUrl: url, botUsername, botDisplayName };
}

function normalizePersonName(s) {
  return String(s ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function nameTokens(name) {
  return normalizePersonName(name)
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean);
}

function scoreNameMatch(fio, displayName) {
  const a = nameTokens(fio);
  const b = nameTokens(displayName);
  if (!a.length || !b.length) return 0;
  const setA = new Set(a);
  let hit = 0;
  for (const t of b) {
    if (setA.has(t)) hit += 1;
  }
  return hit;
}

function telegramDisplayName(from) {
  if (!from || typeof from !== "object") return "";
  return normalizePersonName(`${from.first_name || ""} ${from.last_name || ""}`.trim());
}

function normalizePhoneForMatch(raw) {
  let digits = String(raw || "").replace(/[^\d+]/g, "");
  if (!digits) return "";
  if (digits.startsWith("00")) digits = `+${digits.slice(2)}`;
  if (digits.startsWith("+")) digits = digits.slice(1);
  digits = digits.replace(/\D/g, "");
  if (!digits) return "";
  // Совместимость со старым локальным вводом UZ без кода страны.
  if (digits.length === 9) return `998${digits}`;
  return digits;
}

function normalizePhoneBindingKey(raw) {
  const digits = normalizePhoneForMatch(raw);
  return digits ? `+${digits}` : "";
}

function findEmployeeByPhone(employees, phoneRaw) {
  const rows = employees?.rows || [];
  const phone = normalizePhoneForMatch(phoneRaw);
  if (!phone) return null;
  const hits = rows.filter((row) => normalizePhoneForMatch(row[EMPLOYEE_COLUMNS.phone]) === phone);
  if (hits.length === 1) return hits[0];
  return null;
}

function findEmployeeByPhoneBindingKey(employees, phoneKeyRaw) {
  const want = normalizePhoneBindingKey(phoneKeyRaw);
  if (!want) return null;
  const rows = employees?.rows || [];
  const hits = rows.filter((row) => normalizePhoneBindingKey(row?.[EMPLOYEE_COLUMNS.phone]) === want);
  return hits.length === 1 ? hits[0] : null;
}

function findEmployeeByStartParam(employees, startParam) {
  const raw = String(startParam || "").trim();
  if (!raw) return null;
  const rows = employees?.rows || [];
  const byPhone = findEmployeeByPhone(employees, raw);
  if (byPhone) return byPhone;
  let id = "";
  const m1 = raw.match(/^e[_-]?(\d+)$/i);
  const m2 = raw.match(/^id[_-]?(\d+)$/i);
  const m3 = raw.match(/^p(?:hone)?[_-]?(.+)$/i);
  if (m1) id = m1[1];
  else if (m2) id = m2[1];
  else if (m3) {
    const p = findEmployeeByPhone(employees, m3[1]);
    if (p) return p;
  }
  else if (/^\d+$/.test(raw)) id = raw;
  if (!id) return null;
  return rows.find((row) => String(row[EMPLOYEE_COLUMNS.id] ?? "").trim() === id) || null;
}

function findEmployeeByTelegramName(employees, from) {
  const rows = employees?.rows || [];
  if (!rows.length) return null;
  /** Один сотрудник в справочнике — привязка по /start без совпадения имён в Telegram. */
  if (rows.length === 1) {
    return rows[0];
  }

  const display = telegramDisplayName(from);
  if (!display) return null;

  const scores = rows.map((row) => {
    const fio = String(row[EMPLOYEE_COLUMNS.fullName] || "");
    return { row, sc: scoreNameMatch(fio, display) };
  });
  let bestScore = 0;
  for (const x of scores) {
    if (x.sc > bestScore) bestScore = x.sc;
  }
  if (bestScore >= 2) {
    const top = scores.filter((x) => x.sc === bestScore);
    if (top.length === 1) return top[0].row;
    return null;
  }
  /** Одно общее слово (например только имя в Telegram) и один явный кандидат. */
  if (bestScore >= 1 && nameTokens(display).length >= 1) {
    const top = scores.filter((x) => x.sc === bestScore);
    if (top.length === 1) return top[0].row;
  }
  return null;
}

function clearChatIdFromOtherEmployees(employees, chatIdStr, exceptRow) {
  for (const row of employees.rows) {
    if (row === exceptRow) continue;
    if (String(row[EMPLOYEE_COLUMNS.chatId] ?? "").trim() === chatIdStr) {
      row[EMPLOYEE_COLUMNS.chatId] = "";
    }
  }
}

async function bindEmployeeToChat({ employees, employee, chatId, from, pool, payload, token }) {
  const myChat = String(chatId);
  clearChatIdFromOtherEmployees(employees, myChat, employee);
  employee[EMPLOYEE_COLUMNS.chatId] = myChat;
  employee[EMPLOYEE_COLUMNS.telegram] = "Подключен";
  employee[EMPLOYEE_COLUMNS.activity] = "Активен";
  if (!payload.telegramPhoneChatBindings || typeof payload.telegramPhoneChatBindings !== "object") {
    payload.telegramPhoneChatBindings = {};
  }
  const phoneKey = normalizePhoneBindingKey(employee[EMPLOYEE_COLUMNS.phone]);
  if (phoneKey) {
    payload.telegramPhoneChatBindings[phoneKey] = myChat;
  }
  await savePayload(pool, payload);

  const name = String(employee[EMPLOYEE_COLUMNS.fullName] || "").trim() || "Сотрудник";
  const first = String(from?.first_name || "").trim();
  await tg(token, "sendMessage", {
    chat_id: chatId,
    text: `Здравствуйте${first ? `, ${first}` : ""}!\n\nВы подключены к боту как «${name}». Ваш Telegram ID сохранён в системе — уведомления по задачам будут приходить сюда.`,
    reply_markup: { remove_keyboard: true }
  });
  await tg(token, "sendMessage", {
    chat_id: chatId,
    text: "Доступны быстрые кнопки: «Мои задачи» и «Просроченные задачи».",
    reply_markup: quickUserKeyboard()
  });
}

async function restoreEmployeeBindingFromPhoneMap({ payload, employees, chatId, pool }) {
  const chatKey = String(chatId || "").trim();
  if (!chatKey || !payload || !employees?.rows?.length) return null;
  const map = payload.telegramPhoneChatBindings && typeof payload.telegramPhoneChatBindings === "object"
    ? payload.telegramPhoneChatBindings
    : null;
  if (!map) return null;
  let matchedPhoneKey = "";
  for (const [phoneKey, mappedChat] of Object.entries(map)) {
    if (String(mappedChat || "").trim() !== chatKey) continue;
    matchedPhoneKey = String(phoneKey || "").trim();
    if (matchedPhoneKey) break;
  }
  if (!matchedPhoneKey) return null;

  const employee = findEmployeeByPhoneBindingKey(employees, matchedPhoneKey);
  if (!employee) return null;
  if (String(employee[EMPLOYEE_COLUMNS.chatId] || "").trim() === chatKey
    && String(employee[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен") {
    return employee;
  }

  clearChatIdFromOtherEmployees(employees, chatKey, employee);
  employee[EMPLOYEE_COLUMNS.chatId] = chatKey;
  employee[EMPLOYEE_COLUMNS.telegram] = "Подключен";
  employee[EMPLOYEE_COLUMNS.activity] = "Активен";
  await savePayload(pool, payload);
  return employee;
}

function contactShareKeyboard() {
  return {
    keyboard: [[{ text: "Поделиться контактом", request_contact: true }]],
    resize_keyboard: true,
    one_time_keyboard: true
  };
}

async function handleTelegramStart(msg, pool, token) {
  const chatId = msg.chat?.id;
  const from = msg.from;
  if (chatId == null || !from) return;

  const text = String(msg.text || "").trim();
  const startParam = /^\/start(?:@\w+)?(?:\s+(.+))?$/i.exec(text)?.[1]?.trim() || "";

  const payload = await loadPayload(pool);
  const employees = getEmployeesSection(payload);
  if (!employees?.rows?.length) {
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: "Справочник сотрудников в системе пуст. Обратитесь к администратору."
    });
    return;
  }

  let emp = findEmployeeByStartParam(employees, startParam);
  if (!emp && msg.contact?.phone_number) {
    emp = findEmployeeByPhone(employees, msg.contact.phone_number);
  }
  if (!emp) {
    emp = findEmployeeByTelegramName(employees, from);
  }

  if (!emp) {
    const msgNo =
      startParam.length > 0
        ? "Не найден сотрудник по коду из ссылки. Проверьте ID или попросите у администратора новую ссылку «Подключить бота»."
        : "Не удалось сопоставить профиль автоматически. Отправьте свой номер через кнопку «Поделиться номером» (сравнение по телефону, +/без + поддерживается) или используйте ссылку с параметром e_<ID>.";
    await tg(token, "sendMessage", { chat_id: chatId, text: msgNo });
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: "Нажмите кнопку ниже и отправьте свой номер телефона:",
      reply_markup: contactShareKeyboard()
    });
    return;
  }

  await bindEmployeeToChat({ employees, employee: emp, chatId, from, pool, payload, token });
}

async function processUpdate(update, pool, token) {
  if (update.callback_query) {
    await handleCallback(update.callback_query, pool, token);
    return;
  }
  if (update.message) {
    await handleMessage(update.message, pool, token);
  }
}

async function handleCallback(q, pool, token) {
  const cq = q.data;
  const from = q.from;
  const chatId = q.message?.chat?.id;
  const messageId = q.message?.message_id;
  if (!cq || chatId == null || messageId == null) {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
    return;
  }

  const myCb = parseMyTasksCallbackData(cq);
  if (myCb) {
    const payload = await loadPayload(pool);
    const appTz = resolveAppTimeZone(payload);
    const myRows = getMyTasksForChat(payload, chatId);
    const objects = getMyTaskObjects(myRows);
    if (myCb.action === "bo") {
      if (!objects.length) {
        await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: messageId,
          text: "У вас пока нет доступных задач.",
          reply_markup: { inline_keyboard: [] }
        });
        await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
        return;
      }
      const objKeyboard = {
        inline_keyboard: objects.slice(0, 50).map((name, idx) => [
          { text: name, callback_data: myTasksCb("ob", String(idx)) }
        ])
      };
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: "Мои задачи\n\nВыберите объект:",
        reply_markup: objKeyboard
      });
      await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
      return;
    }
    if (myCb.action === "ob") {
      const idx = Number(myCb.args?.[0]);
      if (!Number.isFinite(idx) || idx < 0 || idx >= objects.length) {
        await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Объект не найден" });
        return;
      }
      const objectName = objects[idx];
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Объект: ${objectName}\n\nВыберите статус:`,
        reply_markup: getStatusPickKeyboardForObject(String(idx), myRows, objectName, appTz)
      });
      await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
      return;
    }
    if (myCb.action === "st") {
      const objIdx = Number(myCb.args?.[0]);
      const statusCode = String(myCb.args?.[1] || "all").trim();
      if (!Number.isFinite(objIdx) || objIdx < 0 || objIdx >= objects.length) {
        await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Объект не найден" });
        return;
      }
      const objectName = objects[objIdx];
      const filtered = filterMyTasksBySelection(myRows, objectName, statusCode, appTz);
      if (!filtered.length) {
        await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: messageId,
          text: `Объект: ${objectName}\nСтатус: ${myTaskStatusLabel(statusCode)}\n\nЗадач не найдено.`,
          reply_markup: {
            inline_keyboard: [
              [{ text: "⬅️ К статусам", callback_data: myTasksCb("ob", String(objIdx)) }],
              [{ text: "⬅️ К объектам", callback_data: myTasksCb("bo", "0") }]
            ]
          }
        });
        await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
        return;
      }
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Объект: ${objectName}\nСтатус: ${myTaskStatusLabel(statusCode)}\n\nВыберите задачу из списка:`,
        reply_markup: myTasksListKeyboard(filtered, payload, chatId)
      });
      await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
      return;
    }
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Некорректное действие" });
    return;
  }

  const overdueCb = parseOverdueCallbackData(cq);
  if (overdueCb) {
    const payload = await loadPayload(pool);
    const appTz = resolveAppTimeZone(payload);
    const overdueRows = getOverdueRowsForChat(payload, chatId, appTz);
    if (overdueCb.action === "bk") {
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: buildOverdueSummaryText(overdueRows.length),
        reply_markup: overdueSummaryKeyboard(overdueRows.length)
      });
      await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
      return;
    }
    if (overdueCb.action === "ls") {
      if (!overdueRows.length) {
        await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: messageId,
          text: "✅ Сейчас у вас нет просроченных задач.",
          reply_markup: { inline_keyboard: [] }
        });
        await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
        return;
      }
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: buildOverdueSummaryText(overdueRows.length),
        reply_markup: overdueListKeyboard(payload, overdueRows)
      });
      await tg(token, "answerCallbackQuery", { callback_query_id: q.id });
      return;
    }
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Некорректное действие" });
    return;
  }

  const parsed = parseCallbackData(cq);
  if (!parsed) {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Некорректные данные" });
    return;
  }

  const payload = await loadPayload(pool);
  const tasks = getTasksSection(payload);
  const employees = getEmployeesSection(payload);
  const requestedTaskId = String(parsed.taskNum || "").trim();
  let row = findTaskRow(tasks, requestedTaskId);
  if (!row && requestedTaskId.includes("/")) {
    row = findTaskRow(tasks, String(requestedTaskId).split("/")[0]);
  }
  if (!row) {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Задача не найдена" });
    return;
  }
  syncTaskMultiStateForRow(payload, row);

  const baseTaskId = String(row[TASK_COLUMNS.number] ?? "").trim();
  const taskId = requestedTaskId || baseTaskId;
  const emp = findEmployeeByChatId(employees, String(chatId));
  const empName = emp ? String(emp[EMPLOYEE_COLUMNS.fullName] || "").trim() : `chat ${chatId}`;
  const appTz = resolveAppTimeZone(payload);
  const viewerRow = buildTaskRowForChat(payload, row, chatId, taskId);
  const viewerAssigneeName = getTaskAssigneeNameByChat(payload, viewerRow, chatId);

  const answerOk = async (text) => {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: text || "", show_alert: Boolean(text && text.length > 200) });
  };

  if (parsed.action === "rd") {
    const nowText = formatRuDateTime(new Date(), appTz);
    const activeForChat = getActiveReassignForChat(payload, row, chatId, taskId);
    if (activeForChat) {
      activeForChat.readAt = nowText;
      activeForChat.updatedAt = nowText;
    } else {
    // Кнопка «📖 Прочитать» приходит только адресату сообщения по задаче,
    // поэтому отмечаем ознакомление сразу по факту нажатия.
    row[TASK_COLUMNS.readState] = composeReadStateValue(true, nowText);
    const assigneeName = getTaskAssigneeNameByChat(payload, viewerRow, chatId);
    if (assigneeName) {
      const state = getTaskMultiStateForRow(payload, row, { create: true });
      if (state && state[assigneeName]) {
        state[assigneeName].readAt = nowText;
        state[assigneeName].updatedAt = nowText;
      }
    }
    }
    appendTaskHistory(payload, taskId, empName, `Telegram: задача прочитана (${nowText})`);
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    const refreshedViewerRow = buildTaskRowForChat(payload, row, chatId, taskId);
    const inlineKeyboard = activeForChat && canChatUseTaskActions(payload, refreshedViewerRow, chatId)
      ? mainKeyboard(taskId, refreshedViewerRow, appTz)
      : mainKeyboardForChat(payload, taskId, refreshedViewerRow, chatId, appTz);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${buildFullTaskMessage(refreshedViewerRow, payload)}\n\nВыберите действие по задаче:`,
      reply_markup: { inline_keyboard: inlineKeyboard }
    });
    await answerOk("Задача отмечена как прочитанная");
    return;
  }

  if (parsed.action === "sm") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Смена статуса доступна только исполнителю");
      return;
    }
    setLastTaskContext(payload, chatId, taskId, messageId);
    const currentStatus = normalizeTaskStatusValue(String(viewerRow[TASK_COLUMNS.status] || "").trim());
    const keyboard = EMPLOYEE_STATUS_OPTIONS
      .map((label, i) => ({ label, i }))
      .filter((item) => item.label !== currentStatus)
      .map((item) => [{ text: statusLabelWithEmoji(item.label), callback_data: cb(taskId, `ss|${item.i}`) }]);
    if (!keyboard.length) {
      await answerOk("Статус уже установлен");
      return;
    }
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "bk") }]);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nВыберите новый статус:`,
      reply_markup: { inline_keyboard: keyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ss") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Смена статуса доступна только исполнителю");
      return;
    }
    const idx = Number(parsed.rest[0]);
    if (!Number.isFinite(idx) || idx < 0 || idx >= EMPLOYEE_STATUS_OPTIONS.length) {
      await answerOk("Неверный статус");
      return;
    }
    const newStatus = EMPLOYEE_STATUS_OPTIONS[idx];
    const isClose = newStatus === "Закрыт";
    const assigneeName = viewerAssigneeName;
    const activeReassign = getActiveReassignForTask(payload, taskId);

    if (isClose) {
      if (isTaskOverdueWithoutDelayReason(viewerRow, appTz)) {
        await answerOk("Сначала укажите причину отставания");
        await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: messageId,
          text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nЗадача просрочена. Перед закрытием укажите причину отставания через кнопку «🚧 Причина отставания».`,
          reply_markup: mainKeyboardForChat(payload, taskId, viewerRow, chatId, appTz)
        });
        return;
      }
      // Перед созданием запроса на закрытие — даём шанс прикрепить файл-обоснование.
      // Сохраняем контекст в сессии и просим прислать документ/фото/видео и т.п.
      // Поддерживаем несколько файлов: после каждого — обновляем prompt и ждём ещё.
      if (!payload.telegramSessions) payload.telegramSessions = {};
      payload.telegramSessions[String(chatId)] = {
        expect: "closeTaskFile",
        taskId,
        baseTaskId,
        assigneeName: assigneeName || "",
        promptMessageId: Number(messageId) || null,
        attachedNames: []
      };
      setLastTaskContext(payload, chatId, taskId, messageId);
      await savePayload(pool, payload);
      // Кнопки одним рядом: Назад слева, Пропустить справа.
      const closeFileKeyboard = [[
        { text: "⬅️ Назад", callback_data: cb(taskId, "bk") },
        { text: "⏭ Пропустить", callback_data: cb(taskId, "csk") }
      ]];
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nПрикрепите файлы-обоснования закрытия (документ, фото, видео и т.п.).\nМожно отправить несколько файлов подряд — каждое сообщение добавит файл в список.\nЛимит Telegram-бота: до 20 МБ на файл.\n\nКогда закончите — нажмите «Готово». Если документа нет — нажмите «Пропустить».`,
        reply_markup: { inline_keyboard: closeFileKeyboard }
      });
      await answerOk();
      return;
    }

    const oldStatus = activeReassign
      ? String(activeReassign.currentStatus || "В процессе").trim()
      : String(row[TASK_COLUMNS.status] ?? "").trim();
    if (activeReassign) {
      activeReassign.currentStatus = newStatus;
      appendTaskHistory(payload, taskId, empName, `Telegram: статус переназначенной задачи «${oldStatus || "—"}» → «${newStatus}»`);
      setLastTaskContext(payload, chatId, taskId, messageId);
      await savePayload(pool, payload);
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${taskCaptionWithPlan(buildTaskRowForChat(payload, row, chatId, taskId), payload)}\n\nСтатус обновлён.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, buildTaskRowForChat(payload, row, chatId, taskId), chatId, appTz) }
      });
      await answerOk();
      return;
    }
    if (assigneeName) {
      const nowText = formatRuDateTime(new Date(), appTz);
      const state = getTaskMultiStateForRow(payload, row, { create: true });
      if (state && state[assigneeName]) {
        state[assigneeName].status = newStatus;
        state[assigneeName].updatedAt = nowText;
        if (newStatus === "Закрыт") state[assigneeName].closedAt = nowText;
      }
      // Для задачи с одним исполнителем агрегатор multi-state не меняет row.status,
      // поэтому синхронизируем напрямую, чтобы Telegram/UI сразу видели новый статус.
      const assignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
      if (assignees.length <= 1) {
        row[TASK_COLUMNS.status] = newStatus;
        if (newStatus !== "Закрыт") {
          row[TASK_COLUMNS.closedDate] = "";
          setSingleTaskClosedAt(payload, row, "");
        } else {
          if (!String(row[TASK_COLUMNS.closedDate] || "").trim()) {
            row[TASK_COLUMNS.closedDate] = formatRuDate(new Date(), appTz);
          }
          setSingleTaskClosedAt(payload, row, nowText);
        }
      } else {
        updateTaskAggregateStatusFromMulti(payload, row, appTz);
      }
    } else {
      row[TASK_COLUMNS.status] = newStatus;
      if (newStatus !== "Закрыт") {
        setSingleTaskClosedAt(payload, row, "");
      }
    }
    appendTaskHistory(payload, taskId, empName, `Telegram: статус «${oldStatus || "—"}» → «${newStatus}»`);
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);

    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(buildTaskRowForChat(payload, row, chatId, taskId), payload)}\n\nСтатус обновлён.`,
      reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, buildTaskRowForChat(payload, row, chatId, taskId), chatId, appTz) }
    });
    await answerOk();
    return;
  }

  // Кнопка "Пропустить" / "Готово" на шаге прикрепления файла-обоснования
  // при закрытии задачи. Логика одинаковая — финализируем close-request.
  // Разница только в показанной надписи (Пропустить = 0 файлов, Готово = N файлов),
  // обработчик любой из них доводит до submitTaskCloseRequest.
  if (parsed.action === "csk" || parsed.action === "cdn") {
    const sess = payload.telegramSessions?.[String(chatId)];
    if (!sess || sess.expect !== "closeTaskFile" || String(sess.taskId || "").trim() !== taskId) {
      await answerOk("Сессия истекла, начните заново");
      return;
    }
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Действие доступно только исполнителю");
      return;
    }
    const viewerAssigneeName = getTaskAssigneeNameByChat(payload, viewerRow, chatId);
    const baseTaskIdForClose = String(sess.baseTaskId || taskId || "").trim();
    const assigneeNameForClose = String(sess.assigneeName || viewerAssigneeName || "").trim();
    const promptMessageIdForClose = Number(sess.promptMessageId) || Number(messageId) || null;
    clearSession(payload, String(chatId));
    await savePayload(pool, payload);
    await submitTaskCloseRequest({
      pool, token, payload, chatId, taskId, row, viewerRow, empName,
      baseTaskId: baseTaskIdForClose,
      assigneeName: assigneeNameForClose,
      messageId: promptMessageIdForClose,
      appTz
    });
    await answerOk();
    return;
  }

  if (parsed.action === "cm") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Комментарии доступны только исполнителю");
      return;
    }
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = { expect: "comment", taskId, promptMessageId: Number(messageId) || null };
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nНапишите комментарий одним сообщением ниже (или /отмена).`,
      reply_markup: { inline_keyboard: backOnlyKeyboard(taskId) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "dr") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Причина доступна только исполнителю");
      return;
    }
    if (!isDelayReasonAllowedNow(viewerRow, appTz)) {
      await answerOk("Причина отставания доступна только после просрочки планового срока");
      return;
    }
    // Шаг 1: выбор фактора. Дальше уже выбор причины (drf → drs) или
    // подтверждение принятия на себя (drf|self).
    const keyboard = [
      [{ text: "🌐 Внешний фактор", callback_data: cb(taskId, "drf|ext") }],
      [{ text: "🏢 Внутренний фактор (внутри компании)", callback_data: cb(taskId, "drf|int") }],
      [{ text: "✋ Принять на себя", callback_data: cb(taskId, "drf|self") }],
      [{ text: "⬅️ Назад", callback_data: cb(taskId, "bk") }]
    ];
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nКакой фактор привёл к отставанию?\n• Внешний фактор — обстоятельства вне компании, сохраняем причину.\n• Внутренний фактор — внутри компании, далее предложим делегировать задачу другому сотруднику.\n• Принять на себя — переназначение не выполняется, ответственность остаётся.`,
      reply_markup: { inline_keyboard: keyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "drf") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Действие доступно только исполнителю");
      return;
    }
    if (!isDelayReasonAllowedNow(viewerRow, appTz)) {
      await answerOk("Причина отставания доступна только после просрочки планового срока");
      return;
    }
    const factor = String(parsed.rest?.[0] || "").trim().toLowerCase();
    if (factor === "self") {
      // Принять на себя — никакого переназначения, просто фиксируем в истории.
      appendTaskHistory(payload, taskId, empName, "Telegram: причина отставания — принял ответственность на себя (без переназначения)");
      setLastTaskContext(payload, chatId, taskId, messageId);
      await savePayload(pool, payload);
      const refreshed = buildTaskRowForChat(payload, row, chatId, taskId);
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${taskCaptionWithPlan(refreshed, payload)}\n\nВы приняли ответственность за отставание. Переназначение не выполнено.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, refreshed, chatId, appTz) }
      });
      await answerOk();
      return;
    }
    if (factor !== "ext" && factor !== "int") {
      await answerOk("Некорректный выбор");
      return;
    }
    // Запоминаем фактор в сессии, чтобы в drs|i ветвление пошло корректно.
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = {
      expect: "delayReasonPick",
      factor,
      taskId,
      promptMessageId: Number(messageId) || null
    };
    const options = getDelayReasonOptions(payload);
    const keyboard = options.map((reason, i) => [{ text: reason, callback_data: cb(taskId, `drs|${i}`) }]);
    if (factor === "ext") {
      // Для внешнего фактора разрешаем ввести свой вариант.
      keyboard.push([{ text: "✍️ Другое", callback_data: cb(taskId, "dro") }]);
    }
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "dr") }]);
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nВыберите причину отставания${factor === "int" ? " (для делегирования)" : ""}:`,
      reply_markup: { inline_keyboard: keyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ra") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Переназначение доступно только исполнителю");
      return;
    }
    const keyboard = [
      [{ text: "Ошибочная задача", callback_data: cb(taskId, "rat|mistake") }],
      [{ text: "Делегирование задачи", callback_data: cb(taskId, "rat|delegation") }],
      [{ text: "⬅️ Назад", callback_data: cb(taskId, "bk") }]
    ];
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nВыберите тип переназначения:`,
      reply_markup: { inline_keyboard: keyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "rat") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Недостаточно прав");
      return;
    }
    const rawType = String(parsed.rest?.[0] || "").trim().toLowerCase();
    // Поддерживаем как новые типы (mistake/delegation), так и старые из live-сообщений
    // отправленных до релиза (obj→mistake, subj→delegation).
    let reasonType = "";
    if (rawType === "mistake" || rawType === "obj") reasonType = "mistake";
    else if (rawType === "delegation" || rawType === "subj") reasonType = "delegation";
    if (!reasonType) {
      await answerOk("Некорректный выбор");
      return;
    }
    if (reasonType === "mistake") {
      if (hasPendingMistakeReassignRequest(payload, taskId)) {
        await answerOk("Уже отправлено администратору");
        return;
      }
      const requestId = `rr_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
      const reassignCode = getNextReassignCode(payload, taskId);
      const reqStore = ensureReassignRequestsStore(payload);
      const requestAllowed = new Set();
      getAlwaysConfirmChatIds(payload).forEach((cid) => requestAllowed.add(cid));
      const adminChat = String(payload?.displaySettings?.telegramAdminChatId || "").trim();
      if (adminChat) requestAllowed.add(adminChat);
      const currentAssignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
      const fromAssignee = String(viewerAssigneeName || currentAssignees[0] || "").trim();
      reqStore[requestId] = {
        id: requestId,
        code: reassignCode,
        taskId,
        status: "pending",
        reasonType: "mistake",
        reasonText: "Ожидает выбора нового ответственного администратором",
        departmentName: "",
        fromEmployeeName: fromAssignee,
        toEmployeeName: "",
        requesterChatId: String(chatId),
        requesterName: empName,
        requesterAssigneeName: String(viewerAssigneeName || "").trim(),
        createdAt: new Date().toISOString(),
        sourceMessageId: Number(messageId) || null,
        allowedConfirmChatIds: Array.from(requestAllowed),
        needsAdminAssignee: true
      };
      row[TASK_COLUMNS.status] = "Передано";
      row[TASK_COLUMNS.reassignReason] = "";
      row[TASK_COLUMNS.reassignType] = "Ошибочная задача";
      appendTaskHistory(payload, taskId, empName, `Telegram: задача отмечена как ошибочная, ожидает выбора нового ответственного администратором (${fromAssignee || "—"})`);
      clearSession(payload, String(chatId));
      await savePayload(pool, payload);

      const requestText =
        `${buildFullTaskMessage(row, payload)}\n\nОшибочная задача\n` +
        `Задача: ${escapeTgHtml(reassignCode)}\n` +
        `От кого: ${buildEmployeeMentionHtml(fromAssignee, payload) || "—"}\n` +
        `Нужно выбрать нового ответственного в веб-приложении.`;
      for (const cid of requestAllowed) {
        await tg(token, "sendMessage", { chat_id: cid, text: requestText });
      }
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${taskCaptionWithPlan(row, payload)}\n\nЗадача отправлена администратору как ошибочная. Новый ответственный будет выбран в веб-приложении.`
      });
      await answerOk("Отправлено администратору");
      return;
    }
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = {
      expect: "reassignDepartment",
      taskId,
      reasonType,
      reasonText: "",
      promptMessageId: Number(messageId) || null
    };
    const departments = buildDepartmentOptions(payload);
    const keyboard = departments.map((name, i) => [{ text: name, callback_data: cb(taskId, `rad|${i}`) }]);
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "ra") }]);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nШаг 1/2: выберите отдел`,
      reply_markup: { inline_keyboard: keyboard.slice(0, 80) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "rar") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Недостаточно прав");
      return;
    }
    const idx = Number(parsed.rest?.[0]);
    const options = getDelayReasonOptions(payload);
    if (!Number.isFinite(idx) || idx < 0 || idx >= options.length) {
      await answerOk("Неверный выбор");
      return;
    }
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = {
      expect: "reassignDepartment",
      taskId,
      reasonType: "objective",
      reasonText: String(options[idx] || "").trim(),
      promptMessageId: Number(messageId) || null
    };
    const departments = buildDepartmentOptions(payload);
    const keyboard = departments.map((name, i) => [{ text: name, callback_data: cb(taskId, `rad|${i}`) }]);
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "ra") }]);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nШаг 1/2: выберите отдел`,
      reply_markup: { inline_keyboard: keyboard.slice(0, 80) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "rad") {
    const sess = payload.telegramSessions?.[String(chatId)];
    if (!sess || sess.expect !== "reassignDepartment" || String(sess.taskId || "").trim() !== taskId) {
      await answerOk("Сессия истекла, начните заново");
      return;
    }
    const idx = Number(parsed.rest?.[0]);
    const departments = buildDepartmentOptions(payload);
    if (!Number.isFinite(idx) || idx < 0 || idx >= departments.length) {
      await answerOk("Отдел не найден");
      return;
    }
    const departmentName = String(departments[idx] || "").trim();
    const currentAssignees = new Set(parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]).map((x) => x.toLowerCase()));
    const employees = buildDepartmentEmployees(payload, departmentName).filter((e) => !currentAssignees.has(e.fullName.toLowerCase()));
    payload.telegramSessions[String(chatId)] = {
      ...sess,
      expect: "reassignEmployee",
      departmentName,
      candidateEmployeeNames: employees.map((x) => x.fullName)
    };
    const keyboard = employees.map((empRow, i) => [{ text: empRow.fullName, callback_data: cb(taskId, `rae|${i}`) }]);
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "ra") }]);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nШаг 2/2: выберите сотрудника отдела «${escapeTgHtml(departmentName)}»`,
      reply_markup: { inline_keyboard: keyboard.slice(0, 80) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "rae") {
    const sess = payload.telegramSessions?.[String(chatId)];
    if (!sess || sess.expect !== "reassignEmployee" || String(sess.taskId || "").trim() !== taskId) {
      await answerOk("Сессия истекла, начните заново");
      return;
    }
    const idx = Number(parsed.rest?.[0]);
    const departmentName = String(sess.departmentName || "").trim();
    const all = buildDepartmentEmployees(payload, departmentName);
    const candidates = Array.isArray(sess.candidateEmployeeNames)
      ? sess.candidateEmployeeNames.map((x) => String(x || "").trim()).filter(Boolean)
      : all.map((x) => x.fullName);
    const employees = all.filter((x) => candidates.includes(x.fullName));
    if (!Number.isFinite(idx) || idx < 0 || idx >= employees.length) {
      await answerOk("Сотрудник не найден");
      return;
    }
    const target = employees[idx];
    const requestId = `rr_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    const reassignCode = getNextReassignCode(payload, taskId);
    const reqStore = ensureReassignRequestsStore(payload);
    const requestAllowed = new Set();
    getDepartmentHeadChatIdsForTask(payload, row).forEach((cid) => requestAllowed.add(cid));
    getAlwaysConfirmChatIds(payload).forEach((cid) => requestAllowed.add(cid));
    const adminChat = String(payload?.displaySettings?.telegramAdminChatId || "").trim();
    if (adminChat) requestAllowed.add(adminChat);

    const currentAssignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
    const fromAssignee = String(viewerAssigneeName || currentAssignees[0] || "").trim();
    reqStore[requestId] = {
      id: requestId,
      code: reassignCode,
      taskId,
      status: "pending",
      reasonType: sess.reasonType || "objective",
      reasonText: String(sess.reasonText || "").trim(),
      departmentName,
      fromEmployeeName: fromAssignee,
      toEmployeeName: String(target.fullName || "").trim(),
      requesterChatId: String(chatId),
      requesterName: empName,
      requesterAssigneeName: String(viewerAssigneeName || "").trim(),
      createdAt: new Date().toISOString(),
      sourceMessageId: Number(messageId) || null,
      allowedConfirmChatIds: Array.from(requestAllowed)
    };
    appendTaskHistory(
      payload,
      taskId,
      empName,
      `Telegram: запрошено переназначение (${fromAssignee || "—"} → ${target.fullName || "—"})`
    );
    clearSession(payload, String(chatId));
    await savePayload(pool, payload);

    const reasonLabel = getReassignTypeLabel(reqStore[requestId].reasonType) || "—";
    const requestText =
      `${buildFullTaskMessage(row, payload)}\n\nЗапрос на переназначение\n` +
      `Задача: ${escapeTgHtml(reassignCode)}\n` +
      `Причина (${escapeTgHtml(reasonLabel)}): ${escapeTgHtml(reqStore[requestId].reasonText || "—")}\n` +
      `С кого: ${buildEmployeeMentionHtml(fromAssignee, payload) || "—"}\n` +
      `На кого: ${buildEmployeeMentionHtml(target.fullName, payload) || "—"}`;
    const reqKb = {
      inline_keyboard: [[
        { text: "Подтвердить", callback_data: cb(taskId, `ar|${requestId}|y`) },
        { text: "Отклонить", callback_data: cb(taskId, `ar|${requestId}|n`) }
      ]]
    };
    for (const cid of Array.from(requestAllowed)) {
      await tg(token, "sendMessage", { chat_id: cid, text: requestText, reply_markup: reqKb });
    }

    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(buildTaskRowForChat(payload, row, chatId, taskId), payload)}\n\nЗапрос ${reassignCode} отправлен на подтверждение.`,
      reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, buildTaskRowForChat(payload, row, chatId, taskId), chatId, appTz) }
    });
    await answerOk("Запрос отправлен");
    return;
  }

  if (parsed.action === "drs") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Причина доступна только исполнителю");
      return;
    }
    if (!isDelayReasonAllowedNow(viewerRow, appTz)) {
      await answerOk("Причина отставания доступна только после просрочки планового срока");
      return;
    }
    const idx = Number(parsed.rest[0]);
    const options = getDelayReasonOptions(payload);
    if (!Number.isFinite(idx) || idx < 0 || idx >= options.length) {
      await answerOk("Неверный выбор");
      return;
    }
    const sess = payload.telegramSessions?.[String(chatId)];
    const factor = String(sess?.factor || "").trim().toLowerCase();
    const prevReason = String(row[TASK_COLUMNS.delayReason] || "").trim();
    const nextReason = String(options[idx] || "").trim();
    row[TASK_COLUMNS.delayReason] = nextReason;

    if (factor === "int") {
      // Внутренний фактор: причина сохранена, переходим к делегированию.
      // Создаём reassign-сессию: будет dept → employee → confirm.
      if (!payload.telegramSessions) payload.telegramSessions = {};
      payload.telegramSessions[String(chatId)] = {
        expect: "reassignDepartment",
        taskId,
        reasonType: "delegation",
        reasonText: nextReason,
        promptMessageId: Number(messageId) || null,
        fromOverdueFactor: "int"
      };
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: причина отставания «${prevReason || "—"}» → «${nextReason}» (внутренний фактор, начато делегирование)`
      );
      setLastTaskContext(payload, chatId, taskId, messageId);
      const departments = buildDepartmentOptions(payload);
      const keyboard = departments.map((name, i) => [{ text: name, callback_data: cb(taskId, `rad|${i}`) }]);
      keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "dr") }]);
      await savePayload(pool, payload);
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${taskCaptionWithPlan(buildTaskRowForChat(payload, row, chatId, taskId), payload)}\n\nПричина сохранена. Шаг 1/2: выберите отдел для делегирования.`,
        reply_markup: { inline_keyboard: keyboard.slice(0, 80) }
      });
      await broadcastTaskCardUpdate(payload, token, row, "Причина отставания обновлена.", String(chatId));
      await answerOk();
      return;
    }

    // Внешний фактор (или legacy без сессии): сохраняем причину и закрываем шаг.
    clearSession(payload, String(chatId));
    appendTaskHistory(payload, taskId, empName, `Telegram: причина отставания «${prevReason || "—"}» → «${nextReason}»${factor === "ext" ? " (внешний фактор)" : ""}`);
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(buildTaskRowForChat(payload, row, chatId, taskId), payload)}\n\nПричина отставания сохранена.`,
      reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, buildTaskRowForChat(payload, row, chatId, taskId), chatId, appTz) }
    });
    await broadcastTaskCardUpdate(payload, token, row, "Причина отставания обновлена.", String(chatId));
    await answerOk();
    return;
  }

  if (parsed.action === "dro") {
    if (!canChatUseTaskActions(payload, viewerRow, chatId)) {
      await answerOk("Причина доступна только исполнителю");
      return;
    }
    if (!isDelayReasonAllowedNow(viewerRow, appTz)) {
      await answerOk("Причина отставания доступна только после просрочки планового срока");
      return;
    }
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = { expect: "delayReason", taskId, promptMessageId: Number(messageId) || null };
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nНапишите причину отставания одним сообщением (или /отмена).`,
      reply_markup: { inline_keyboard: backOnlyKeyboard(taskId) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "bk") {
    const activeForChat = getActiveReassignForChat(payload, row, chatId, taskId);
    const currentRead = getTaskReadStateParts(row[TASK_COLUMNS.readState]);
    const assigneeName = getTaskAssigneeNameByChat(payload, viewerRow, chatId);
    const state = assigneeName ? getTaskMultiStateForRow(payload, row, { create: true }) : null;
    const assigneeReadAt = assigneeName && state?.[assigneeName] ? String(state[assigneeName].readAt || "").trim() : "";
    const activeReadAt = activeForChat ? String(activeForChat.readAt || "").trim() : "";
    if (activeForChat && canChatUseTaskActions(payload, viewerRow, chatId) && !activeReadAt) {
      const nowText = formatRuDateTime(new Date(), appTz);
      activeForChat.readAt = nowText;
      activeForChat.updatedAt = nowText;
      appendTaskHistory(payload, taskId, empName, `Telegram: задача открыта/прочитана (${nowText})`);
    } else if (canChatUseTaskActions(payload, viewerRow, chatId) && (!currentRead.isRead || (assigneeName && !assigneeReadAt))) {
      const nowText = formatRuDateTime(new Date(), appTz);
      row[TASK_COLUMNS.readState] = composeReadStateValue(true, nowText);
      if (assigneeName && state?.[assigneeName]) {
        state[assigneeName].readAt = nowText;
        state[assigneeName].updatedAt = nowText;
      }
      appendTaskHistory(payload, taskId, empName, `Telegram: задача открыта/прочитана (${nowText})`);
    }
    clearSession(payload, String(chatId));
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    const refreshedViewerRow = buildTaskRowForChat(payload, row, chatId, taskId);
    const inlineKeyboard = activeForChat && canChatUseTaskActions(payload, refreshedViewerRow, chatId)
      ? mainKeyboard(taskId, refreshedViewerRow, appTz)
      : mainKeyboardForChat(payload, taskId, refreshedViewerRow, chatId, appTz);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(refreshedViewerRow, payload)}\n\nВыберите действие по задаче:`,
      reply_markup: { inline_keyboard: inlineKeyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ad") {
    const decision = parsed.rest[0];
    if (!["y", "n"].includes(decision)) {
      await answerOk();
      return;
    }
    const req = payload.telegramCloseRequests?.[taskId];
    if (!req) {
      await answerOk("Запрос уже обработан");
      return;
    }
    const allowed = Array.isArray(req.allowedConfirmChatIds)
      ? req.allowedConfirmChatIds.map((x) => String(x).trim()).filter(Boolean)
      : [];
    const canByRequest = allowed.includes(String(chatId));
    if (!canByRequest && !canConfirmCloseInPayload(payload, String(chatId))) {
      await answerOk("Недостаточно прав");
      return;
    }

    if (decision === "y") {
      const appTz = resolveAppTimeZone(payload);
      if (isTaskOverdueWithoutDelayReason(row, appTz)) {
        await answerOk("Сначала укажите причину отставания");
        return;
      }
      const now = new Date();
      const confirmedAt = formatRuDateTime(now, appTz);
      const requesterAssignee = normalizePersonName(req.requesterAssigneeName || "");
      const closeCode = String(req.reassignCode || taskId).trim();
      const closeBaseTaskId = String(req.baseTaskId || baseTaskId || "").trim() || String(closeCode.split("/")[0] || "").trim();
      const isReassignClose = closeCode.includes("/");
      const assignees = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
      if (isReassignClose) {
        const reassignLogStore = ensureTaskReassignLogStore(payload);
        if (!Array.isArray(reassignLogStore[closeBaseTaskId])) reassignLogStore[closeBaseTaskId] = [];
        const reassignEntry = reassignLogStore[closeBaseTaskId].find((item) => String(item?.code || "").trim() === closeCode);
        if (reassignEntry) {
          reassignEntry.currentStatus = "Закрыт";
          reassignEntry.closedAt = confirmedAt;
          reassignEntry.updatedAt = confirmedAt;
        }
        // Если закрыта переназначенная ветка, считаем родительскую задачу закрытой.
        row[TASK_COLUMNS.status] = "Закрыт";
        row[TASK_COLUMNS.closedDate] = formatRuDate(now, appTz);
        setSingleTaskClosedAt(payload, row, confirmedAt);
      } else if (requesterAssignee && assignees.length > 1) {
        const state = getTaskMultiStateForRow(payload, row, { create: true });
        if (state && state[requesterAssignee]) {
          state[requesterAssignee].status = "Закрыт";
          state[requesterAssignee].closedAt = confirmedAt;
          state[requesterAssignee].updatedAt = confirmedAt;
        }
        updateTaskAggregateStatusFromMulti(payload, row, appTz);
      } else {
        row[TASK_COLUMNS.status] = "Закрыт";
        row[TASK_COLUMNS.closedDate] = formatRuDate(now, appTz);
        setSingleTaskClosedAt(payload, row, confirmedAt);
      }
      const confirmInfo = requesterAssignee && assignees.length > 1
        ? `Закрытие подтверждено: ${escapeTgHtml(requesterAssignee)} (${escapeTgHtml(empName || "—")}), ${escapeTgHtml(confirmedAt)}`
        : `Закрытие подтверждено: ${escapeTgHtml(empName || "—")}, ${escapeTgHtml(confirmedAt)}`;
      delete payload.telegramCloseRequests[taskId];
      appendTaskHistory(payload, closeBaseTaskId || taskId, empName, `Telegram: закрытие задачи подтверждено (${closeCode || taskId}, ${confirmedAt})`);
      await savePayload(pool, payload);

      const sourceMid = Number(req.sourceMessageId) || 0;
      const requesterRow = buildTaskRowForChat(payload, row, req.chatId, closeCode || taskId);
      const requesterText = `${buildFullTaskMessage(requesterRow, payload)}\n\n✅ ${confirmInfo}`;
      if (sourceMid) {
        await tg(token, "editMessageText", {
          chat_id: req.chatId,
          message_id: sourceMid,
          text: requesterText,
          reply_markup: { inline_keyboard: mainKeyboardForChat(payload, closeCode || taskId, requesterRow, req.chatId, resolveAppTimeZone(payload)) }
        });
      } else {
        await tg(token, "sendMessage", {
          chat_id: req.chatId,
          text: requesterText,
          reply_markup: { inline_keyboard: mainKeyboardForChat(payload, closeCode || taskId, requesterRow, req.chatId, resolveAppTimeZone(payload)) }
        });
      }
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача №${taskId}: закрытие подтверждено. Исполнителю отправлено уведомление.`,
        reply_markup: { inline_keyboard: [] }
      });
      // Гасим prompt'ы у других админских чатов, кому ушло уведомление.
      await clearCloseConfirmPrompts(token, req, taskId, `закрытие подтверждено (${empName || "admin"})`, chatId);
      await broadcastTaskCardUpdate(payload, token, row, confirmInfo, req.chatId);
    } else {
      // Откат: возвращаем статус, который был до перехода в "Проверка".
      const prevStatus = String(req.previousStatus || "В процессе").trim() || "В процессе";
      if (String(row[TASK_COLUMNS.status] || "").trim() === "Проверка") {
        row[TASK_COLUMNS.status] = prevStatus;
      }
      delete payload.telegramCloseRequests[taskId];
      appendTaskHistory(payload, taskId, empName, `Telegram: отклонён запрос на закрытие, статус возвращён → «${prevStatus}»`);
      await savePayload(pool, payload);
      await tg(token, "sendMessage", {
        chat_id: req.chatId,
        text: `Задача №${taskId}: запрос на закрытие отклонён. Статус задачи возвращён в «${prevStatus}».`
      });
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача №${taskId}: закрытие отклонено. Исполнитель уведомлён.`,
        reply_markup: { inline_keyboard: [] }
      });
      // Гасим prompt'ы у других админских чатов.
      await clearCloseConfirmPrompts(token, req, taskId, `закрытие отклонено (${empName || "admin"})`, chatId);
    }
    await answerOk();
    return;
  }

  if (parsed.action === "ar") {
    const requestId = String(parsed.rest?.[0] || "").trim();
    const decision = String(parsed.rest?.[1] || "").trim().toLowerCase();
    if (!requestId || !["y", "n"].includes(decision)) {
      await answerOk("Некорректные данные");
      return;
    }
    const reqStore = ensureReassignRequestsStore(payload);
    const reqEntry = reqStore[requestId];
    if (!reqEntry) {
      await answerOk("Заявка не найдена");
      return;
    }
    if (String(reqEntry.status || "").trim() !== "pending") {
      await answerOk("Заявка уже обработана");
      return;
    }
    const allowed = Array.isArray(reqEntry.allowedConfirmChatIds)
      ? reqEntry.allowedConfirmChatIds.map((x) => String(x).trim()).filter(Boolean)
      : [];
    if (!allowed.includes(String(chatId).trim()) && !canConfirmCloseInPayload(payload, String(chatId))) {
      await answerOk("Недостаточно прав");
      return;
    }
    const rqTaskId = String(reqEntry.taskId || "").trim();
    const rqRow = findTaskRow(tasks, rqTaskId);
    if (!rqRow) {
      await answerOk("Задача не найдена");
      return;
    }
    const nowIso = new Date().toISOString();
    const nowText = formatRuDateTime(new Date(), resolveAppTimeZone(payload));
    const actor = empName || `chat ${chatId}`;
    const fromName = String(reqEntry.fromEmployeeName || "").trim();
    const toName = String(reqEntry.toEmployeeName || "").trim();
    const reassignCode = String(reqEntry.code || buildReassignChildTaskId(rqTaskId, 1)).trim();
    const logStore = ensureTaskReassignLogStore(payload);
    if (!Array.isArray(logStore[rqTaskId])) logStore[rqTaskId] = [];

    if (decision === "y") {
      reqEntry.status = "approved";
      reqEntry.decidedAt = nowIso;
      reqEntry.decidedBy = actor;
      rqRow[TASK_COLUMNS.status] = "Передано";
      rqRow[TASK_COLUMNS.reassignReason] = String(reqEntry.reasonText || "").trim();
      // Тип переназначения отображается в новом столбце "Тип переназначения".
      const _typeLabel = getReassignTypeLabel(reqEntry.reasonType);
      if (_typeLabel) rqRow[TASK_COLUMNS.reassignType] = _typeLabel;
      appendTaskHistory(payload, rqTaskId, actor, `Telegram: переназначение подтверждено (${fromName || "—"} → ${toName || "—"})${_typeLabel ? `, тип: ${_typeLabel}` : ""}`);
      const employeesRows = Array.isArray(getEmployeesSection(payload)?.rows) ? getEmployeesSection(payload).rows : [];
      const targetEmp = employeesRows.find((r) => String(r?.[EMPLOYEE_COLUMNS.fullName] || "").trim().toLowerCase() === toName.toLowerCase());
      const targetChat = String(targetEmp?.[EMPLOYEE_COLUMNS.chatId] || "").trim();
      const canSendToTarget = targetChat && String(targetEmp?.[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен";
      let reassignEntry = logStore[rqTaskId].find((item) => String(item?.code || "").trim() === reassignCode) || null;
      if (!reassignEntry) {
        reassignEntry = {
          id: requestId,
          code: reassignCode,
          status: reqEntry.status,
          currentStatus: "В процессе",
          comment: "",
          readAt: "",
          sentAt: "",
          from: fromName,
          to: toName,
          reasonType: reqEntry.reasonType || "",
          reasonText: reqEntry.reasonText || "",
          department: reqEntry.departmentName || "",
          requestedBy: reqEntry.requesterName || "",
          decidedBy: actor,
          createdAt: reqEntry.createdAt || nowIso,
          decidedAt: nowIso
        };
        logStore[rqTaskId].unshift(reassignEntry);
      } else {
        reassignEntry.status = reqEntry.status;
        reassignEntry.currentStatus = "В процессе";
        reassignEntry.from = fromName;
        reassignEntry.to = toName;
        reassignEntry.reasonType = reqEntry.reasonType || "";
        reassignEntry.reasonText = reqEntry.reasonText || "";
        reassignEntry.department = reqEntry.departmentName || "";
        reassignEntry.requestedBy = reqEntry.requesterName || "";
        reassignEntry.decidedBy = actor;
        reassignEntry.decidedAt = nowIso;
      }
      if (logStore[rqTaskId].length > 100) logStore[rqTaskId].length = 100;
      await savePayload(pool, payload);

      if (canSendToTarget) {
        const readKeyboard = { inline_keyboard: [[{ text: "📖 Прочитать", callback_data: cb(reassignCode, "rd") }]] };
        const targetRow = buildTaskRowForReassign(rqRow, reassignEntry, rqTaskId);
        await tg(token, "sendMessage", {
          chat_id: targetChat,
          text: `${buildFullTaskMessage(targetRow, payload)}\n\nВам назначена задача (${reassignCode}).`,
          reply_markup: readKeyboard
        });
        reassignEntry.sentAt = nowIso;
        await savePayload(pool, payload);
      }
      const requesterChat = String(reqEntry.requesterChatId || "").trim();
      const sourceMid = Number(reqEntry.sourceMessageId) || 0;
      const reasonLabel = getReassignTypeLabel(reqEntry.reasonType) || "—";
      const requesterText =
        `${buildFullTaskMessage(buildTaskRowForChat(payload, rqRow, requesterChat), payload)}\n\n` +
        `Вы переназначили задачу (${escapeTgHtml(reassignCode)}).\n` +
        `Причина (${escapeTgHtml(reasonLabel)}): ${escapeTgHtml(String(reqEntry.reasonText || "—").trim() || "—")}\n` +
        `Передано: ${buildEmployeeMentionHtml(toName, payload) || "—"}\n` +
        `Подтвердил: ${escapeTgHtml(actor)}\n` +
        `Дата/время: ${escapeTgHtml(nowText)}`;
      if (requesterChat) {
        if (sourceMid) {
          await tg(token, "editMessageText", {
            chat_id: requesterChat,
            message_id: sourceMid,
            text: requesterText,
            reply_markup: { inline_keyboard: [] }
          });
        } else {
          await tg(token, "sendMessage", { chat_id: requesterChat, text: requesterText });
        }
      }
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача ${reassignCode}: подтверждена.\nПодтвердил: ${actor}\nДата/время: ${nowText}`,
        reply_markup: { inline_keyboard: [] }
      });
    } else {
      reqEntry.status = "rejected";
      reqEntry.decidedAt = nowIso;
      reqEntry.decidedBy = actor;
      appendTaskHistory(payload, rqTaskId, actor, "Telegram: переназначение отклонено");
      await tg(token, "sendMessage", {
        chat_id: String(reqEntry.requesterChatId || ""),
        text:
          `Задача №${rqTaskId}: запрос на переназначение (${reassignCode}) отклонён.\n` +
          `Отклонил: ${actor}\nДата/время: ${nowText}`
      });
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача ${reassignCode}: отклонена.\nОтклонил: ${actor}\nДата/время: ${nowText}`,
        reply_markup: { inline_keyboard: [] }
      });
    }
    if (decision !== "y") {
      logStore[rqTaskId].unshift({
        id: requestId,
        code: reassignCode,
        status: reqEntry.status,
        currentStatus: "",
        comment: "",
        readAt: "",
        sentAt: "",
        from: fromName,
        to: toName,
        reasonType: reqEntry.reasonType || "",
        reasonText: reqEntry.reasonText || "",
        department: reqEntry.departmentName || "",
        requestedBy: reqEntry.requesterName || "",
        decidedBy: actor,
        createdAt: reqEntry.createdAt || nowIso,
        decidedAt: nowIso
      });
      if (logStore[rqTaskId].length > 100) logStore[rqTaskId].length = 100;
      await savePayload(pool, payload);
    }
    await answerOk();
    return;
  }

  await answerOk();
}

function canConfirmCloseInPayload(payload, chatId) {
  const reqMap = payload?.telegramCloseRequests;
  if (reqMap && typeof reqMap === "object") {
    const c = String(chatId).trim();
    for (const key of Object.keys(reqMap)) {
      const req = reqMap[key];
      const allowed = Array.isArray(req?.allowedConfirmChatIds)
        ? req.allowedConfirmChatIds.map((x) => String(x).trim()).filter(Boolean)
        : [];
      if (allowed.includes(c)) return true;
    }
  }

  const c = String(chatId).trim();
  const emp = findEmployeeByChatId(getEmployeesSection(payload), c);
  if (!emp) return false;
  if (String(emp[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") return false;
  const empId = String(emp[EMPLOYEE_COLUMNS.id] ?? "").trim();
  if (!empId) return false;
  const ds = payload.displaySettings || {};
  const allow = new Set(
    Array.isArray(ds.telegramCloseConfirmAllowedIds)
      ? ds.telegramCloseConfirmAllowedIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  const dup = new Set(
    Array.isArray(ds.telegramGlobalDuplicateRecipientIds)
      ? ds.telegramGlobalDuplicateRecipientIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  return allow.has(empId) && dup.has(empId);
}

/**
 * Создаёт запрос на закрытие задачи (telegramCloseRequests[taskId]), редактирует
 * исходное сообщение «Запрос отправлен …» и уведомляет согласующих. Используется
 * после шага «прикрепить файл / пропустить» при закрытии задачи через бот.
 */
async function submitTaskCloseRequest({ pool, token, payload, chatId, taskId, row, viewerRow, empName, baseTaskId, assigneeName, messageId, appTz }) {
  if (!payload.telegramCloseRequests) payload.telegramCloseRequests = {};
  const requestAllowed = new Set();
  getDepartmentHeadChatIdsForTask(payload, row).forEach((cid) => requestAllowed.add(cid));
  getAlwaysConfirmChatIds(payload).forEach((cid) => requestAllowed.add(cid));
  const ds = payload.displaySettings || {};
  const allow = new Set(
    Array.isArray(ds.telegramCloseConfirmAllowedIds)
      ? ds.telegramCloseConfirmAllowedIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  const dup = new Set(
    Array.isArray(ds.telegramGlobalDuplicateRecipientIds)
      ? ds.telegramGlobalDuplicateRecipientIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  const employees = getEmployeesSection(payload);
  const rows = employees?.rows || [];
  for (const r of rows) {
    const id = String(r[EMPLOYEE_COLUMNS.id] ?? "").trim();
    if (!id || !allow.has(id) || !dup.has(id)) continue;
    if (String(r[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") continue;
    const cid = String(r[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (cid) requestAllowed.add(cid);
  }
  // Запоминаем статус до закрытия (для отката при отклонении админом)
  // и переводим задачу в "Проверка" — будет видно админу в одноимённой вкладке/фильтре.
  const previousStatus = String(row[TASK_COLUMNS.status] || "В процессе").trim() || "В процессе";
  payload.telegramCloseRequests[taskId] = {
    chatId: String(chatId),
    employeeName: empName,
    requesterAssigneeName: assigneeName || "",
    baseTaskId,
    reassignCode: taskId.includes("/") ? taskId : "",
    at: Date.now(),
    sourceMessageId: Number(messageId) || null,
    allowedConfirmChatIds: Array.from(requestAllowed),
    previousStatus
  };
  row[TASK_COLUMNS.status] = "Проверка";
  appendTaskHistory(
    payload,
    taskId,
    empName,
    `Telegram: запрошено закрытие задачи, статус → «Проверка» (ожидает подтверждения${requestAllowed.size ? `, согласующих: ${requestAllowed.size}` : ""})`
  );
  setLastTaskContext(payload, chatId, taskId, messageId);
  await savePayload(pool, payload);
  if (messageId) {
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nЗапрос на закрытие отправлен администратору. Ожидайте подтверждения.`,
      reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, viewerRow, chatId, appTz) }
    });
  } else {
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: `${taskCaptionWithPlan(viewerRow, payload)}\n\nЗапрос на закрытие отправлен администратору. Ожидайте подтверждения.`,
      reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, viewerRow, chatId, appTz) }
    });
  }
  await notifyCloseConfirmRecipients(pool, token, payload, taskId, row, empName);
}

async function notifyCloseConfirmRecipients(pool, token, payload, taskId, row, requesterName) {
  const closeReq = payload?.telegramCloseRequests?.[String(taskId)] || {};
  const targetChatIds = Array.isArray(closeReq.allowedConfirmChatIds)
    ? closeReq.allowedConfirmChatIds.map((x) => String(x).trim()).filter(Boolean)
    : [];
  const text = `${buildFullTaskMessage(row, payload)}\n\nЗапрос на закрытие задачи №${escapeTgHtml(taskId)}\nОт: ${escapeTgHtml(requesterName)}\n\nПодтвердите или отклоните закрытие.`;
  const kb = {
    inline_keyboard: [
      [
        { text: "Подтвердить закрытие", callback_data: cb(taskId, "ad|y") },
        { text: "Отклонить", callback_data: cb(taskId, "ad|n") }
      ]
    ]
  };
  // Запоминаем message_id каждого admin-уведомления. Позже при решении (web/bot)
  // отредактируем эти сообщения, чтобы кнопки исчезли и текст сменился на «уже обработано».
  const promptMessages = {};
  for (const cid of targetChatIds) {
    const sent = await tg(token, "sendMessage", { chat_id: cid, text, reply_markup: kb });
    const mid = Number(sent?.result?.message_id) || 0;
    if (mid) promptMessages[cid] = mid;
  }
  // Сохраняем обратно в payload, если есть свежий closeReq в БД.
  if (Object.keys(promptMessages).length > 0) {
    try {
      const currentPayload = await loadPayload(pool);
      if (currentPayload?.telegramCloseRequests?.[String(taskId)]) {
        currentPayload.telegramCloseRequests[String(taskId)].confirmPrompts = promptMessages;
        await savePayload(pool, currentPayload);
      }
    } catch (_) { /* noop */ }
  }
}

/**
 * После решения админа (через бот или веб) — редактируем все ранее посланные
 * "Подтвердить/Отклонить" сообщения, убирая кнопки и подменяя текст на финальную
 * заметку, чтобы другие админские чаты не видели уже устаревший prompt.
 */
export async function clearCloseConfirmPrompts(token, closeReq, taskId, finalNote, exceptChatId = null) {
  if (!closeReq || typeof closeReq !== "object") return;
  const prompts = closeReq.confirmPrompts && typeof closeReq.confirmPrompts === "object"
    ? closeReq.confirmPrompts
    : {};
  const noteText = `Задача №${String(taskId).trim()}: ${String(finalNote || "запрос обработан").trim()}.`;
  for (const [cid, mid] of Object.entries(prompts)) {
    if (exceptChatId && String(cid) === String(exceptChatId)) continue;
    try {
      await tg(token, "editMessageText", {
        chat_id: cid,
        message_id: Number(mid) || 0,
        text: noteText,
        reply_markup: { inline_keyboard: [] }
      });
    } catch (_) { /* noop — old message could be gone */ }
  }
}

async function handleMessage(msg, pool, token) {
  const chatId = msg.chat?.id;
  if (chatId == null) return;
  const chatKey = String(chatId);
  const text = String(msg.text || "").trim();
  const hasPhoto = Array.isArray(msg.photo) && msg.photo.length > 0;
  const docMime = String(msg.document?.mime_type || "").toLowerCase();
  const hasImageDocument = Boolean(msg.document?.file_id) && docMime.startsWith("image/");
  const messageId = Number(msg.message_id) || null;

  if (msg.contact?.phone_number) {
    const payload = await loadPayload(pool);
    const employees = getEmployeesSection(payload);
    const emp = findEmployeeByPhone(employees, msg.contact.phone_number);
    if (!emp) {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text:
          "Не найден сотрудник с таким номером в справочнике. Проверьте номер в карточке сотрудника (международный формат, например +998..., +7..., +90...) и попробуйте снова.",
        reply_markup: contactShareKeyboard()
      });
      return;
    }
    await bindEmployeeToChat({ employees, employee: emp, chatId, from: msg.from, pool, payload, token });
    return;
  }

  if (text === "/отмена" || text === "/cancel") {
    const payload = await loadPayload(pool);
    const sess = payload.telegramSessions?.[chatKey];
    const last = payload.telegramLastTaskByChat?.[chatKey];
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await safeDeleteMessage(token, chatId, messageId);

    const taskId = String(sess?.taskId || last?.taskId || "").trim();
    const promptMessageId = Number(sess?.promptMessageId || last?.promptMessageId) || null;
    if (taskId) {
      const tasks = getTasksSection(payload);
      let row = findTaskRow(tasks, taskId);
      if (!row && taskId.includes("/")) {
        row = findTaskRow(tasks, String(taskId).split("/")[0]);
      }
      const appTz = resolveAppTimeZone(payload);
      const scopedRow = row ? buildTaskRowForChat(payload, row, chatId, taskId) : null;
      if (row && promptMessageId) {
        const edited = await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: promptMessageId,
          text: `${taskCaptionWithPlan(scopedRow, payload)}\n\nДействие отменено.`,
          reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, scopedRow, chatId, appTz) }
        });
        if (edited?.ok) return;
      }
      if (row) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `${taskCaptionWithPlan(scopedRow, payload)}\n\nДействие отменено.`,
          reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, scopedRow, chatId, appTz) }
        });
        return;
      }
    }
    await tg(token, "sendMessage", { chat_id: chatId, text: "Действие отменено." });
    return;
  }

  if (/^\/start(?:@\w+)?(?:\s|$)/i.test(text)) {
    await handleTelegramStart(msg, pool, token);
    return;
  }

  // Команды для восстановления быстрой клавиатуры, если она пропала.
  // Telegram-клиенты иногда скрывают reply-keyboard или она сбрасывается
  // другими ботами/действиями — теперь её можно вернуть командой.
  if (/^\/(menu|меню|keyboard|клавиатура)(?:@\w+)?(?:\s|$)/i.test(text)
    || text === "Меню"
    || text === "Главное меню") {
    const payloadForMenu = await loadPayload(pool);
    const employeesForMenu = getEmployeesSection(payloadForMenu);
    const empForMenu = findEmployeeByChatId(employeesForMenu, chatKey);
    if (!empForMenu) {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: "Сначала подключитесь к боту: отправьте /start или поделитесь номером.",
        reply_markup: contactShareKeyboard()
      });
      return;
    }
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: "Главное меню. Используйте кнопки ниже.",
      reply_markup: quickUserKeyboard()
    });
    return;
  }

  if (text === "Мои задачи") {
    const payload = await loadPayload(pool);
    const employees = getEmployeesSection(payload);
    if (!findEmployeeByChatId(employees, chatKey)) {
      await restoreEmployeeBindingFromPhoneMap({ payload, employees, chatId, pool });
    }
    const myRows = getMyTasksForChat(payload, chatId);
    const objects = getMyTaskObjects(myRows);
    if (!objects.length) {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: "У вас пока нет доступных задач.",
        reply_markup: quickUserKeyboard()
      });
      return;
    }
    const kb = {
      inline_keyboard: objects.slice(0, 50).map((name, idx) => [
        { text: name, callback_data: myTasksCb("ob", String(idx)) }
      ])
    };
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: "Мои задачи\n\nВыберите объект:",
      reply_markup: kb
    });
    return;
  }

  if (text === "Просроченные задачи") {
    const payload = await loadPayload(pool);
    const employees = getEmployeesSection(payload);
    if (!findEmployeeByChatId(employees, chatKey)) {
      await restoreEmployeeBindingFromPhoneMap({ payload, employees, chatId, pool });
    }
    const appTz = resolveAppTimeZone(payload);
    const overdueRows = getOverdueRowsForChat(payload, chatId, appTz);
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: buildOverdueSummaryText(overdueRows.length),
      reply_markup: overdueSummaryKeyboard(overdueRows.length)
    });
    return;
  }

  const payload = await loadPayload(pool);
  const employeesForRestore = getEmployeesSection(payload);
  if (!findEmployeeByChatId(employeesForRestore, chatKey)) {
    await restoreEmployeeBindingFromPhoneMap({ payload, employees: employeesForRestore, chatId, pool });
  }
  let sess = payload.telegramSessions?.[chatKey];

  if ((!sess || !sess.taskId) && (text || hasPhoto || hasImageDocument)) {
    const last = payload.telegramLastTaskByChat?.[chatKey];
    const lastTaskId = String(last?.taskId || "").trim();
    const lastAt = Number(last?.at) || 0;
    const freshEnough = lastAt > 0 && Date.now() - lastAt <= 6 * 60 * 60 * 1000;
    if (lastTaskId && freshEnough) {
      sess = {
        expect: "comment",
        taskId: lastTaskId,
        promptMessageId: Number(last?.promptMessageId) || null
      };
      if (!payload.telegramSessions) payload.telegramSessions = {};
      payload.telegramSessions[chatKey] = sess;
      await savePayload(pool, payload);
    }
  }
  if (!sess || !sess.taskId) {
    const employees = getEmployeesSection(payload);
    const emp = findEmployeeByChatId(employees, chatKey);
    if (!emp && (text || hasPhoto || hasImageDocument)) {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text:
          "Чтобы подключиться к задачам, нажмите кнопку ниже и отправьте свой контакт. Либо используйте ссылку вида /start e_<ID>.",
        reply_markup: contactShareKeyboard()
      });
      return;
    }
    if (emp && (hasPhoto || hasImageDocument)) {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: "Отправка фото по задачам отключена. Используйте комментарий и файлы в веб-интерфейсе."
      });
      return;
    }
    if (emp && text) {
      // Подключённый пользователь прислал текст без активной сессии и без совпадения
      // со встроенными командами. Вернём клавиатуру с подсказкой — а то она у
      // некоторых пользователей пропадает после долгого простоя или сторонних действий.
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: "Используйте кнопки ниже:\n• «Мои задачи» — список ваших задач по объектам.\n• «Просроченные задачи» — задачи с просрочкой.\nИли наберите /menu чтобы снова показать меню.",
        reply_markup: quickUserKeyboard()
      });
    }
    return;
  }

  const tasks = getTasksSection(payload);
  let row = findTaskRow(tasks, sess.taskId);
  if (!row && String(sess.taskId || "").includes("/")) {
    row = findTaskRow(tasks, String(sess.taskId || "").split("/")[0]);
  }
  if (!row) {
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    return;
  }

  const employees = getEmployeesSection(payload);
  const emp = findEmployeeByChatId(employees, chatKey);
  const empName = emp ? String(emp[EMPLOYEE_COLUMNS.fullName] || "").trim() : `chat ${chatKey}`;
  const taskId = String(sess.taskId || row[TASK_COLUMNS.number] || "").trim();
  const promptMessageId = Number(sess.promptMessageId) || null;
  const appTz = resolveAppTimeZone(payload);
  const getChatRow = () => buildTaskRowForChat(payload, row, chatId, taskId);

  // Шаг прикрепления файла-обоснования при закрытии задачи через бот.
  if (sess.expect === "closeTaskFile") {
    const viewerRow = getChatRow();
    const viewerAssigneeName = getTaskAssigneeNameByChat(payload, viewerRow, chatId);
    const baseTaskId = String(sess.baseTaskId || taskId || "").trim();
    const assigneeName = String(sess.assigneeName || viewerAssigneeName || "").trim();

    const lowerText = String(text || "").toLowerCase();
    const wantSkip = lowerText === "/пропустить" || lowerText === "/skip" || lowerText === "пропустить" || lowerText === "skip";
    const wantDone = lowerText === "/готово" || lowerText === "готово" || lowerText === "/done" || lowerText === "done";

    // Принимаем любой медиа-тип: документ, фото, видео, video_note, аудио,
    // голосовое, анимация (GIF). Берём первый непустой file_id + размер.
    let pickedFileId = "";
    let pickedName = "";
    let pickedMime = "";
    let pickedSize = 0;
    if (msg.document?.file_id) {
      pickedFileId = String(msg.document.file_id);
      pickedName = String(msg.document.file_name || "");
      pickedMime = String(msg.document.mime_type || "");
      pickedSize = Number(msg.document.file_size) || 0;
    } else if (msg.video?.file_id) {
      pickedFileId = String(msg.video.file_id);
      pickedName = String(msg.video.file_name || "video.mp4");
      pickedMime = String(msg.video.mime_type || "video/mp4");
      pickedSize = Number(msg.video.file_size) || 0;
    } else if (msg.animation?.file_id) {
      pickedFileId = String(msg.animation.file_id);
      pickedName = String(msg.animation.file_name || "animation.mp4");
      pickedMime = String(msg.animation.mime_type || "video/mp4");
      pickedSize = Number(msg.animation.file_size) || 0;
    } else if (msg.video_note?.file_id) {
      pickedFileId = String(msg.video_note.file_id);
      pickedName = "video_note.mp4";
      pickedMime = "video/mp4";
      pickedSize = Number(msg.video_note.file_size) || 0;
    } else if (Array.isArray(msg.photo) && msg.photo.length > 0) {
      const largest = msg.photo[msg.photo.length - 1];
      pickedFileId = String(largest?.file_id || "");
      pickedName = "";
      pickedMime = "image/jpeg";
      pickedSize = Number(largest?.file_size) || 0;
    } else if (msg.audio?.file_id) {
      pickedFileId = String(msg.audio.file_id);
      pickedName = String(msg.audio.file_name || "audio.mp3");
      pickedMime = String(msg.audio.mime_type || "audio/mpeg");
      pickedSize = Number(msg.audio.file_size) || 0;
    } else if (msg.voice?.file_id) {
      pickedFileId = String(msg.voice.file_id);
      pickedName = "voice.ogg";
      pickedMime = String(msg.voice.mime_type || "audio/ogg");
      pickedSize = Number(msg.voice.file_size) || 0;
    }

    const attachedNames = Array.isArray(sess.attachedNames) ? sess.attachedNames.slice() : [];

    // Рендер блока «прикреплено»+клавиатура. Зависит от того, что уже привязано.
    const renderPromptText = (extraNote = "") => {
      let body = `${taskCaptionWithPlan(viewerRow, payload)}\n\nПрикрепите файлы-обоснования закрытия.`;
      if (attachedNames.length > 0) {
        body += `\n\n📎 Прикреплено: ${attachedNames.length}`;
        body += `\n${attachedNames.map((n, i) => `${i + 1}. ${escapeTgHtml(n)}`).join("\n")}`;
        body += `\n\nМожете отправить ещё файл или нажмите «Готово».`;
      } else {
        body += `\nМожно отправить несколько файлов подряд. Лимит Telegram-бота — до 20 МБ на файл.\n\nКогда закончите — нажмите «Готово». Если документа нет — нажмите «Пропустить».`;
      }
      if (extraNote) body += `\n\n${extraNote}`;
      return body;
    };
    const renderKeyboard = () => {
      // Если уже прикреплён хотя бы один файл — Пропустить заменяется на Готово.
      const rightBtn = attachedNames.length > 0
        ? { text: "✅ Готово", callback_data: cb(taskId, "cdn") }
        : { text: "⏭ Пропустить", callback_data: cb(taskId, "csk") };
      return [[
        { text: "⬅️ Назад", callback_data: cb(taskId, "bk") },
        rightBtn
      ]];
    };
    const editOrSendPrompt = async (extraNote = "") => {
      const promptText = renderPromptText(extraNote);
      const keyboard = renderKeyboard();
      const target = promptMessageId;
      if (target) {
        const edited = await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: target,
          text: promptText,
          reply_markup: { inline_keyboard: keyboard }
        });
        if (edited?.ok) return;
      }
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: promptText,
        reply_markup: { inline_keyboard: keyboard }
      });
    };

    if (wantSkip && attachedNames.length === 0) {
      clearSession(payload, chatKey);
      await savePayload(pool, payload);
      await submitTaskCloseRequest({
        pool, token, payload, chatId, taskId, row, viewerRow, empName,
        baseTaskId, assigneeName, messageId: promptMessageId, appTz
      });
      return;
    }

    if (wantDone || (wantSkip && attachedNames.length > 0)) {
      clearSession(payload, chatKey);
      await savePayload(pool, payload);
      await submitTaskCloseRequest({
        pool, token, payload, chatId, taskId, row, viewerRow, empName,
        baseTaskId, assigneeName, messageId: promptMessageId, appTz
      });
      return;
    }

    if (pickedFileId) {
      // Telegram Bot API позволяет скачивать через getFile только файлы до 20 МБ.
      // Файлы больше — отклоняем явно, иначе getFile вернёт ошибку «file is too big».
      const MAX_TG_BOT_DOWNLOAD = 20 * 1024 * 1024;
      if (pickedSize > MAX_TG_BOT_DOWNLOAD) {
        const mb = Math.round((pickedSize / 1024 / 1024) * 10) / 10;
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `Файл ~${mb} МБ — это больше лимита Telegram-бота (20 МБ). Telegram не разрешает ботам скачивать файлы крупнее. Отправьте файл поменьше или загрузите его через веб-интерфейс задачи.`
        });
        return;
      }
      const attachment = await downloadTelegramFileAsAttachment(token, pickedFileId, pickedName, pickedMime);
      if (!attachment) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: "Не удалось сохранить файл. Попробуйте ещё раз или нажмите «Готово»/«Пропустить»."
        });
        return;
      }
      appendTaskAttachmentEntry(payload, taskId, attachment);
      if (baseTaskId && baseTaskId !== taskId) {
        appendTaskAttachmentEntry(payload, baseTaskId, attachment);
      }
      attachedNames.push(attachment.name || "(без имени)");
      sess.attachedNames = attachedNames;
      if (!payload.telegramSessions) payload.telegramSessions = {};
      payload.telegramSessions[chatKey] = sess;
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: прикреплён файл-обоснование закрытия — ${attachment.name || attachment.stored}`
      );
      await savePayload(pool, payload);
      // НЕ удаляем сообщение пользователя — пусть остаётся как подтверждение.
      await editOrSendPrompt();
      return;
    }

    // Ни команда, ни медиа — переспрашиваем и перечисляем что принимаем.
    await editOrSendPrompt("⚠️ Пожалуйста, отправьте файл одним сообщением (документ любого формата, фото, видео, GIF, аудио или голосовое), либо нажмите кнопку.");
    return;
  }

  if (sess.expect === "comment" && text) {
    const activeForChat = getActiveReassignForChat(payload, row, chatId, taskId);
    const prevPlan = String(row[TASK_COLUMNS.plan] || "").trim();
    const nextPlan = String(text || "").trim().slice(0, 4000);
    if (activeForChat) {
      activeForChat.comment = nextPlan;
      activeForChat.updatedAt = formatRuDateTime(new Date(), appTz);
    } else {
      row[TASK_COLUMNS.plan] = nextPlan;
    }
    const assigneeName = getTaskAssigneeNameByChat(payload, getChatRow(), chatId);
    if (assigneeName && !activeForChat) {
      const nowText = formatRuDateTime(new Date(), appTz);
      const state = getTaskMultiStateForRow(payload, row, { create: true });
      if (state && state[assigneeName]) {
        state[assigneeName].comment = nextPlan;
        state[assigneeName].updatedAt = nowText;
      }
      updateTaskAggregateStatusFromMulti(payload, row, appTz);
    }
    if (activeForChat) {
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: обновлены «Комментарии сотрудника (Результат)» по переназначению — ${nextPlan.slice(0, 2000)}`
      );
    } else if (prevPlan) {
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: обновлены «Комментарии сотрудника (Результат)»\nБыло: ${prevPlan.slice(0, 2000)}\nСтало: ${nextPlan.slice(0, 2000)}`
      );
    } else {
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: заполнены «Комментарии сотрудника (Результат)» — ${nextPlan.slice(0, 2000)}`
      );
    }
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await safeDeleteMessage(token, chatId, messageId);
    if (promptMessageId) {
      const edited = await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: promptMessageId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nКомментарий сохранён.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
      });
      if (!edited?.ok) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nКомментарий сохранён.`,
          reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
        });
      }
    } else {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nКомментарий сохранён.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
      });
    }
    await broadcastTaskCardUpdate(payload, token, row, "Комментарий сотрудника обновлён.", chatKey);
    return;
  }
  syncTaskMultiStateForRow(payload, row);

  if (sess.expect === "comment" && !text) {
    await tg(token, "sendMessage", { chat_id: chatId, text: "Пожалуйста, отправьте комментарий текстом или /отмена." });
    return;
  }

  if (sess.expect === "delayReason" && !isDelayReasonAllowedNow(getChatRow(), appTz)) {
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    if (promptMessageId) {
      const edited = await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: promptMessageId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nПричина отставания станет доступна после просрочки планового срока.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
      });
      if (edited?.ok) return;
    }
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nПричина отставания станет доступна после просрочки планового срока.`,
      reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
    });
    return;
  }

  if (sess.expect === "delayReason" && text) {
    const prevReason = String(row[TASK_COLUMNS.delayReason] || "").trim();
    const nextReason = String(text || "").trim().slice(0, 1000);
    row[TASK_COLUMNS.delayReason] = nextReason;
    appendTaskHistory(payload, taskId, empName, `Telegram: причина отставания «${prevReason || "—"}» → «${nextReason || "—"}»`);
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await safeDeleteMessage(token, chatId, messageId);
    if (promptMessageId) {
      const edited = await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: promptMessageId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nПричина отставания сохранена.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
      });
      if (!edited?.ok) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nПричина отставания сохранена.`,
          reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
        });
      }
    } else {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nПричина отставания сохранена.`,
        reply_markup: { inline_keyboard: mainKeyboardForChat(payload, taskId, getChatRow(), chatId, appTz) }
      });
    }
    await broadcastTaskCardUpdate(payload, token, row, "Причина отставания обновлена.", chatKey);
    return;
  }

  if (sess.expect === "delayReason" && !text) {
    await tg(token, "sendMessage", { chat_id: chatId, text: "Пожалуйста, отправьте причину отставания текстом или /отмена." });
    return;
  }

  if (sess.expect === "reassignReasonText" && text) {
    const reasonText = String(text || "").trim().slice(0, 1000);
    const departments = buildDepartmentOptions(payload);
    payload.telegramSessions[chatKey] = {
      expect: "reassignDepartment",
      taskId,
      reasonType: "subjective",
      reasonText,
      promptMessageId
    };
    await savePayload(pool, payload);
    await safeDeleteMessage(token, chatId, messageId);
    const keyboard = departments.map((name, i) => [{ text: name, callback_data: cb(taskId, `rad|${i}`) }]);
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "ra") }]);
    if (promptMessageId) {
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: promptMessageId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nШаг 1/2: выберите отдел`,
        reply_markup: { inline_keyboard: keyboard.slice(0, 80) }
      });
    } else {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: `${taskCaptionWithPlan(getChatRow(), payload)}\n\nШаг 1/2: выберите отдел`,
        reply_markup: { inline_keyboard: keyboard.slice(0, 80) }
      });
    }
    return;
  }

  if (sess.expect === "reassignReasonText" && !text) {
    await tg(token, "sendMessage", { chat_id: chatId, text: "Пожалуйста, отправьте причину текстом или /отмена." });
    return;
  }

  if (sess.expect === "photo") {
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: "Отправка фото по задачам отключена. Используйте комментарий и файлы в веб-интерфейсе."
    });
  }
}

function clearSession(payload, chatKey) {
  if (payload.telegramSessions && payload.telegramSessions[chatKey]) {
    delete payload.telegramSessions[chatKey];
  }
}

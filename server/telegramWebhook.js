/**
 * Обработка обновлений Telegram: /start → привязка chat_id к сотруднику, inline-кнопки у задачи,
 * комментарии, фото, согласование закрытия, запись в taskHistory в payload.
 * Токен: TELEGRAM_BOT_TOKEN в окружении или displaySettings.telegramBotToken в JSON приложения (после синхронизации).
 */

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
  lastSentAt: 19
};
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
  ["[объект]", "object"]
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
  const out = new Array(TASK_COLUMNS.lastSentAt + 1).fill("");
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

function mainKeyboard(taskNumber) {
  const n = encodeTaskNum(taskNumber);
  return [
    [
      { text: "⌛️ Сменить статус", callback_data: cb(n, "sm") },
      { text: "🗣 Комментарий", callback_data: cb(n, "cm") }
    ],
    [{ text: "📸 Отправить фото", callback_data: cb(n, "ph") }]
  ];
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

function getDepartmentHeadChatIdsForTask(payload, row) {
  const employees = getEmployeesSection(payload);
  const assignedName = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
  if (!assignedName) return [];
  const assignedEmp = findEmployeeByFullName(employees, assignedName);
  if (!assignedEmp) return [];
  const departmentName = String(assignedEmp[EMPLOYEE_COLUMNS.department] || "").trim();
  if (!departmentName) return [];
  const depRow = findDepartmentRowByName(payload, departmentName);
  if (!depRow) return [];
  const headName = String(depRow[2] || "").trim();
  if (!headName) return [];
  const headEmp = findEmployeeByFullName(employees, headName);
  if (!headEmp) return [];
  if (String(headEmp[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") return [];
  const chatId = String(headEmp[EMPLOYEE_COLUMNS.chatId] || "").trim();
  return chatId ? [chatId] : [];
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

function statusLabelWithEmoji(status) {
  const s = String(status || "").trim();
  if (!s) return "";
  const emoji = STATUS_EMOJI[s] || "•";
  const text = s === "Закрыт" ? "Закрыто" : s;
  return `${emoji} ${text}`;
}

async function tg(token, method, body) {
  const url = `https://api.telegram.org/bot${token}/${method}`;
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
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

function composeReadStateValue(isRead, whenText = "—") {
  return `${isRead ? "Прочитано" : "Не прочитано"}\n${String(whenText || "—").trim() || "—"}`;
}

function isAssignedEmployeeReader(payload, row, chatId) {
  const employees = getEmployeesSection(payload);
  const clickChat = String(chatId || "").trim();
  const clickEmp = findEmployeeByChatId(employees, clickChat);
  if (!clickEmp) return false;
  if (String(clickEmp[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") return false;

  const clickName = String(clickEmp[EMPLOYEE_COLUMNS.fullName] || "").trim();
  const assignedName = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
  const responsibleName = String(row[TASK_COLUMNS.responsible] || "").trim();

  // Строгий путь: исполнитель или ответственный по задаче.
  if (assignedName && findEmployeeByFullName(employees, assignedName) === clickEmp) return true;
  if (responsibleName && findEmployeeByFullName(employees, responsibleName) === clickEmp) return true;

  // Фолбэк: если исполнитель не сопоставился со справочником (например, старые/шаблонные ФИО),
  // считаем прочтение по факту нажатия у подключённого сотрудника, чтобы не блокировать процесс.
  const assignedEmp = assignedName ? findEmployeeByFullName(employees, assignedName) : null;
  if (!assignedEmp) return true;

  // Если у исполнителя не заполнен/неверен chat_id, также разрешаем отметку текущему подключённому сотруднику.
  const assignedChat = String(assignedEmp[EMPLOYEE_COLUMNS.chatId] || "").trim();
  if (!assignedChat) return true;

  // Иначе только сам исполнитель.
  return assignedChat === clickChat;
}

function buildFullTaskMessage(row) {
  const lines = [];
  const st = String(row[TASK_COLUMNS.status] || "").trim();
  const statusLine = st ? `${STATUS_EMOJI[st] || "⚪"} ${st}` : "—";
  lines.push(`📝 Задача №${String(row[TASK_COLUMNS.number] || "").trim() || "—"}: ${String(row[TASK_COLUMNS.task] || "").trim() || "—"}`);
  lines.push(`🏢 Объект: ${String(row[TASK_COLUMNS.object] || "").trim() || "—"}`);
  lines.push(`📌 Статус: ${statusLine}`);
  lines.push(`⚡ Приоритет: ${String(row[TASK_COLUMNS.priority] || "").trim() || "—"}`);
  lines.push(`📅 Дата: ${String(row[TASK_COLUMNS.addedDate] || "").trim() || "—"}`);
  lines.push(`🏗 Фаза: ${String(row[TASK_COLUMNS.phase] || "").trim() || "—"}`);
  lines.push(`📂 Раздел: ${String(row[TASK_COLUMNS.phaseSection] || "").trim() || "—"}`);
  lines.push(`🗂 Подраздел: ${String(row[TASK_COLUMNS.phaseSubsection] || "").trim() || "—"}`);
  lines.push(`👤 Ответственный: ${String(row[TASK_COLUMNS.assignedResponsible] || "").trim() || "—"}`);
  lines.push(`👤 Постановщик задачи: ${String(row[TASK_COLUMNS.responsible] || "").trim() || "—"}`);
  lines.push(`⏳ Срок: ${String(row[TASK_COLUMNS.dueDate] || "").trim() || "—"}`);
  const note = String(row[TASK_COLUMNS.note] || "").trim();
  if (note) lines.push(`💬 Комментарий: ${note}`);
  return lines.join("\n");
}

function taskCaptionWithPlan(row) {
  const base = buildFullTaskMessage(row);
  const plan = String(row[TASK_COLUMNS.plan] || "").trim();
  if (!plan) return base;
  const compactPlan = plan.length > 800 ? `${plan.slice(0, 797)}...` : plan;
  return `${base}\n🧩 Комментарии сотрудника (Результат):\n${compactPlan.replace(/\r\n?/g, "\n")}`;
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
  const assigned = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
  const responsible = String(row[TASK_COLUMNS.responsible] || "").trim();
  if (assigned) names.add(assigned);
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
  const assignedName = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
  const assignedEmp = assignedName ? findEmployeeByFullName(employees, assignedName) : null;
  const assignedChat = String(assignedEmp?.[EMPLOYEE_COLUMNS.chatId] || "").trim();
  if (assignedChat && String(assignedEmp?.[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен") {
    out.add(assignedChat);
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
  const actionChats = getTaskActionChatIds(payload, row);
  if (!actionChats.size) return true;
  return actionChats.has(String(chatId || "").trim());
}

async function broadcastTaskCardUpdate(payload, token, row, reasonText, excludeChatId = "") {
  const taskId = String(row[TASK_COLUMNS.number] ?? "").trim();
  const text = `${taskCaptionWithPlan(row)}\n\n${String(reasonText || "").trim() || "Обновление по задаче."}`;
  const chatIds = resolveTaskUpdateRecipientChatIds(payload, row, excludeChatId);
  const actionChats = getTaskActionChatIds(payload, row);
  for (const cid of chatIds) {
    const body = {
      chat_id: cid,
      text
    };
    if (actionChats.has(String(cid || "").trim())) {
      body.reply_markup = { inline_keyboard: mainKeyboard(taskId) };
    }
    await tg(token, "sendMessage", body);
  }
}

function defaultAcceptTemplate() {
  return "✅ Закрытие подтверждено.\n📝 Задача [ид_задачи] ([название_задачи]).";
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

async function savePayload(pool, payload) {
  await pool.query(
    `INSERT INTO app_state (id, payload, updated_at) VALUES (1, $1::jsonb, NOW())
     ON CONFLICT (id) DO UPDATE SET payload = EXCLUDED.payload, updated_at = NOW()`,
    [JSON.stringify(payload)]
  );
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

function findEmployeeByPhone(employees, phoneRaw) {
  const rows = employees?.rows || [];
  const phone = normalizePhoneForMatch(phoneRaw);
  if (!phone) return null;
  const hits = rows.filter((row) => normalizePhoneForMatch(row[EMPLOYEE_COLUMNS.phone]) === phone);
  if (hits.length === 1) return hits[0];
  return null;
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
  await savePayload(pool, payload);

  const name = String(employee[EMPLOYEE_COLUMNS.fullName] || "").trim() || "Сотрудник";
  const first = String(from?.first_name || "").trim();
  await tg(token, "sendMessage", {
    chat_id: chatId,
    text: `Здравствуйте${first ? `, ${first}` : ""}!\n\nВы подключены к боту как «${name}». Ваш Telegram ID сохранён в системе — уведомления по задачам будут приходить сюда.`
  });
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
      reply_markup: {
        keyboard: [[{ text: "Поделиться номером", request_contact: true }]],
        resize_keyboard: true,
        one_time_keyboard: true
      }
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

  const parsed = parseCallbackData(cq);
  if (!parsed) {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Некорректные данные" });
    return;
  }

  const payload = await loadPayload(pool);
  const tasks = getTasksSection(payload);
  const employees = getEmployeesSection(payload);
  const row = findTaskRow(tasks, parsed.taskNum);
  if (!row) {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: "Задача не найдена" });
    return;
  }

  const taskId = String(row[TASK_COLUMNS.number] ?? "").trim();
  const emp = findEmployeeByChatId(employees, String(chatId));
  const empName = emp ? String(emp[EMPLOYEE_COLUMNS.fullName] || "").trim() : `chat ${chatId}`;

  const answerOk = async (text) => {
    await tg(token, "answerCallbackQuery", { callback_query_id: q.id, text: text || "", show_alert: Boolean(text && text.length > 200) });
  };

  if (parsed.action === "rd") {
    const nowText = formatRuDateTime(new Date(), resolveAppTimeZone(payload));
    // Кнопка «📖 Прочитать» приходит только адресату сообщения по задаче,
    // поэтому отмечаем ознакомление сразу по факту нажатия.
    row[TASK_COLUMNS.readState] = composeReadStateValue(true, nowText);
    appendTaskHistory(payload, taskId, empName, `Telegram: задача прочитана (${nowText})`);
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${buildFullTaskMessage(row)}\n\nВыберите действие по задаче:`,
      reply_markup: { inline_keyboard: mainKeyboard(taskId) }
    });
    await answerOk("Задача отмечена как прочитанная");
    return;
  }

  if (parsed.action === "sm") {
    if (!canChatUseTaskActions(payload, row, chatId)) {
      await answerOk("Смена статуса доступна только исполнителю");
      return;
    }
    setLastTaskContext(payload, chatId, taskId, messageId);
    const keyboard = EMPLOYEE_STATUS_OPTIONS.map((label, i) => [{ text: statusLabelWithEmoji(label), callback_data: cb(taskId, `ss|${i}`) }]);
    keyboard.push([{ text: "⬅️ Назад", callback_data: cb(taskId, "bk") }]);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(row)}\n\nВыберите новый статус:`,
      reply_markup: { inline_keyboard: keyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ss") {
    if (!canChatUseTaskActions(payload, row, chatId)) {
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

    if (isClose) {
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
      payload.telegramCloseRequests[taskId] = {
        chatId: String(chatId),
        employeeName: empName,
        at: Date.now(),
        sourceMessageId: Number(messageId) || null,
        allowedConfirmChatIds: Array.from(requestAllowed)
      };
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: запрошено закрытие задачи (ожидает подтверждения${requestAllowed.size ? `, согласующих: ${requestAllowed.size}` : ""})`
      );
      setLastTaskContext(payload, chatId, taskId, messageId);
      await savePayload(pool, payload);

      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${taskCaptionWithPlan(row)}\n\nЗапрос на закрытие отправлен администратору. Ожидайте подтверждения.`,
        reply_markup: { inline_keyboard: mainKeyboard(taskId) }
      });
      await notifyCloseConfirmRecipients(pool, token, payload, taskId, row, empName);
      await answerOk();
      return;
    }

    const oldStatus = String(row[TASK_COLUMNS.status] ?? "").trim();
    row[TASK_COLUMNS.status] = newStatus;
    appendTaskHistory(payload, taskId, empName, `Telegram: статус «${oldStatus || "—"}» → «${newStatus}»`);
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);

    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(row)}\n\nСтатус обновлён.`,
      reply_markup: { inline_keyboard: mainKeyboard(taskId) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "cm") {
    if (!canChatUseTaskActions(payload, row, chatId)) {
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
      text: `${taskCaptionWithPlan(row)}\n\nНапишите комментарий одним сообщением ниже (или /отмена).`,
      reply_markup: { inline_keyboard: backOnlyKeyboard(taskId) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ph") {
    if (!canChatUseTaskActions(payload, row, chatId)) {
      await answerOk("Отправка фото доступна только исполнителю");
      return;
    }
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = { expect: "photo", taskId, promptMessageId: Number(messageId) || null };
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(row)}\n\nПришлите фото одним сообщением (или /отмена).`,
      reply_markup: { inline_keyboard: backOnlyKeyboard(taskId) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "bk") {
    clearSession(payload, String(chatId));
    setLastTaskContext(payload, chatId, taskId, messageId);
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${taskCaptionWithPlan(row)}\n\nВыберите действие по задаче:`,
      reply_markup: { inline_keyboard: mainKeyboard(taskId) }
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
      row[TASK_COLUMNS.status] = "Закрыт";
      const confirmedAt = formatRuDateTime(new Date(), resolveAppTimeZone(payload));
      const confirmInfo = `Закрытие подтверждено: ${empName || "—"}, ${confirmedAt}`;
      delete payload.telegramCloseRequests[taskId];
      const ds = payload.displaySettings || {};
      const tpl = String(ds.telegramCloseAcceptedTemplate || "").trim() || defaultAcceptTemplate();
      const msg = applySimpleTemplate(tpl, row);
      appendTaskHistory(payload, taskId, empName, `Telegram: закрытие задачи подтверждено (${confirmedAt})`);
      await savePayload(pool, payload);

      const sourceMid = Number(req.sourceMessageId) || 0;
      if (sourceMid) {
        await tg(token, "editMessageText", {
          chat_id: req.chatId,
          message_id: sourceMid,
          text: `${buildFullTaskMessage(row)}\n\n${confirmInfo}`,
          reply_markup: { inline_keyboard: mainKeyboard(taskId) }
        });
      }
      await tg(token, "sendMessage", {
        chat_id: req.chatId,
        text: `${buildFullTaskMessage(row)}\n\n${confirmInfo}\n\n${msg}`
      });
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача №${taskId}: закрытие подтверждено. Исполнителю отправлено уведомление.`,
        reply_markup: { inline_keyboard: [] }
      });
      await broadcastTaskCardUpdate(payload, token, row, confirmInfo, req.chatId);
    } else {
      delete payload.telegramCloseRequests[taskId];
      appendTaskHistory(payload, taskId, empName, `Telegram: отклонён запрос на закрытие`);
      await savePayload(pool, payload);
      await tg(token, "sendMessage", {
        chat_id: req.chatId,
        text: `Задача №${taskId}: запрос на закрытие отклонён.`
      });
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача №${taskId}: закрытие отклонено. Исполнитель уведомлён.`,
        reply_markup: { inline_keyboard: [] }
      });
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

async function notifyCloseConfirmRecipients(pool, token, payload, taskId, row, requesterName) {
  const closeReq = payload?.telegramCloseRequests?.[String(taskId)] || {};
  const targetChatIds = Array.isArray(closeReq.allowedConfirmChatIds)
    ? closeReq.allowedConfirmChatIds.map((x) => String(x).trim()).filter(Boolean)
    : [];
  const text = `${buildFullTaskMessage(row)}\n\nЗапрос на закрытие задачи №${taskId}\nОт: ${requesterName}\n\nПодтвердите или отклоните закрытие.`;
  const kb = {
    inline_keyboard: [
      [
        { text: "Подтвердить закрытие", callback_data: cb(taskId, "ad|y") },
        { text: "Отклонить", callback_data: cb(taskId, "ad|n") }
      ]
    ]
  };
  for (const cid of targetChatIds) {
    await tg(token, "sendMessage", { chat_id: cid, text, reply_markup: kb });
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
          "Не найден сотрудник с таким номером в справочнике. Проверьте номер в карточке сотрудника (международный формат, например +998..., +7..., +90...) и попробуйте снова."
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
      const row = findTaskRow(tasks, taskId);
      if (row && promptMessageId) {
        const edited = await tg(token, "editMessageText", {
          chat_id: chatId,
          message_id: promptMessageId,
          text: `${taskCaptionWithPlan(row)}\n\nДействие отменено.`,
          reply_markup: { inline_keyboard: mainKeyboard(taskId) }
        });
        if (edited?.ok) return;
      }
      if (row) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `${taskCaptionWithPlan(row)}\n\nДействие отменено.`,
          reply_markup: { inline_keyboard: mainKeyboard(taskId) }
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

  const payload = await loadPayload(pool);
  let sess = payload.telegramSessions?.[chatKey];

  if ((!sess || !sess.taskId) && (text || hasPhoto || hasImageDocument)) {
    const last = payload.telegramLastTaskByChat?.[chatKey];
    const lastTaskId = String(last?.taskId || "").trim();
    const lastAt = Number(last?.at) || 0;
    const freshEnough = lastAt > 0 && Date.now() - lastAt <= 6 * 60 * 60 * 1000;
    if (lastTaskId && freshEnough) {
      sess = {
        expect: text ? "comment" : "photo",
        taskId: lastTaskId,
        promptMessageId: Number(last?.promptMessageId) || null
      };
      if (!payload.telegramSessions) payload.telegramSessions = {};
      payload.telegramSessions[chatKey] = sess;
      await savePayload(pool, payload);
    }
  }
  if (!sess || !sess.taskId) return;

  const tasks = getTasksSection(payload);
  const row = findTaskRow(tasks, sess.taskId);
  if (!row) {
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    return;
  }

  const employees = getEmployeesSection(payload);
  const emp = findEmployeeByChatId(employees, chatKey);
  const empName = emp ? String(emp[EMPLOYEE_COLUMNS.fullName] || "").trim() : `chat ${chatKey}`;
  const taskId = String(row[TASK_COLUMNS.number] ?? "").trim();
  const promptMessageId = Number(sess.promptMessageId) || null;

  if (sess.expect === "comment" && text) {
    const prevPlan = String(row[TASK_COLUMNS.plan] || "").trim();
    const nextPlan = String(text || "").trim().slice(0, 4000);
    row[TASK_COLUMNS.plan] = nextPlan;
    if (prevPlan) {
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
        text: `${taskCaptionWithPlan(row)}\n\nКомментарий сохранён.`,
        reply_markup: { inline_keyboard: mainKeyboard(taskId) }
      });
      if (!edited?.ok) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `${taskCaptionWithPlan(row)}\n\nКомментарий сохранён.`,
          reply_markup: { inline_keyboard: mainKeyboard(taskId) }
        });
      }
    } else {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: `${taskCaptionWithPlan(row)}\n\nКомментарий сохранён.`,
        reply_markup: { inline_keyboard: mainKeyboard(taskId) }
      });
    }
    await broadcastTaskCardUpdate(payload, token, row, "Комментарий сотрудника обновлён.", chatKey);
    return;
  }

  if (sess.expect === "comment" && !text) {
    await tg(token, "sendMessage", { chat_id: chatId, text: "Пожалуйста, отправьте комментарий текстом или /отмена." });
    return;
  }

  if (sess.expect === "photo" && (hasPhoto || hasImageDocument)) {
    const fileId = hasPhoto ? String(msg.photo[msg.photo.length - 1]?.file_id || "") : String(msg.document?.file_id || "");
    if (!fileId) {
      await tg(token, "sendMessage", { chat_id: chatId, text: "Не удалось прочитать фото. Попробуйте ещё раз." });
      return;
    }
    let storedName = "";
    try {
      const media = await addTelegramPhotoToTaskMediaAfter(row, token, fileId);
      storedName = String(media?.fileName || "");
    } catch (_) {
      storedName = "";
    }
    appendTaskHistory(
      payload,
      taskId,
      empName,
      `Telegram: отправлено фото${storedName ? ` (добавлено в «Медиа после»: ${storedName})` : ""} (file_id ${fileId})`
    );
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await safeDeleteMessage(token, chatId, messageId);
    if (promptMessageId) {
      const edited = await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: promptMessageId,
        text: `${taskCaptionWithPlan(row)}\n\nФото добавлено в «Медиа после».`,
        reply_markup: { inline_keyboard: mainKeyboard(taskId) }
      });
      if (!edited?.ok) {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `${taskCaptionWithPlan(row)}\n\nФото добавлено в «Медиа после».`,
          reply_markup: { inline_keyboard: mainKeyboard(taskId) }
        });
      }
    } else {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text: `${taskCaptionWithPlan(row)}\n\nФото добавлено в «Медиа после».`,
        reply_markup: { inline_keyboard: mainKeyboard(taskId) }
      });
    }
    return;
  }

  if (sess.expect === "photo") {
    await tg(token, "sendMessage", { chat_id: chatId, text: "Пожалуйста, отправьте фото (как фото или документ-изображение) или /отмена." });
  }
}

function clearSession(payload, chatKey) {
  if (payload.telegramSessions && payload.telegramSessions[chatKey]) {
    delete payload.telegramSessions[chatKey];
  }
}

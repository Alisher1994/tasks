/**
 * Обработка обновлений Telegram: /start → привязка chat_id к сотруднику, inline-кнопки у задачи,
 * комментарии, фото, согласование закрытия, запись в taskHistory в payload.
 * Токен: TELEGRAM_BOT_TOKEN в окружении или displaySettings.telegramBotToken в JSON приложения (после синхронизации).
 */

const STATUS_OPTIONS = ["Новый", "В процессе", "Треб. реш. рук.", "Закрыт"];
const TASK_COLUMNS = {
  number: 0,
  object: 1,
  status: 2,
  priority: 3,
  addedDate: 4,
  phase: 5,
  phaseSection: 6,
  phaseSubsection: 7,
  assignedResponsible: 8,
  task: 9,
  responsible: 10,
  note: 11,
  plan: 12,
  fact: 13,
  dueDate: 14,
  closedDate: 15,
  mediaBefore: 16,
  mediaAfter: 17
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

const TASK_HISTORY_MAX = 300;

const PLACEHOLDERS = [
  ["[ид_задачи]", "number"],
  ["[Ид]", "number"],
  ["[название_задачи]", "task"],
  ["[статус]", "status"],
  ["[объект]", "object"]
];

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
  return out;
}

function mainKeyboard(taskNumber) {
  const n = encodeTaskNum(taskNumber);
  return [
    [
      { text: "Сменить статус", callback_data: cb(n, "sm") },
      { text: "Комментарий", callback_data: cb(n, "cm") }
    ],
    [{ text: "Отправить фото", callback_data: cb(n, "ph") }]
  ];
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
  return `№ ${num}${title ? `: ${title}` : ""}\nСтатус: ${st || "—"}`;
}

function defaultAcceptTemplate() {
  return "Задача [ид_задачи] ([название_задачи]): запрос на закрытие принят администратором.";
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
  return raw && typeof raw === "object" ? JSON.parse(JSON.stringify(raw)) : {};
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
  if (digits.startsWith("+")) digits = digits.slice(1);
  digits = digits.replace(/\D/g, "");
  if (!digits) return "";
  /** Локальный формат Узбекистана без кода страны → приводим к 998XXXXXXXXX. */
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

  if (parsed.action === "sm") {
    const keyboard = STATUS_OPTIONS.map((label, i) => [{ text: label, callback_data: cb(taskId, `ss|${i}`) }]);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${shortTaskCaption(row)}\n\nВыберите новый статус:`,
      reply_markup: { inline_keyboard: keyboard }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ss") {
    const idx = Number(parsed.rest[0]);
    if (!Number.isFinite(idx) || idx < 0 || idx >= STATUS_OPTIONS.length) {
      await answerOk("Неверный статус");
      return;
    }
    const newStatus = STATUS_OPTIONS[idx];
    const isClose = newStatus === "Закрыт";

    if (isClose) {
      if (!payload.telegramCloseRequests) payload.telegramCloseRequests = {};
      payload.telegramCloseRequests[taskId] = {
        chatId: String(chatId),
        employeeName: empName,
        at: Date.now()
      };
      appendTaskHistory(
        payload,
        taskId,
        empName,
        `Telegram: запрошено закрытие задачи (ожидает подтверждения)`
      );
      await savePayload(pool, payload);

      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `${shortTaskCaption(row)}\n\nЗапрос на закрытие отправлен администратору. Ожидайте подтверждения.`,
        reply_markup: { inline_keyboard: mainKeyboard(taskId) }
      });
      await notifyCloseConfirmRecipients(pool, token, payload, taskId, row, empName);
      await answerOk();
      return;
    }

    const oldStatus = String(row[TASK_COLUMNS.status] ?? "").trim();
    row[TASK_COLUMNS.status] = newStatus;
    appendTaskHistory(payload, taskId, empName, `Telegram: статус «${oldStatus || "—"}» → «${newStatus}»`);
    await savePayload(pool, payload);

    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${shortTaskCaption(row)}\n\nСтатус обновлён.`,
      reply_markup: { inline_keyboard: mainKeyboard(taskId) }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "cm") {
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = { expect: "comment", taskId };
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${shortTaskCaption(row)}\n\nНапишите комментарий одним сообщением ниже (или /отмена).`,
      reply_markup: { inline_keyboard: [] }
    });
    await answerOk();
    return;
  }

  if (parsed.action === "ph") {
    if (!payload.telegramSessions) payload.telegramSessions = {};
    payload.telegramSessions[String(chatId)] = { expect: "photo", taskId };
    await savePayload(pool, payload);
    await tg(token, "editMessageText", {
      chat_id: chatId,
      message_id: messageId,
      text: `${shortTaskCaption(row)}\n\nПришлите фото одним сообщением (или /отмена).`,
      reply_markup: { inline_keyboard: [] }
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
    if (!canConfirmCloseInPayload(payload, String(chatId))) {
      await answerOk("Недостаточно прав");
      return;
    }
    const req = payload.telegramCloseRequests?.[taskId];
    if (!req) {
      await answerOk("Запрос уже обработан");
      return;
    }

    if (decision === "y") {
      row[TASK_COLUMNS.status] = "Закрыт";
      delete payload.telegramCloseRequests[taskId];
      const ds = payload.displaySettings || {};
      const tpl = String(ds.telegramCloseAcceptedTemplate || "").trim() || defaultAcceptTemplate();
      const msg = applySimpleTemplate(tpl, row);
      appendTaskHistory(payload, taskId, empName, `Telegram: закрытие задачи подтверждено`);
      await savePayload(pool, payload);

      await tg(token, "sendMessage", {
        chat_id: req.chatId,
        text: msg
      });
      await tg(token, "editMessageText", {
        chat_id: chatId,
        message_id: messageId,
        text: `Задача №${taskId}: закрытие подтверждено. Исполнителю отправлено уведомление.`,
        reply_markup: { inline_keyboard: [] }
      });
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
  const targets = rows.filter((r) => {
    const id = String(r[EMPLOYEE_COLUMNS.id] ?? "").trim();
    if (!id || !allow.has(id) || !dup.has(id)) return false;
    if (String(r[EMPLOYEE_COLUMNS.telegram] || "").trim() !== "Подключен") return false;
    return Boolean(String(r[EMPLOYEE_COLUMNS.chatId] || "").trim());
  });
  const text = `Запрос на закрытие задачи №${taskId}\n${String(row[TASK_COLUMNS.task] || "").trim()}\nОт: ${requesterName}`;
  const kb = {
    inline_keyboard: [
      [
        { text: "Подтвердить закрытие", callback_data: cb(taskId, "ad|y") },
        { text: "Отклонить", callback_data: cb(taskId, "ad|n") }
      ]
    ]
  };
  for (const a of targets) {
    const cid = String(a[EMPLOYEE_COLUMNS.chatId] || "").trim();
    await tg(token, "sendMessage", { chat_id: cid, text, reply_markup: kb });
  }
}

async function handleMessage(msg, pool, token) {
  const chatId = msg.chat?.id;
  if (chatId == null) return;
  const chatKey = String(chatId);
  const text = String(msg.text || "").trim();

  if (msg.contact?.phone_number) {
    const payload = await loadPayload(pool);
    const employees = getEmployeesSection(payload);
    const emp = findEmployeeByPhone(employees, msg.contact.phone_number);
    if (!emp) {
      await tg(token, "sendMessage", {
        chat_id: chatId,
        text:
          "Не найден сотрудник с таким номером в справочнике. Проверьте номер в карточке сотрудника (формат +998... или без +) и попробуйте снова."
      });
      return;
    }
    await bindEmployeeToChat({ employees, employee: emp, chatId, from: msg.from, pool, payload, token });
    return;
  }

  if (text === "/отмена" || text === "/cancel") {
    const payload = await loadPayload(pool);
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await tg(token, "sendMessage", { chat_id: chatId, text: "Действие отменено." });
    return;
  }

  if (/^\/start(?:@\w+)?(?:\s|$)/i.test(text)) {
    await handleTelegramStart(msg, pool, token);
    return;
  }

  const payload = await loadPayload(pool);
  const sess = payload.telegramSessions?.[chatKey];
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

  if (sess.expect === "comment" && text) {
    appendTaskHistory(payload, taskId, empName, `Telegram: комментарий — ${text.slice(0, 2000)}`);
    clearSession(payload, chatKey);
    await savePayload(pool, payload);
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: `Комментарий сохранён в истории задачи №${taskId}.`,
      reply_markup: { inline_keyboard: mainKeyboard(taskId) }
    });
    return;
  }

  if (sess.expect === "photo" && msg.photo && msg.photo.length) {
    const best = msg.photo[msg.photo.length - 1];
    const fileId = best.file_id;
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
    await tg(token, "sendMessage", {
      chat_id: chatId,
      text: `Фото добавлено в «Медиа после» задачи №${taskId}.`,
      reply_markup: { inline_keyboard: mainKeyboard(taskId) }
    });
    return;
  }
}

function clearSession(payload, chatKey) {
  if (payload.telegramSessions && payload.telegramSessions[chatKey]) {
    delete payload.telegramSessions[chatKey];
  }
}

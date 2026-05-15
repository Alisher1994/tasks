import { randomBytes } from "crypto";

const SMS_INVITE_LOG_LIMIT = 500;
const SMS_DEFAULT_TIMEOUT_MS = 15000;
const SMS_DEFAULT_PROVIDER = "generic";
const SMS_GATE_PROVIDER_ID = "sms-gate.app";
const SMS_DEFAULT_PHONE_FIELD = "phone";
const SMS_DEFAULT_MESSAGE_FIELD = "message";
const SMS_DEFAULT_SENDER_FIELD = "sender";
const SMS_DEFAULT_API_KEY_HEADER = "Authorization";
const SMS_DEFAULT_AUTH_TYPE = "header";
const SMS_DEFAULT_GATE_URL = "https://api.sms-gate.app/3rdparty/v1/messages";
const SMS_DEFAULT_INVITE_TEMPLATE =
  "Здравствуйте, [ФИО]. Пожалуйста, пройдите регистрацию в Telegram-боте [Бот] по ссылке: [Ссылка_бота]. После регистрации вы будете получать задачи от руководителей.";
const SMS_DEFAULT_TASK_TEMPLATE =
  "У вас есть задача №[ID_задачи]: [Название_задачи]. Для подробностей перейдите в Telegram-бот [Бот]: [Ссылка_бота].";

function normalizeSmsGatewayMethod(raw) {
  const value = String(raw || "").trim().toLowerCase();
  return value === "get-query" ? "get-query" : "post-json";
}

function normalizeSmsGatewayProvider(raw) {
  const value = String(raw || "").trim().toLowerCase();
  if (value === SMS_GATE_PROVIDER_ID) return SMS_GATE_PROVIDER_ID;
  return SMS_DEFAULT_PROVIDER;
}

function normalizeSmsGatewayAuthType(raw) {
  const value = String(raw || "").trim().toLowerCase();
  if (value === "basic") return "basic";
  if (value === "bearer") return "bearer";
  if (value === "none") return "none";
  return SMS_DEFAULT_AUTH_TYPE;
}

function normalizeSmsFieldKey(raw, fallback) {
  const value = String(raw || "").trim();
  if (!value) return fallback;
  if (!/^[A-Za-z_][A-Za-z0-9_.-]{0,63}$/.test(value)) return fallback;
  return value;
}

function normalizeSmsHeaderName(raw, fallback = SMS_DEFAULT_API_KEY_HEADER) {
  const value = String(raw || "").trim();
  if (!value) return fallback;
  if (!/^[A-Za-z0-9-]{1,64}$/.test(value)) return fallback;
  return value;
}

function normalizeSmsInviteBoolean(raw, fallback = false) {
  if (typeof raw === "boolean") return raw;
  const value = String(raw || "").trim().toLowerCase();
  if (!value) return fallback;
  if (["1", "true", "yes", "on", "да"].includes(value)) return true;
  if (["0", "false", "no", "off", "нет"].includes(value)) return false;
  return fallback;
}

function ensureSmsInviteLogArray(payload) {
  if (!Array.isArray(payload.smsInviteLog)) payload.smsInviteLog = [];
  return payload.smsInviteLog;
}

function buildSmsInviteBotLink(displaySettings, employeeId) {
  const botUsername = String(displaySettings?.telegramBotUsername || "").trim().replace(/^@+/, "");
  if (!botUsername) return "";
  const id = String(employeeId || "").trim();
  if (!id) return `https://t.me/${botUsername}`;
  return `https://t.me/${botUsername}?start=e_${encodeURIComponent(id)}`;
}

function applySmsInviteTemplate(rawTemplate, context, truncateForLog) {
  let tpl = String(rawTemplate || "").trim();
  if (!tpl) tpl = SMS_DEFAULT_INVITE_TEMPLATE;
  const replacements = new Map([
    ["[ФИО]", context.fullName || "сотрудник"],
    ["[ID]", context.employeeId || ""],
    ["[Телефон]", context.phone || ""],
    ["[Бот]", context.botLabel || "бот"],
    ["[Ссылка_бота]", context.botLink || "ссылка будет предоставлена администратором"],
    ["[full_name]", context.fullName || "employee"],
    ["[employee_id]", context.employeeId || ""],
    ["[phone]", context.phone || ""],
    ["[bot]", context.botLabel || "bot"],
    ["[bot_link]", context.botLink || ""]
  ]);
  let out = tpl;
  replacements.forEach((value, token) => {
    out = out.split(token).join(String(value || ""));
  });
  return truncateForLog(out, 2000);
}

function applySmsTaskTemplate(rawTemplate, context, truncateForLog) {
  let tpl = String(rawTemplate || "").trim();
  if (!tpl) tpl = SMS_DEFAULT_TASK_TEMPLATE;
  const replacements = new Map([
    ["[ID_задачи]", context.taskId || ""],
    ["[Название_задачи]", context.taskTitle || "задача"],
    ["[Объект]", context.objectName || ""],
    ["[Срок]", context.dueDate || ""],
    ["[ФИО]", context.fullName || "сотрудник"],
    ["[Бот]", context.botLabel || "бот"],
    ["[Ссылка_бота]", context.botLink || "ссылка будет предоставлена администратором"],
    ["[task_id]", context.taskId || ""],
    ["[task_title]", context.taskTitle || ""],
    ["[object]", context.objectName || ""],
    ["[due_date]", context.dueDate || ""],
    ["[full_name]", context.fullName || ""],
    ["[bot]", context.botLabel || "bot"],
    ["[bot_link]", context.botLink || ""]
  ]);
  let out = tpl;
  replacements.forEach((value, token) => {
    out = out.split(token).join(String(value || ""));
  });
  return truncateForLog(out, 2000);
}

function getSmsGatewaySettings(displaySettings) {
  const ds = displaySettings && typeof displaySettings === "object" ? displaySettings : {};
  const provider = normalizeSmsGatewayProvider(ds.smsGatewayProvider);
  const timeoutRaw = Number(ds.smsGatewayTimeoutMs);
  const simRaw = Number(ds.smsGatewaySimNumber);
  const ttlRaw = Number(ds.smsGatewayTtlSeconds);
  const prioRaw = Number(ds.smsGatewayPriority);
  const deviceWithinRaw = Number(ds.smsGatewayDeviceActiveWithinHours);
  return {
    enabled: ds.smsGatewayEnabled === true,
    provider,
    url: String(ds.smsGatewayUrl || "").trim() || SMS_DEFAULT_GATE_URL,
    method: provider === SMS_GATE_PROVIDER_ID ? "post-json" : normalizeSmsGatewayMethod(ds.smsGatewayMethod),
    authType: normalizeSmsGatewayAuthType(ds.smsGatewayAuthType),
    username: String(ds.smsGatewayUsername || "").trim(),
    password: String(ds.smsGatewayPassword || "").trim(),
    apiKey: String(ds.smsGatewayApiKey || "").trim(),
    apiKeyHeader: normalizeSmsHeaderName(ds.smsGatewayApiKeyHeader, SMS_DEFAULT_API_KEY_HEADER),
    sender: String(ds.smsGatewaySender || "").trim(),
    deviceId: String(ds.smsGatewayDeviceId || "").trim(),
    simNumber: Number.isFinite(simRaw) && simRaw >= 1 ? Math.floor(simRaw) : 0,
    ttlSeconds: Number.isFinite(ttlRaw) && ttlRaw > 0 ? Math.floor(ttlRaw) : 0,
    priority: Number.isFinite(prioRaw) ? Math.max(0, Math.min(100, Math.floor(prioRaw))) : 0,
    skipPhoneValidation: normalizeSmsInviteBoolean(ds.smsGatewaySkipPhoneValidation, true),
    deviceActiveWithinHours: Number.isFinite(deviceWithinRaw) && deviceWithinRaw > 0
      ? Math.min(720, Math.max(1, Math.floor(deviceWithinRaw)))
      : 0,
    phoneField: normalizeSmsFieldKey(ds.smsGatewayPhoneField, SMS_DEFAULT_PHONE_FIELD),
    messageField: normalizeSmsFieldKey(ds.smsGatewayMessageField, SMS_DEFAULT_MESSAGE_FIELD),
    senderField: normalizeSmsFieldKey(ds.smsGatewaySenderField, SMS_DEFAULT_SENDER_FIELD),
    timeoutMs: Number.isFinite(timeoutRaw)
      ? Math.min(60000, Math.max(3000, Math.floor(timeoutRaw)))
      : SMS_DEFAULT_TIMEOUT_MS,
    inviteTemplate: String(ds.smsInviteTemplate || "").trim(),
    taskTemplate: String(ds.smsTaskTemplate || "").trim()
  };
}

function buildSmsInviteMessage(payload, employeeRow, { normalizePhone, truncateForLog }) {
  const displaySettings = payload?.displaySettings || {};
  const settings = getSmsGatewaySettings(displaySettings);
  const employeeId = String(employeeRow?.[0] || "").trim();
  const fullName = String(employeeRow?.[1] || "").trim() || "Сотрудник";
  const phone = normalizePhone(employeeRow?.[4] || "");
  const botUsername = String(displaySettings?.telegramBotUsername || "").trim().replace(/^@+/, "");
  const botLabel = botUsername ? `@${botUsername}` : "Telegram-бот";
  const botLink = buildSmsInviteBotLink(displaySettings, employeeId);
  const text = applySmsInviteTemplate(settings.inviteTemplate, {
    employeeId,
    fullName,
    phone,
    botLabel,
    botLink
  }, truncateForLog);
  return { text, botLink, botLabel };
}

function buildSmsTaskMessage(payload, taskRow, employeeRow, deps) {
  const {
    truncateForLog,
    EMPLOYEE_FULL_NAME_COL,
    TASK_NUMBER_COL,
    TASK_TITLE_COL,
    TASK_DUE_DATE_COL
  } = deps;
  const displaySettings = payload?.displaySettings || {};
  const settings = getSmsGatewaySettings(displaySettings);
  const employeeId = String(employeeRow?.[0] || "").trim();
  const fullName = String(employeeRow?.[EMPLOYEE_FULL_NAME_COL] || "").trim() || "Сотрудник";
  const botUsername = String(displaySettings?.telegramBotUsername || "").trim().replace(/^@+/, "");
  const botLabel = botUsername ? `@${botUsername}` : "Telegram-бот";
  const botLink = buildSmsInviteBotLink(displaySettings, employeeId);
  const text = applySmsTaskTemplate(settings.taskTemplate, {
    taskId: String(taskRow?.[TASK_NUMBER_COL] || "").trim(),
    taskTitle: String(taskRow?.[TASK_TITLE_COL] || "").trim(),
    objectName: String(taskRow?.[1] || "").trim(),
    dueDate: String(taskRow?.[TASK_DUE_DATE_COL] || "").trim(),
    fullName,
    botLabel,
    botLink
  }, truncateForLog);
  return { text, botLink, botLabel };
}

async function sendSmsViaGateway(settings, { phone, text }, truncateForLog) {
  const url = String(settings?.url || "").trim();
  if (!url) {
    return { ok: false, reason: "Не указан URL SMS Gateway." };
  }
  const provider = normalizeSmsGatewayProvider(settings?.provider);
  const method = normalizeSmsGatewayMethod(settings?.method);
  const phoneField = normalizeSmsFieldKey(settings?.phoneField, SMS_DEFAULT_PHONE_FIELD);
  const messageField = normalizeSmsFieldKey(settings?.messageField, SMS_DEFAULT_MESSAGE_FIELD);
  const senderField = normalizeSmsFieldKey(settings?.senderField, SMS_DEFAULT_SENDER_FIELD);
  const sender = String(settings?.sender || "").trim();
  const timeoutMs = Number.isFinite(Number(settings?.timeoutMs))
    ? Math.min(60000, Math.max(3000, Math.floor(Number(settings.timeoutMs))))
    : SMS_DEFAULT_TIMEOUT_MS;

  const headers = {};
  const authType = normalizeSmsGatewayAuthType(settings?.authType);
  const username = String(settings?.username || "").trim();
  const password = String(settings?.password || "").trim();
  const apiKey = String(settings?.apiKey || "").trim();
  if (provider === SMS_GATE_PROVIDER_ID) {
    if (authType === "basic") {
      if (!username || !password) {
        return { ok: false, reason: "Для sms-gate.app (Basic) укажите username и password." };
      }
      const credentials = Buffer.from(`${username}:${password}`, "utf8").toString("base64");
      headers.Authorization = `Basic ${credentials}`;
    } else if (authType === "bearer") {
      if (!apiKey) {
        return { ok: false, reason: "Для sms-gate.app (Bearer) укажите токен." };
      }
      headers.Authorization = apiKey.toLowerCase().startsWith("bearer ")
        ? apiKey
        : `Bearer ${apiKey}`;
    } else if (authType === "header" && apiKey) {
      const keyHeader = normalizeSmsHeaderName(settings?.apiKeyHeader, SMS_DEFAULT_API_KEY_HEADER);
      headers[keyHeader] = apiKey;
    }
  } else if (authType === "basic") {
    if (!username || !password) {
      return { ok: false, reason: "Для Basic авторизации укажите username и password." };
    }
    const credentials = Buffer.from(`${username}:${password}`, "utf8").toString("base64");
    headers.Authorization = `Basic ${credentials}`;
  } else if (authType === "bearer") {
    if (!apiKey) {
      return { ok: false, reason: "Для Bearer авторизации укажите токен." };
    }
    headers.Authorization = apiKey.toLowerCase().startsWith("bearer ")
      ? apiKey
      : `Bearer ${apiKey}`;
  } else if (apiKey) {
    const keyHeader = normalizeSmsHeaderName(settings?.apiKeyHeader, SMS_DEFAULT_API_KEY_HEADER);
    headers[keyHeader] = apiKey;
  }

  const phoneNormalized = String(phone || "").trim();
  const messageText = String(text || "");
  let payload = {
    [phoneField]: phoneNormalized,
    [messageField]: messageText
  };
  if (sender) payload[senderField] = sender;

  const params = new URLSearchParams();
  if (provider === SMS_GATE_PROVIDER_ID) {
    payload = {
      textMessage: { text: messageText },
      phoneNumbers: [phoneNormalized]
    };
    const deviceId = String(settings?.deviceId || "").trim();
    const simNumber = Number(settings?.simNumber) || 0;
    const ttlSeconds = Number(settings?.ttlSeconds) || 0;
    const priority = Number(settings?.priority) || 0;
    const deviceActiveWithinHours = Number(settings?.deviceActiveWithinHours) || 0;
    const skipPhoneValidation = normalizeSmsInviteBoolean(settings?.skipPhoneValidation, true);
    if (deviceId) payload.deviceId = deviceId;
    if (simNumber > 0) payload.simNumber = simNumber;
    if (ttlSeconds > 0) payload.ttl = ttlSeconds;
    if (priority > 0) payload.priority = priority;
    params.set("skipPhoneValidation", skipPhoneValidation ? "true" : "false");
    if (deviceActiveWithinHours > 0) {
      params.set("deviceActiveWithin", String(deviceActiveWithinHours));
    }
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  let response;
  try {
    const requestUrl = params.size > 0
      ? `${url}${url.includes("?") ? "&" : "?"}${params.toString()}`
      : url;
    if (method === "get-query") {
      const u = new URL(requestUrl);
      Object.entries(payload).forEach(([k, v]) => {
        u.searchParams.set(k, String(v || ""));
      });
      response = await fetch(u.toString(), {
        method: "GET",
        headers,
        signal: controller.signal
      });
    } else {
      headers["Content-Type"] = "application/json";
      response = await fetch(requestUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
        signal: controller.signal
      });
    }
  } catch (e) {
    clearTimeout(timer);
    if (String(e?.name || "") === "AbortError") {
      return { ok: false, reason: `Таймаут SMS Gateway (${timeoutMs} мс).` };
    }
    return { ok: false, reason: String(e?.message || "Ошибка сети при отправке SMS.") };
  }
  clearTimeout(timer);

  const responseText = await response.text().catch(() => "");
  let parsed = null;
  try {
    parsed = responseText ? JSON.parse(responseText) : null;
  } catch (_) {
    parsed = null;
  }
  let providerMessage = truncateForLog(
    parsed?.message
      || parsed?.description
      || parsed?.status
      || parsed?.error
      || responseText
      || "",
    800
  );

  if (provider === SMS_GATE_PROVIDER_ID && parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
    const messageId = String(parsed?.id || "").trim();
    const state = String(parsed?.state || "").trim() || "Pending";
    const recipients = Array.isArray(parsed?.recipients) ? parsed.recipients.length : 0;
    providerMessage = truncateForLog(
      `Принято SMS Gate: state=${state}${messageId ? `, id=${messageId}` : ""}${recipients > 0 ? `, recipients=${recipients}` : ""}.`,
      300
    );
  }

  if (!response.ok) {
    return {
      ok: false,
      httpStatus: response.status,
      reason: providerMessage || `HTTP ${response.status}`,
      responsePreview: truncateForLog(responseText, 1200)
    };
  }

  return {
    ok: true,
    httpStatus: response.status,
    providerMessage: providerMessage || (provider === SMS_GATE_PROVIDER_ID ? "Запрос принят SMS Gate." : ""),
    responsePreview: truncateForLog(responseText, 1200)
  };
}

export function registerSmsRoutes(app, deps) {
  const {
    authMiddleware,
    requireAdmin,
    loadAppPayload,
    saveAppPayload,
    normalizePhone,
    findEmployeeByIdInPayload,
    findEmployeeByFullNameInPayload,
    collectTaskTelegramRecipientNames,
    getTaskRows,
    truncateForLog,
    EMPLOYEE_FULL_NAME_COL,
    EMPLOYEE_CHAT_ID_COL,
    TASK_NUMBER_COL,
    TASK_TITLE_COL,
    TASK_DUE_DATE_COL
  } = deps;

  app.post("/api/sms/invite/send", authMiddleware, requireAdmin, async (req, res) => {
    const nowMs = Date.now();
    let payload = null;
    let logEntry = null;
    try {
      const employeeId = String(req.body?.employeeId || "").trim();
      if (!employeeId) {
        return res.status(400).json({ ok: false, error: "employeeId обязателен." });
      }
      payload = await loadAppPayload();
      const employeeRow = findEmployeeByIdInPayload(payload, employeeId);
      if (!employeeRow) {
        return res.status(404).json({ ok: false, error: "Сотрудник не найден." });
      }

      const fullName = String(employeeRow?.[EMPLOYEE_FULL_NAME_COL] || "").trim() || "Сотрудник";
      const phone = normalizePhone(employeeRow?.[4] || "");
      const chatId = String(employeeRow?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
      const smsSettings = getSmsGatewaySettings(payload?.displaySettings || {});
      const inviteMessage = buildSmsInviteMessage(payload, employeeRow, { normalizePhone, truncateForLog });
      const actorName = String(req.user?.name || "").trim() || "Администратор";

      logEntry = {
        id: randomBytes(10).toString("hex"),
        atMs: nowMs,
        atIso: new Date(nowMs).toISOString(),
        employeeId,
        employeeName: fullName,
        phone,
        chatId,
        text: truncateForLog(inviteMessage.text, 2000),
        actor: actorName,
        status: "failed",
        ok: false,
        gatewayProvider: smsSettings.provider,
        gatewayMethod: smsSettings.method,
        gatewayUrl: truncateForLog(smsSettings.url, 240),
        httpStatus: 0,
        resultMessage: "",
        responsePreview: ""
      };

      if (!phone) {
        logEntry.resultMessage = "У сотрудника не указан корректный телефон.";
      } else if (!smsSettings.enabled) {
        logEntry.resultMessage = "SMS Gateway выключен в настройках.";
      } else if (!smsSettings.url) {
        logEntry.resultMessage = "Не указан URL SMS Gateway в настройках.";
      } else if (!inviteMessage.text) {
        logEntry.resultMessage = "Текст SMS пустой. Проверьте шаблон.";
      } else {
        const sendResult = await sendSmsViaGateway(smsSettings, { phone, text: inviteMessage.text }, truncateForLog);
        logEntry.ok = sendResult.ok === true;
        logEntry.status = sendResult.ok ? "sent" : "failed";
        logEntry.httpStatus = Number(sendResult.httpStatus) || 0;
        logEntry.resultMessage = truncateForLog(
          sendResult.providerMessage || sendResult.reason || (sendResult.ok ? "SMS отправлено." : "Неизвестная ошибка."),
          800
        );
        logEntry.responsePreview = truncateForLog(sendResult.responsePreview || "", 1200);
      }

      const logStore = ensureSmsInviteLogArray(payload);
      logStore.unshift(logEntry);
      if (logStore.length > SMS_INVITE_LOG_LIMIT) {
        logStore.length = SMS_INVITE_LOG_LIMIT;
      }
      await saveAppPayload(payload);

      if (logEntry.ok) {
        return res.json({ ok: true, entry: logEntry });
      }
      return res.status(400).json({ ok: false, error: logEntry.resultMessage || "Не удалось отправить SMS.", entry: logEntry });
    } catch (e) {
      console.error(e);
      try {
        if (payload && logEntry) {
          logEntry.resultMessage = truncateForLog(`Ошибка сервера: ${String(e?.message || e)}`, 800);
          logEntry.status = "failed";
          logEntry.ok = false;
          const logStore = ensureSmsInviteLogArray(payload);
          logStore.unshift(logEntry);
          if (logStore.length > SMS_INVITE_LOG_LIMIT) logStore.length = SMS_INVITE_LOG_LIMIT;
          await saveAppPayload(payload);
        }
      } catch (_) {
        /* noop */
      }
      return res.status(500).json({ ok: false, error: "Ошибка отправки SMS-приглашения." });
    }
  });

  app.post("/api/sms/task/send", authMiddleware, requireAdmin, async (req, res) => {
    const nowMs = Date.now();
    try {
      const taskId = String(req.body?.taskId || "").trim();
      if (!taskId) {
        return res.status(400).json({ ok: false, error: "taskId обязателен." });
      }

      const payload = await loadAppPayload();
      const smsSettings = getSmsGatewaySettings(payload?.displaySettings || {});
      if (!smsSettings.enabled) {
        return res.status(400).json({ ok: false, error: "SMS Gateway выключен в настройках." });
      }
      if (!smsSettings.url) {
        return res.status(400).json({ ok: false, error: "Не указан URL SMS Gateway в настройках." });
      }

      const tasksRows = getTaskRows(payload);
      const taskRow = tasksRows.find((row) => String(row?.[TASK_NUMBER_COL] || "").trim() === taskId);
      if (!taskRow) {
        return res.status(404).json({ ok: false, error: "Задача не найдена." });
      }
      const taskTitle = String(taskRow?.[TASK_TITLE_COL] || "").trim() || `Задача ${taskId}`;
      const actorName = String(req.user?.name || "").trim() || "Администратор";

      const recipientNames = collectTaskTelegramRecipientNames(taskRow);
      if (!recipientNames.length) {
        return res.status(400).json({ ok: false, error: "В задаче не указан исполнитель/ответственный для SMS." });
      }

      const recipientsByPhone = new Map();
      for (const name of recipientNames) {
        const employeeRow = findEmployeeByFullNameInPayload(payload, name);
        if (!employeeRow) continue;
        const phone = normalizePhone(employeeRow?.[4] || "");
        if (!phone) continue;
        if (!recipientsByPhone.has(phone)) {
          recipientsByPhone.set(phone, employeeRow);
        }
      }

      if (!recipientsByPhone.size) {
        return res.status(400).json({ ok: false, error: "Не найдены сотрудники с корректным телефоном для отправки SMS." });
      }

      const logStore = ensureSmsInviteLogArray(payload);
      let sentCount = 0;
      let failCount = 0;
      const errors = [];
      const entries = [];

      for (const [phone, employeeRow] of recipientsByPhone.entries()) {
        const employeeId = String(employeeRow?.[0] || "").trim();
        const employeeName = String(employeeRow?.[EMPLOYEE_FULL_NAME_COL] || "").trim() || "Сотрудник";
        const chatId = String(employeeRow?.[EMPLOYEE_CHAT_ID_COL] || "").trim();
        const taskMessage = buildSmsTaskMessage(payload, taskRow, employeeRow, {
          truncateForLog,
          EMPLOYEE_FULL_NAME_COL,
          TASK_NUMBER_COL,
          TASK_TITLE_COL,
          TASK_DUE_DATE_COL
        });
        const logEntry = {
          id: randomBytes(10).toString("hex"),
          atMs: nowMs,
          atIso: new Date(nowMs).toISOString(),
          employeeId,
          employeeName,
          phone,
          chatId,
          text: truncateForLog(taskMessage.text, 2000),
          actor: actorName,
          status: "failed",
          ok: false,
          gatewayProvider: smsSettings.provider,
          gatewayMethod: smsSettings.method,
          gatewayUrl: truncateForLog(smsSettings.url, 240),
          httpStatus: 0,
          resultMessage: "",
          responsePreview: "",
          smsKind: "task",
          taskId,
          taskTitle: truncateForLog(taskTitle, 240)
        };

        if (!taskMessage.text) {
          logEntry.resultMessage = "Текст SMS по задаче пустой. Проверьте шаблон.";
        } else {
          const sendResult = await sendSmsViaGateway(smsSettings, { phone, text: taskMessage.text }, truncateForLog);
          logEntry.ok = sendResult.ok === true;
          logEntry.status = sendResult.ok ? "sent" : "failed";
          logEntry.httpStatus = Number(sendResult.httpStatus) || 0;
          logEntry.resultMessage = truncateForLog(
            sendResult.providerMessage || sendResult.reason || (sendResult.ok ? "SMS отправлено." : "Неизвестная ошибка."),
            800
          );
          logEntry.responsePreview = truncateForLog(sendResult.responsePreview || "", 1200);
        }

        if (logEntry.ok) {
          sentCount += 1;
        } else {
          failCount += 1;
          errors.push(`${employeeName}: ${String(logEntry.resultMessage || "ошибка")}`);
        }
        entries.push(logEntry);
        logStore.unshift(logEntry);
      }

      if (logStore.length > SMS_INVITE_LOG_LIMIT) {
        logStore.length = SMS_INVITE_LOG_LIMIT;
      }
      await saveAppPayload(payload);

      if (sentCount > 0) {
        return res.json({
          ok: true,
          taskId,
          taskTitle,
          totalRecipients: recipientsByPhone.size,
          sentCount,
          failCount,
          errors,
          entries
        });
      }
      return res.status(400).json({
        ok: false,
        error: errors[0] || "Не удалось отправить SMS по задаче.",
        taskId,
        taskTitle,
        totalRecipients: recipientsByPhone.size,
        sentCount,
        failCount,
        errors,
        entries
      });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ ok: false, error: "Ошибка отправки SMS по задаче." });
    }
  });

  app.get("/api/sms/invite/history", authMiddleware, requireAdmin, async (_req, res) => {
    try {
      const payload = await loadAppPayload();
      const list = Array.isArray(payload?.smsInviteLog)
        ? payload.smsInviteLog.filter((item) => item && typeof item === "object")
        : [];
      return res.json({ ok: true, entries: list });
    } catch (e) {
      console.error(e);
      return res.status(500).json({ ok: false, error: "Не удалось загрузить историю SMS." });
    }
  });
}

const loginForm = document.getElementById("loginForm");
const loginError = document.getElementById("loginError");
const loginSection = document.getElementById("loginSection");
const appSection = document.getElementById("appSection");
const phoneInput = document.getElementById("phone");
const tabsRoot = document.getElementById("tabs");
const tableContainer = document.getElementById("tableContainer");
const logoutBtn = document.getElementById("logoutBtn");
const currentUser = document.getElementById("currentUser");
const loginBtn = document.getElementById("loginBtn");
const sidebarBrandToggle = document.getElementById("sidebarBrandToggle");
const passwordInput = document.getElementById("password");
const togglePasswordBtn = document.getElementById("togglePasswordBtn");
const loginPhoneFlag = document.getElementById("loginPhoneFlag");
const loginPhoneCountryBtn = document.getElementById("loginPhoneCountryBtn");
const turboTopLoader = document.getElementById("turboTopLoader");
const turboTopLoaderBar = document.getElementById("turboTopLoaderBar");

const AUTH_PHONE = "+998994067406";
const AUTH_PASSWORD = "7406";
const DEFAULT_PHONE_PREFIX = "+";
const PHONE_MIN_DIGITS = 8;
const PHONE_MAX_DIGITS = 15;
const PHONE_MAX_LENGTH = PHONE_MAX_DIGITS + 1;
const REPORT_SHARE_STORAGE_KEY = "mbc_report_share_links";
const OVERDUE_NOTIFY_RUNTIME_KEY = "mbc_overdue_notify_runtime";
/** JWT при работе с сервером (Railway) */
const AUTH_TOKEN_KEY = "mbc_jwt";
const SESSION_STORAGE_KEY = "mbc_task_auth_user";
const SESSION_LAST_ACTIVITY_KEY = "mbc_task_last_activity_ts";
const ACTIVE_SECTION_STORAGE_KEY = "mbc_task_active_section";
const DISPLAY_SETTINGS_KEY = "mbc_task_display_settings";
const DATA_STORAGE_KEY = "mbc_task_sections_data";
const SESSION_IDLE_TIMEOUT_MS = 12 * 60 * 60 * 1000;
const SESSION_IDLE_CHECK_INTERVAL_MS = 60 * 1000;
/** Data URL превью фото объектов (ключ obj-ph-{id}) — переживает перезагрузку; в ячейке по-прежнему имя файла */
const OBJECT_PHOTO_THUMBS_STORAGE_KEY = "mbc_object_photo_thumbs";
const TRASH_STORAGE_KEY = "mbc_task_trash_data";
/** История действий по задачам: { [taskId]: Array<{ t, who, action }> } */
const TASK_HISTORY_STORAGE_KEY = "mbc_task_action_history";
const TASK_MULTI_STATE_STORAGE_KEY = "mbc_task_multi_state";
const TASK_HISTORY_MAX_PER_TASK = 300;
const REPORT_CHART_ORDER_STORAGE_KEY = "mbc_report_chart_tile_order";
/** «row» — Топ фаз / Разделы / Подразделы в одной строке; «separate» — каждый график на всю ширину */
const REPORT_PHASE_GROUP_LAYOUT_KEY = "mbc_report_phase_group_layout";
const REPORT_TOP_TRIO_IDS = ["status", "priority", "priorityDonut"];
const REPORT_PHASE_GROUP_IDS = ["phase", "phaseSection", "phaseSubsection"];
const REPORT_PHASE_GROUP_SET = new Set(REPORT_PHASE_GROUP_IDS);
const REPORT_PHASE_DONUT_GROUP_IDS = ["phaseDonut", "phaseSectionDonut", "phaseSubsectionDonut"];
const REPORT_PHASE_DONUT_GROUP_SET = new Set(REPORT_PHASE_DONUT_GROUP_IDS);
/** id плитки → широкая колонка на всю строку */
const REPORT_CHART_TILE_META = {
  status: { wide: false },
  priority: { wide: false },
  priorityDonut: { wide: false },
  phase: { wide: false },
  phaseSection: { wide: false },
  phaseSubsection: { wide: true },
  phaseDonut: { wide: false },
  phaseSectionDonut: { wide: false },
  phaseSubsectionDonut: { wide: false },
  months: { wide: true },
  overdue: { wide: true },
  object: { wide: true },
  department: { wide: true },
  responsible: { wide: true }
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
const OBJECT_COLUMNS = {
  id: 0,
  name: 1,
  address: 2,
  status: 3,
  rp: 4,
  zrp: 5,
  /** Одно изображение: имя файла в ячейке; превью в objectPhotoPreviewStore */
  photo: 6
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
const EMPLOYEE_TELEGRAM_OPTIONS = ["Подключен", "Не подключен"];
const SYSTEM_ROLES = [
  "РП",
  "ЗРП",
  "Директор строительства",
  "Администратор",
  "Инженер планового отдела",
  "Генеральный директор",
  "ГИП",
  "Технадзор",
  "ОТиТБ",
  "Финансист",
  "Главный финансист",
  "Главный бухгалтер",
  "Бухгалтер",
  "Инженер ПТО",
  "Руководитель отдела ПТО"
];
const SYSTEM_DEPARTMENTS = [
  "ОТН",
  "ОТиТБ",
  "Финансовый отдел",
  "ПТО",
  "Бухгалтерия",
  "Производство",
  "Плановый отдел",
  "Проектная группа",
  "АУП",
  "Снабжение",
  "Тендер",
  "Аналитика",
  "Отдел развития ИТ"
];
const STATUS_DECISION_OLD = "Треб. реш. рук.";
const STATUS_DECISION = "Требует решение руководителя";
const STATUS_OPTIONS = ["Новый", "В процессе", "Закрыт"];

/** Подписи месяцев для графика «Добавление задач по месяцам» (текущий год, ось X). */
const REPORT_MONTH_LABELS_RU = ["Янв.", "Февр.", "Мар.", "Апр.", "Май", "Июн.", "Июл.", "Авг.", "Сен.", "Окт.", "Ноя.", "Дек."];
/** Строки без значения в колонке «Статус» — не отдельный столбец, а пустая ячейка */
const REPORT_NO_STATUS_LABEL = "Без статуса";
/** Только реальные статусы справочника (чекбоксы фильтра) */
const REPORT_FILTER_ALL_STATUSES = [...STATUS_OPTIONS];
/** Мягкие цвета статусов (градиентные/пастельные оттенки бренда) */
const STATUS_CHART_COLORS = {
  Новый: "#a8b4f0",
  "В процессе": "#e8c9a8",
  [STATUS_DECISION]: "#e0a8a8",
  [STATUS_DECISION_OLD]: "#e0a8a8",
  Закрыт: "#9dceb0",
  [REPORT_NO_STATUS_LABEL]: "#c5cad3",
  Прочее: "#b8c5d6"
};
/** Для подписей/легенды — чуть насыщеннее текста */
const STATUS_CHART_COLORS_ACCENT = {
  Новый: "#5c6bc0",
  "В процессе": "#c08457",
  [STATUS_DECISION]: "#b06a6a",
  [STATUS_DECISION_OLD]: "#b06a6a",
  Закрыт: "#4a9d6a",
  [REPORT_NO_STATUS_LABEL]: "#64748b",
  Прочее: "#6b7c93"
};
/** Класс фона карточки напоминания в настройках */
const REMINDER_CARD_UI = {
  Новый: "new",
  "В процессе": "progress",
  [STATUS_DECISION]: "decision",
  [STATUS_DECISION_OLD]: "decision",
  Закрыт: "closed"
};
const TELEGRAM_STATUS_EMOJI = {
  Новый: "🟣",
  "В процессе": "🟡",
  [STATUS_DECISION]: "🔴",
  [STATUS_DECISION_OLD]: "🔴",
  Закрыт: "🟢"
};

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

const DATE_DISPLAY_FORMAT_OPTIONS = [
  { id: "DMY_DOT", label: "ДД.ММ.ГГГГ (31.12.2025)" },
  { id: "ISO", label: "ГГГГ-ММ-ДД (2025-12-31)" },
  { id: "DMY_SLASH", label: "ДД/ММ/ГГГГ (31/12/2025)" },
  { id: "MDY_SLASH", label: "ММ/ДД/ГГГГ (12/31/2025)" }
];

const TIME_DISPLAY_FORMAT_OPTIONS = [
  { id: "24", label: "24 часа (13:45)" },
  { id: "12", label: "12 часов (1:45 PM)" }
];

const SERVER_TIMEZONE_OPTIONS = [
  { id: "", label: "Как в браузере (авто)" },
  { id: "UTC", label: "UTC" },
  { id: "Asia/Tashkent", label: "Asia/Tashkent" },
  { id: "Asia/Almaty", label: "Asia/Almaty" },
  { id: "Asia/Dubai", label: "Asia/Dubai" },
  { id: "Europe/Moscow", label: "Europe/Moscow" },
  { id: "Europe/Kyiv", label: "Europe/Kyiv" },
  { id: "Europe/Berlin", label: "Europe/Berlin" },
  { id: "America/New_York", label: "America/New_York" }
];
const REMINDER_DAYS_OPTIONS = [
  { value: "none", label: "Не оповещать" },
  { value: "1", label: "Каждый 1 день" },
  { value: "2", label: "Каждые 2 дня" },
  { value: "3", label: "Каждые 3 дня" },
  { value: "5", label: "Каждые 5 дней" },
  { value: "7", label: "Каждые 7 дней" },
  { value: "14", label: "Каждые 14 дней" },
  { value: "30", label: "Каждые 30 дней" }
];
/** Подстановки для текста отправки задачи (Telegram и т.п.): токен → колонка задачи */
const TASK_MESSAGE_PLACEHOLDERS_UI = [
  { token: "[ид_задачи]", col: "number", label: "ID" },
  { token: "[название_объекта]", col: "object", label: "Название объекта" },
  { token: "[статус]", col: "status", label: "Статус" },
  { token: "[приоритет]", col: "priority", label: "Приоритет" },
  { token: "[дата_постановки_задачи]", col: "addedDate", label: "Дата постановки задачи" },
  { token: "[фаза]", col: "phase", label: "Фаза" },
  { token: "[раздел]", col: "phaseSection", label: "Раздел" },
  { token: "[подраздел]", col: "phaseSubsection", label: "Подраздел" },
  { token: "[задача]", col: "task", label: "Задача" },
  { token: "[постановщик_задачи]", col: "responsible", label: "Постановщик задачи" },
  { token: "[ответственный]", col: "assignedResponsible", label: "Ответственный" },
  { token: "[коментарии_к_задаче]", col: "note", label: "Коментарии к задаче" },
  { token: "[комментарии_сотрудника_результат]", col: "plan", label: "Комментарии сотрудника (Результат)" },
  { token: "[коментарии_администратора]", col: "fact", label: "Коментарии администратора" },
  { token: "[плановый_срок_устранения]", col: "dueDate", label: "Плановый срок устранения" },
  { token: "[факт_даты_устранения]", col: "closedDate", label: "Факт даты устранения" },
  { token: "[медиа_до]", col: "mediaBefore", label: "Медиа до (5)" },
  { token: "[медиа_после]", col: "mediaAfter", label: "Медиа после (5)" }
];
/** Старые токены оставляем только для совместимости шаблонов. */
const TASK_MESSAGE_PLACEHOLDERS_LEGACY = [
  { token: "[объект]", col: "object", label: "legacy" },
  { token: "[дата_задачи]", col: "addedDate", label: "legacy" },
  { token: "[название_задачи]", col: "task", label: "legacy" },
  { token: "[закреплённый]", col: "assignedResponsible", label: "legacy" },
  { token: "[примечание]", col: "note", label: "legacy" },
  { token: "[план]", col: "plan", label: "legacy" },
  { token: "[факт]", col: "fact", label: "legacy" },
  { token: "[срок_задачи]", col: "dueDate", label: "legacy" },
  { token: "[дата_закрытия]", col: "closedDate", label: "legacy" },
  { token: "[Ид]", col: "number", label: "legacy" },
  { token: "[ФИО]", col: "assignedResponsible", label: "legacy" },
  { token: "{Название объекта}", col: "object", label: "legacy" }
];
const TASK_MESSAGE_PLACEHOLDERS = [...TASK_MESSAGE_PLACEHOLDERS_UI, ...TASK_MESSAGE_PLACEHOLDERS_LEGACY];

function isHostedRuntime() {
  return typeof window !== "undefined" && (window.location.protocol === "http:" || window.location.protocol === "https:");
}

function getAuthToken() {
  try {
    return sessionStorage.getItem(AUTH_TOKEN_KEY) || "";
  } catch (_) {
    return "";
  }
}

function setAuthToken(token) {
  try {
    if (token) sessionStorage.setItem(AUTH_TOKEN_KEY, token);
    else sessionStorage.removeItem(AUTH_TOKEN_KEY);
  } catch (_) {
    /* noop */
  }
}

function setAppBootLoading(isLoading) {
  const active = Boolean(isLoading);
  clearTimeout(bootLoaderCloseTimer);
  if (active) {
    document.body.classList.remove("app-boot-closing");
    document.body.classList.add("app-booting");
  } else {
    if (!document.body.classList.contains("app-booting")) return;
    document.body.classList.remove("app-booting");
    document.body.classList.add("app-boot-closing");
    bootLoaderCloseTimer = setTimeout(() => {
      document.body.classList.remove("app-boot-closing");
    }, 240);
  }
  const loader = document.getElementById("appBootLoader");
  if (loader) loader.setAttribute("aria-hidden", active ? "false" : "true");
}

function hideBootLoaderAfterRender() {
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      setAppBootLoading(false);
    });
  });
}

function setTurboLoaderProgress(progress) {
  if (!turboTopLoaderBar) return;
  const safe = Math.max(0, Math.min(100, Number(progress) || 0));
  turboTopLoaderBar.style.width = `${safe}%`;
}

function startTurboLoader() {
  if (!turboTopLoader || !turboTopLoaderBar) return;
  if (document.body.classList.contains("app-booting")) return;
  clearTimeout(turboLoaderHideTimer);
  turboTopLoader.classList.add("is-active");
  if (turboLoaderProgress <= 0 || turboLoaderProgress >= 100) {
    turboLoaderProgress = 8;
    setTurboLoaderProgress(turboLoaderProgress);
  }
  clearInterval(turboLoaderProgressTimer);
  turboLoaderProgressTimer = setInterval(() => {
    turboLoaderProgress = Math.min(90, turboLoaderProgress + (turboLoaderProgress < 45 ? 12 : turboLoaderProgress < 75 ? 6 : 2));
    setTurboLoaderProgress(turboLoaderProgress);
    if (turboLoaderProgress >= 90) {
      clearInterval(turboLoaderProgressTimer);
      turboLoaderProgressTimer = null;
    }
  }, 90);
}

function finishTurboLoader() {
  if (!turboTopLoader || !turboTopLoaderBar) return;
  clearInterval(turboLoaderProgressTimer);
  turboLoaderProgressTimer = null;
  turboLoaderProgress = 100;
  setTurboLoaderProgress(100);
  clearTimeout(turboLoaderHideTimer);
  turboLoaderHideTimer = setTimeout(() => {
    turboTopLoader.classList.remove("is-active");
    turboLoaderProgress = 0;
    setTurboLoaderProgress(0);
  }, 170);
}

function pulseTurboLoader() {
  startTurboLoader();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      finishTurboLoader();
    });
  });
}

function readLastSessionActivityTs() {
  try {
    const raw = sessionStorage.getItem(SESSION_LAST_ACTIVITY_KEY);
    const n = Number(raw);
    return Number.isFinite(n) && n > 0 ? n : 0;
  } catch (_) {
    return 0;
  }
}

function writeLastSessionActivityTs(ts = Date.now()) {
  try {
    sessionStorage.setItem(SESSION_LAST_ACTIVITY_KEY, String(Math.max(0, Math.floor(ts))));
  } catch (_) {
    /* noop */
  }
}

function clearLastSessionActivityTs() {
  try {
    sessionStorage.removeItem(SESSION_LAST_ACTIVITY_KEY);
  } catch (_) {
    /* noop */
  }
}

function trackSessionActivity(force = false) {
  if (!getAuthToken()) return;
  const now = Date.now();
  if (!force && now - lastSessionActivityWriteAt < 5000) return;
  lastSessionActivityWriteAt = now;
  writeLastSessionActivityTs(now);
}

function bindSessionActivityListeners() {
  if (sessionActivityListenersBound) return;
  sessionActivityListenersBound = true;
  const handler = () => trackSessionActivity(false);
  const opts = { passive: true, capture: true };
  ["pointerdown", "keydown", "touchstart", "wheel", "scroll"].forEach((eventName) => {
    window.addEventListener(eventName, handler, opts);
  });
}

function handleIdleSessionExpired() {
  if (!getAuthToken()) return;
  clearSession();
  showLogin();
  showStatusDialog({
    title: "Сессия завершена",
    message: "Вы были неактивны слишком долго. Войдите снова, чтобы продолжить работу и синхронизацию Telegram.",
    type: "error"
  });
}

function checkSessionIdleTimeout() {
  if (!getAuthToken()) return;
  const last = readLastSessionActivityTs();
  if (!last) {
    trackSessionActivity(true);
    return;
  }
  if (Date.now() - last >= SESSION_IDLE_TIMEOUT_MS) {
    handleIdleSessionExpired();
  }
}

function startSessionIdleWatcher() {
  if (!getAuthToken()) return;
  bindSessionActivityListeners();
  trackSessionActivity(true);
  clearInterval(sessionIdleCheckTimer);
  sessionIdleCheckTimer = setInterval(() => {
    checkSessionIdleTimeout();
  }, SESSION_IDLE_CHECK_INTERVAL_MS);
}

function stopSessionIdleWatcher() {
  clearInterval(sessionIdleCheckTimer);
  sessionIdleCheckTimer = null;
}

function buildAppPayload() {
  normalizeTaskMultiStateStore();
  return {
    sections: JSON.parse(JSON.stringify(sections)),
    displaySettings: JSON.parse(JSON.stringify(displaySettings)),
    trashBySection: JSON.parse(JSON.stringify(trashBySection)),
    taskHistory: loadTaskHistoryStore(),
    taskMultiState: JSON.parse(JSON.stringify(taskMultiState)),
    reportShares: loadReportShares(),
    reportChartOrder: loadReportChartOrder(),
    reportPhaseLayout: loadReportPhaseGroupLayout()
  };
}

let serverSyncTimer = null;
let remotePullTimer = null;
let overdueNotifyTimer = null;
let overdueNotifyInFlight = false;
let hasUnsyncedLocalChanges = false;
let serverPushInFlight = false;
let authExpiredNoticeShown = false;
let bootLoaderCloseTimer = null;
let sessionIdleCheckTimer = null;
let sessionActivityListenersBound = false;
let lastSessionActivityWriteAt = 0;
let turboLoaderProgress = 0;
let turboLoaderProgressTimer = null;
let turboLoaderHideTimer = null;

function handleServerAuthExpired() {
  if (authExpiredNoticeShown) return;
  authExpiredNoticeShown = true;
  setAuthToken("");
  clearSession();
  showLogin();
  showStatusDialog({
    title: "Сессия истекла",
    message: "Повторно войдите в систему. Без этого ответы из Telegram не синхронизируются в таблицу.",
    type: "error"
  });
}

function isKnownSectionId(sectionId) {
  const id = String(sectionId || "").trim();
  if (!id) return false;
  if (id === "tasks" || id === "report" || id === "otherSettings") return true;
  return sections.some((section) => section.id === id);
}

function saveActiveSection(sectionId) {
  const id = String(sectionId || "").trim();
  if (!isKnownSectionId(id)) return;
  try {
    localStorage.setItem(ACTIVE_SECTION_STORAGE_KEY, id);
  } catch (_) {
    /* noop */
  }
}

function restoreActiveSection() {
  try {
    const saved = String(localStorage.getItem(ACTIVE_SECTION_STORAGE_KEY) || "").trim();
    return isKnownSectionId(saved) ? saved : "tasks";
  } catch (_) {
    return "tasks";
  }
}

async function refreshCurrentViewData() {
  if (isHostedRuntime() && getAuthToken()) {
    try {
      await pullRemoteAppState({ rerender: true });
      return;
    } catch (_) {
      /* noop */
    }
  }
  renderTablePreserveScroll();
}

function scheduleServerSync() {
  if (!isHostedRuntime() || !getAuthToken()) return;
  hasUnsyncedLocalChanges = true;
  clearTimeout(serverSyncTimer);
  serverSyncTimer = setTimeout(() => {
    pushAppToServer().catch(() => {});
  }, 1200);
}

async function pushAppToServer() {
  if (!isHostedRuntime() || !getAuthToken()) return;
  serverPushInFlight = true;
  try {
    const data = buildAppPayload();
    await mergeTaskReadStateFromServer(data);
    const r = await fetch("/api/data", {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${getAuthToken()}`
      },
      body: JSON.stringify({ data })
    });
    if (r.status === 401) {
      handleServerAuthExpired();
      return;
    }
    if (r.ok) {
      hasUnsyncedLocalChanges = false;
    }
  } finally {
    serverPushInFlight = false;
  }
}

/** Немедленная синхронизация с сервером (без debounce), например перед регистрацией Telegram webhook. */
async function pushAppToServerImmediate() {
  if (!isHostedRuntime() || !getAuthToken()) return false;
  hasUnsyncedLocalChanges = true;
  serverPushInFlight = true;
  try {
    const data = buildAppPayload();
    await mergeTaskReadStateFromServer(data);
    const r = await fetch("/api/data", {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${getAuthToken()}`
      },
      body: JSON.stringify({ data })
    });
    if (r.status === 401) {
      handleServerAuthExpired();
      return false;
    }
    if (r.ok) {
      hasUnsyncedLocalChanges = false;
    }
    return r.ok;
  } finally {
    serverPushInFlight = false;
  }
}

function mergeBotManagedTaskFieldsIntoLocalRow(localRow, remoteRow) {
  if (!Array.isArray(localRow) || !Array.isArray(remoteRow)) return false;
  let changed = false;

  // Статус и комментарий сотрудника приходят из Telegram и должны иметь приоритет над устаревшей локальной копией.
  const remoteStatus = normalizeTaskStatusValue(remoteRow[TASK_COLUMNS.status]);
  const localStatus = normalizeTaskStatusValue(localRow[TASK_COLUMNS.status]);
  if (String(remoteStatus || "") !== String(localStatus || "")) {
    localRow[TASK_COLUMNS.status] = remoteStatus;
    changed = true;
  }

  const syncTextField = (colIndex) => {
    const remoteVal = String(remoteRow[colIndex] ?? "");
    const localVal = String(localRow[colIndex] ?? "");
    if (remoteVal !== localVal) {
      localRow[colIndex] = remoteVal;
      changed = true;
    }
  };

  syncTextField(TASK_COLUMNS.plan);
  syncTextField(TASK_COLUMNS.mediaAfter);
  syncTextField(TASK_COLUMNS.closedDate);

  // Поле «Прочитано» обновляем безопасно: всегда берём серверный признак прочтения, если он true.
  const localRead = getTaskReadStateParts(localRow[TASK_COLUMNS.readState]);
  const remoteRead = getTaskReadStateParts(remoteRow[TASK_COLUMNS.readState]);
  if (remoteRead.isRead && !localRead.isRead) {
    localRow[TASK_COLUMNS.readState] = String(remoteRow[TASK_COLUMNS.readState] || composeTaskReadState(true, "—"));
    changed = true;
  } else if (String(localRow[TASK_COLUMNS.readState] ?? "") !== String(remoteRow[TASK_COLUMNS.readState] ?? "")) {
    localRow[TASK_COLUMNS.readState] = String(remoteRow[TASK_COLUMNS.readState] ?? "");
    changed = true;
  }

  const localSent = String(localRow[TASK_COLUMNS.lastSentAt] || "").trim();
  const remoteSent = String(remoteRow[TASK_COLUMNS.lastSentAt] || "").trim();
  if ((!localSent || localSent === "—") && remoteSent && remoteSent !== "—") {
    localRow[TASK_COLUMNS.lastSentAt] = remoteSent;
    changed = true;
  } else if (localSent !== remoteSent && remoteSent && remoteSent !== "—") {
    localRow[TASK_COLUMNS.lastSentAt] = remoteSent;
    changed = true;
  }

  return changed;
}

async function mergeTaskReadStateFromServer(localPayload) {
  if (!isHostedRuntime() || !getAuthToken()) return;
  if (!localPayload || typeof localPayload !== "object") return;
  if (!Array.isArray(localPayload.sections)) return;
  try {
    const r = await fetch("/api/data", {
      headers: { Authorization: `Bearer ${getAuthToken()}` }
    });
    if (!r.ok) return;
    const json = await r.json().catch(() => null);
    const remotePayload = json?.data;
    const localTasks = Array.isArray(localPayload.sections)
      ? localPayload.sections.find((s) => s?.id === "tasks")
      : null;
    const remoteTasks = Array.isArray(remotePayload?.sections)
      ? remotePayload.sections.find((s) => s?.id === "tasks")
      : null;
    if (!Array.isArray(localTasks?.rows) || !Array.isArray(remoteTasks?.rows)) return;
    const remoteById = new Map();
    remoteTasks.rows.forEach((row) => {
      const id = String(row?.[TASK_COLUMNS.number] || "").trim();
      if (id) remoteById.set(id, row);
    });
    localTasks.rows.forEach((row) => {
      const id = String(row?.[TASK_COLUMNS.number] || "").trim();
      if (!id) return;
      const remoteRow = remoteById.get(id);
      if (!remoteRow) return;
      mergeBotManagedTaskFieldsIntoLocalRow(row, remoteRow);
    });
    if (remotePayload?.taskMultiState && typeof remotePayload.taskMultiState === "object" && !Array.isArray(remotePayload.taskMultiState)) {
      localPayload.taskMultiState = JSON.parse(JSON.stringify(remotePayload.taskMultiState));
    }
  } catch (_) {
    /* noop */
  }
}

async function pullTaskReadStateFromServerIntoLocal({ rerender = true } = {}) {
  if (!isHostedRuntime() || !getAuthToken()) return false;
  try {
    const r = await fetch("/api/data", {
      headers: { Authorization: `Bearer ${getAuthToken()}` }
    });
    if (r.status === 401) {
      handleServerAuthExpired();
      return false;
    }
    if (!r.ok) return false;
    const json = await r.json().catch(() => null);
    const remoteTasks = Array.isArray(json?.data?.sections)
      ? json.data.sections.find((s) => s?.id === "tasks")
      : null;
    const localTasks = getSectionById("tasks");
    if (!Array.isArray(remoteTasks?.rows) || !Array.isArray(localTasks?.rows)) return false;
    const remoteById = new Map();
    remoteTasks.rows.forEach((row) => {
      const id = String(row?.[TASK_COLUMNS.number] || "").trim();
      if (!id) return;
      remoteById.set(id, row);
    });
    let changed = false;
    localTasks.rows.forEach((row) => {
      const id = String(row?.[TASK_COLUMNS.number] || "").trim();
      if (!id) return;
      const remoteRow = remoteById.get(id);
      if (!remoteRow) return;
      changed = mergeBotManagedTaskFieldsIntoLocalRow(row, remoteRow) || changed;
    });
    let multiChanged = false;
    try {
      if (json?.data?.taskMultiState && typeof json.data.taskMultiState === "object" && !Array.isArray(json.data.taskMultiState)) {
        const nextMulti = JSON.stringify(json.data.taskMultiState);
        const prevMulti = JSON.stringify(taskMultiState || {});
        if (nextMulti !== prevMulti) {
          taskMultiState = JSON.parse(nextMulti);
          localStorage.setItem(TASK_MULTI_STATE_STORAGE_KEY, JSON.stringify(taskMultiState));
          multiChanged = true;
        }
      }
    } catch (_) {
      /* noop */
    }
    if (changed) {
      try {
        localStorage.setItem(DATA_STORAGE_KEY, JSON.stringify(sections));
      } catch (_) {
        /* noop */
      }
    }
    if ((changed || multiChanged) && rerender) renderTablePreserveScroll();
    return changed;
  } catch (_) {
    return false;
  }
}

function telegramBotDisplayNameFromGetMeResult(result) {
  if (!result || typeof result !== "object") return "";
  const fn = String(result.first_name || "").trim();
  const ln = String(result.last_name || "").trim();
  return [fn, ln].filter(Boolean).join(" ").trim();
}

function updateTelegramBotProfileReadonlyDom() {
  const uEl = document.getElementById("telegramBotUsernameReadonly");
  const nEl = document.getElementById("telegramBotDisplayNameReadonly");
  const u = String(displaySettings.telegramBotUsername || "").trim();
  const n = String(displaySettings.telegramBotDisplayName || "").trim();
  if (uEl) uEl.value = u ? `@${u}` : "";
  if (nEl) nEl.value = n || "";
}

/**
 * getMe по токену → displaySettings и поля только для чтения (ник и имя бота в Telegram).
 * @returns {Promise<{ ok: boolean, soft?: boolean, description?: string }>}
 */
async function refreshTelegramBotProfileFromToken(token) {
  const tok = String(token || "").trim();
  if (!tok) {
    displaySettings.telegramBotUsername = "";
    displaySettings.telegramBotDisplayName = "";
    updateTelegramBotProfileReadonlyDom();
    return { ok: false };
  }
  try {
    const r = await fetch(`https://api.telegram.org/bot${tok}/getMe`);
    const j = await r.json().catch(() => ({}));
    if (!j?.ok) {
      displaySettings.telegramBotUsername = "";
      displaySettings.telegramBotDisplayName = "";
      saveDisplaySettings({ skipServerSync: true });
      updateTelegramBotProfileReadonlyDom();
      return { ok: false, description: String(j?.description || "").trim() };
    }
    const res = j.result;
    displaySettings.telegramBotUsername = res?.username ? String(res.username) : "";
    displaySettings.telegramBotDisplayName = telegramBotDisplayNameFromGetMeResult(res);
    saveDisplaySettings({ skipServerSync: true });
    updateTelegramBotProfileReadonlyDom();
    return { ok: true };
  } catch (_) {
    updateTelegramBotProfileReadonlyDom();
    return { ok: false, soft: true };
  }
}

/**
 * Сохраняет токен в БД и вызывает setWebhook на текущем домене (нужен вход в приложение).
 * @param {{ skipPush?: boolean }} [options] — если skipPush, PUT /api/data уже выполнен вызывающим кодом.
 * @returns {Promise<{ ok: boolean, webhookUrl?: string, botUsername?: string, botDisplayName?: string, error?: string }>}
 */
async function registerTelegramWebhookOnServer(options = {}) {
  const skipPush = Boolean(options.skipPush);
  if (!isHostedRuntime() || !getAuthToken()) {
    return { ok: false, error: "Войдите в систему на хостинге (Railway), чтобы подключить webhook." };
  }
  if (!skipPush) {
    const synced = await pushAppToServerImmediate();
    if (!synced) {
      return { ok: false, error: "Не удалось сохранить данные на сервер перед регистрацией бота." };
    }
  }
  const botTok = String(displaySettings.telegramBotToken || "").trim();
  if (!botTok) {
    return {
      ok: true,
      webhookUrl: "",
      botUsername: displaySettings.telegramBotUsername || "",
      botDisplayName: displaySettings.telegramBotDisplayName || ""
    };
  }
  try {
    const r = await fetch("/api/telegram/set-webhook", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${getAuthToken()}`
      },
      body: JSON.stringify({})
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      return { ok: false, error: j.error || `Ошибка ${r.status}` };
    }
    if (j.botUsername) {
      displaySettings.telegramBotUsername = String(j.botUsername);
    }
    if (typeof j.botDisplayName === "string") {
      displaySettings.telegramBotDisplayName = String(j.botDisplayName).trim();
    }
    return {
      ok: true,
      webhookUrl: j.webhookUrl,
      botUsername: j.botUsername,
      botDisplayName: j.botDisplayName
    };
  } catch (e) {
    return { ok: false, error: String(e?.message || e) };
  }
}

/**
 * Подтягивает токен из поля ввода (если оно есть), сбрасывает отложенную синхронизацию,
 * сразу отправляет полный payload на сервер и регистрирует webhook — чтобы при смене токена на сервере не оставался старый снимок из debounce.
 */
async function flushTelegramBotTokenToServer(options = {}) {
  const silent = Boolean(options.silent);
  clearTimeout(serverSyncTimer);
  serverSyncTimer = null;
  const inp = document.getElementById("telegramBotTokenInput");
  if (inp) {
    displaySettings.telegramBotToken = String(inp.value || "").trim();
  }
  saveDisplaySettings({ skipServerSync: true });
  if (!isHostedRuntime() || !getAuthToken()) {
    const tokLocal = String(displaySettings.telegramBotToken || "").trim();
    if (tokLocal) {
      await refreshTelegramBotProfileFromToken(tokLocal);
    } else {
      displaySettings.telegramBotUsername = "";
      displaySettings.telegramBotDisplayName = "";
      updateTelegramBotProfileReadonlyDom();
    }
    saveDisplaySettings();
    return { ok: false, error: "Нет входа в систему на сервере — токен только в браузере." };
  }
  const synced = await pushAppToServerImmediate();
  if (!synced) {
    saveDisplaySettings();
    return { ok: false, error: "Не удалось сохранить данные на сервер." };
  }
  const tok = String(displaySettings.telegramBotToken || "").trim();
  let reg = {
    ok: true,
    webhookUrl: "",
    botUsername: displaySettings.telegramBotUsername || "",
    botDisplayName: displaySettings.telegramBotDisplayName || ""
  };
  if (tok) {
    reg = await registerTelegramWebhookOnServer({ skipPush: true });
    await refreshTelegramBotProfileFromToken(tok);
  } else {
    displaySettings.telegramBotUsername = "";
    displaySettings.telegramBotDisplayName = "";
    updateTelegramBotProfileReadonlyDom();
  }
  saveDisplaySettings();
  if (!silent && !reg.ok && reg.error) {
    window.alert(`Токен записан на сервер, но webhook: ${reg.error}`);
  }
  return reg;
}

async function pullRemoteAppState(options = {}) {
  const rerender = Boolean(options.rerender);
  if (!isHostedRuntime() || !getAuthToken()) return;
  const r = await fetch("/api/data", {
    headers: { Authorization: `Bearer ${getAuthToken()}` }
  });
  if (r.status === 401) {
    handleServerAuthExpired();
    stopRemoteAutoPull();
    return;
  }
  if (!r.ok) return;
  const json = await r.json();
  applyServerBundle(json.data, { rerender });
}

function applyServerBundle(data, options = {}) {
  const rerender = Boolean(options.rerender);
  if (!data || typeof data !== "object") return;
  if (data.displaySettings && typeof data.displaySettings === "object") {
    try {
      localStorage.setItem(DISPLAY_SETTINGS_KEY, JSON.stringify(data.displaySettings));
      restoreDisplaySettings();
      startOverdueTaskNotificationsScheduler();
    } catch (_) {
      /* noop */
    }
  }
  if (Array.isArray(data.sections)) {
    try {
      localStorage.setItem(DATA_STORAGE_KEY, JSON.stringify(data.sections));
      restoreSectionsData();
      loadObjectPhotoThumbsFromStorage();
    } catch (_) {
      /* noop */
    }
  }
  if (data.trashBySection && typeof data.trashBySection === "object") {
    try {
      localStorage.setItem(TRASH_STORAGE_KEY, JSON.stringify(data.trashBySection));
      restoreTrashData();
      clearTasksTrashNow({ save: true });
    } catch (_) {
      /* noop */
    }
  }
  if (data.taskHistory && typeof data.taskHistory === "object") {
    try {
      localStorage.setItem(TASK_HISTORY_STORAGE_KEY, JSON.stringify(data.taskHistory));
    } catch (_) {
      /* noop */
    }
  }
  if (data.taskMultiState && typeof data.taskMultiState === "object" && !Array.isArray(data.taskMultiState)) {
    try {
      localStorage.setItem(TASK_MULTI_STATE_STORAGE_KEY, JSON.stringify(data.taskMultiState));
      restoreTaskMultiState();
    } catch (_) {
      /* noop */
    }
  }
  if (data.reportShares && typeof data.reportShares === "object") {
    try {
      localStorage.setItem(REPORT_SHARE_STORAGE_KEY, JSON.stringify(data.reportShares));
    } catch (_) {
      /* noop */
    }
  }
  if (Array.isArray(data.reportChartOrder)) {
    try {
      localStorage.setItem(REPORT_CHART_ORDER_STORAGE_KEY, JSON.stringify(data.reportChartOrder));
    } catch (_) {
      /* noop */
    }
  }
  if (data.reportPhaseLayout === "row" || data.reportPhaseLayout === "separate") {
    try {
      localStorage.setItem(REPORT_PHASE_GROUP_LAYOUT_KEY, data.reportPhaseLayout);
    } catch (_) {
      /* noop */
    }
  }
  if (rerender) {
    renderTablePreserveScroll();
  }
}

function isUserEditingNow() {
  const el = document.activeElement;
  if (!el) return false;
  if (el.isContentEditable) return true;
  const tag = String(el.tagName || "").toUpperCase();
  return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT";
}

function startRemoteAutoPull() {
  clearInterval(remotePullTimer);
  remotePullTimer = null;
  if (!isHostedRuntime() || !getAuthToken()) return;
  remotePullTimer = setInterval(() => {
    if (document.hidden) return;
    // Bot-managed поля задач (прочитано/статус/комментарий и т.п.) подтягиваем всегда,
    // даже если пользователь сейчас в настройках или есть локальные несинхр. правки.
    pullTaskReadStateFromServerIntoLocal({ rerender: activeSectionId === "tasks" }).catch(() => {});

    // Полный pull — только на экране задач и когда пользователь не редактирует,
    // чтобы не мешать вводу в формах.
    if (activeSectionId !== "tasks") return;
    if (isUserEditingNow()) return;
    if (serverPushInFlight || hasUnsyncedLocalChanges) return;
    pullRemoteAppState({ rerender: true }).catch(() => {});
  }, 8000);
  pullTaskReadStateFromServerIntoLocal({ rerender: activeSectionId === "tasks" }).catch(() => {});
}

function stopRemoteAutoPull() {
  clearInterval(remotePullTimer);
  remotePullTimer = null;
}

function escapeHtmlText(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeHtmlAttr(s) {
  return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function getSessionUserDisplayName() {
  return String(localStorage.getItem(SESSION_STORAGE_KEY) || "").trim() || "Пользователь";
}

function canAccessSettingsMenu() {
  if (!isHostedRuntime() || !getAuthToken()) return true;
  return currentAuthRole === "admin";
}

async function refreshAuthMeProfile() {
  if (!isHostedRuntime() || !getAuthToken()) return null;
  try {
    const r = await fetch("/api/auth/me", {
      headers: { Authorization: `Bearer ${getAuthToken()}` }
    });
    if (r.status === 401) {
      handleServerAuthExpired();
      return null;
    }
    if (!r.ok) return null;
    const j = await r.json().catch(() => ({}));
    const role = String(j?.role || "user").trim().toLowerCase() === "admin" ? "admin" : "user";
    const displayName = String(j?.displayName || "").trim();
    currentAuthRole = role;
    return { role, displayName };
  } catch (_) {
    return null;
  }
}

function loadTaskHistoryStore() {
  try {
    const raw = localStorage.getItem(TASK_HISTORY_STORAGE_KEY);
    if (!raw) return {};
    const p = JSON.parse(raw);
    return typeof p === "object" && p && !Array.isArray(p) ? p : {};
  } catch (_) {
    return {};
  }
}

function saveTaskHistoryStore(store) {
  try {
    localStorage.setItem(TASK_HISTORY_STORAGE_KEY, JSON.stringify(store));
    scheduleServerSync();
  } catch (_) {
    /* noop */
  }
}

function formatTaskHistoryWhen(ms) {
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return "—";
  const dateStr = d.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit", year: "numeric" });
  const use12 = normalizeTimeDisplayFormatId(displaySettings.timeDisplayFormat) === "12";
  const showSec = Boolean(displaySettings.timeShowSeconds);
  const timeStr = d.toLocaleTimeString("ru-RU", {
    hour: "2-digit",
    minute: "2-digit",
    second: showSec ? "2-digit" : undefined,
    hour12: use12
  });
  return `${dateStr} ${timeStr}`;
}

function appendTaskHistoryEntry(taskId, actionText) {
  const id = String(taskId ?? "").trim() || "—";
  const who = getSessionUserDisplayName();
  const action = String(actionText || "").trim();
  if (!action) return;
  const store = loadTaskHistoryStore();
  if (!store[id]) store[id] = [];
  store[id].unshift({ t: Date.now(), who, action });
  if (store[id].length > TASK_HISTORY_MAX_PER_TASK) {
    store[id].length = TASK_HISTORY_MAX_PER_TASK;
  }
  saveTaskHistoryStore(store);
}

function getTaskHistoryEntries(taskId) {
  const id = String(taskId ?? "").trim();
  const store = loadTaskHistoryStore();
  const list = store[id];
  return Array.isArray(list) ? list : [];
}

function renderTaskHistoryTableHtml(taskId) {
  const entries = getTaskHistoryEntries(taskId);
  if (!entries.length) {
    return `<p class="task-history-empty">Пока нет записей. Изменения в таблице задач и сохранение карточки будут отображаться здесь.</p>`;
  }
  const rows = entries
    .map(
      (e) => `
    <tr>
      <td class="task-history-when">${escapeHtmlText(formatTaskHistoryWhen(e.t))}</td>
      <td class="task-history-who">${escapeHtmlText(e.who || "—")}</td>
      <td class="task-history-action">${escapeHtmlText(e.action)}</td>
    </tr>`
    )
    .join("");
  return `
    <div class="task-history-scroll">
      <table class="task-history-table">
        <thead>
          <tr>
            <th>Дата и время</th>
            <th>Кто</th>
            <th>Действие</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function taskHistoryCtx(section, rowIndex, colIndex) {
  if (section.id !== "tasks") return undefined;
  const row = section.rows[rowIndex];
  if (!row) return undefined;
  return {
    taskId: String(row[TASK_COLUMNS.number] ?? ""),
    columnLabel: String(section.columns[colIndex] ?? ""),
    getOld: () => row[colIndex]
  };
}

function shortenHistorySnippet(s, maxLen = 100) {
  const t = String(s ?? "");
  if (t.length <= maxLen) return t;
  return `${t.slice(0, maxLen)}…`;
}

/** Подставляет значения строки задачи в шаблон с токенами вида [ид_задачи] */
function applyTaskMessageTemplate(template, row) {
  if (!template || !row) return "";
  let out = String(template);
  const formatTokenValue = (colKey, rawValue) => {
    if (colKey === "status") {
      const st = String(rawValue ?? "").trim();
      if (!st) return "";
      return `${TELEGRAM_STATUS_EMOJI[st] || "⚪"} ${st}`;
    }
    if (colKey !== "mediaBefore" && colKey !== "mediaAfter") {
      return String(rawValue ?? "");
    }
    const items = getMediaItems(rawValue);
    if (!items.length) return "Нет";
    return `Фото приложено: ${items.length}`;
  };
  for (const item of TASK_MESSAGE_PLACEHOLDERS) {
    const col = TASK_COLUMNS[item.col];
    if (col === undefined) continue;
    const val = formatTokenValue(item.col, row[col]);
    out = out.split(item.token).join(val);
  }
  return out.replace(/\r\n?/g, "\n");
}

function normalizePersonName(s) {
  return String(s ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

/** Строка объекта по наименованию из задачи: РП и ЗРП с этой же строки (не путать с однофамильцами на других объектах). */
function getObjectRpZrpForTask(taskRow) {
  const objectName = normalizePersonName(taskRow[TASK_COLUMNS.object]);
  if (!objectName) return null;
  const objectsRows = getSectionById("objects")?.rows || [];
  for (const objRow of objectsRows) {
    const n = normalizePersonName(objRow[OBJECT_COLUMNS.name]);
    if (n && n === objectName) {
      return {
        rp: normalizePersonName(objRow[OBJECT_COLUMNS.rp]),
        zrp: normalizePersonName(objRow[OBJECT_COLUMNS.zrp])
      };
    }
  }
  return null;
}

function addGlobalDuplicateRecipientNames(namesSet) {
  const ids = Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
    ? displaySettings.telegramGlobalDuplicateRecipientIds
    : [];
  if (!ids.length) return;
  const idSet = new Set(ids.map((x) => String(x).trim()).filter(Boolean));
  const empRows = getSectionById("employees")?.rows || [];
  for (const er of empRows) {
    const eid = String(er[EMPLOYEE_COLUMNS.id] ?? "").trim();
    if (!eid || !idSet.has(eid)) continue;
    const fn = normalizePersonName(er[EMPLOYEE_COLUMNS.fullName]);
    if (fn) namesSet.add(fn);
  }
}

function getDepartmentHeadNameByEmployeeName(employeeFullName) {
  const empName = normalizePersonName(employeeFullName);
  if (!empName) return "";
  const empRows = getSectionById("employees")?.rows || [];
  const empRow = empRows.find((r) => normalizePersonName(r[EMPLOYEE_COLUMNS.fullName]) === empName);
  if (!empRow) return "";
  const departmentName = normalizePersonName(empRow[EMPLOYEE_COLUMNS.department]);
  if (!departmentName) return "";
  const depRows = getSectionById("departments")?.rows || [];
  const depRow = depRows.find((r) => normalizePersonName(r[1]) === departmentName);
  if (!depRow) return "";
  const headName = normalizePersonName(depRow[2]);
  return headName || "";
}

function getEmployeeRowsByDisplayName(fullName) {
  const want = normalizePersonName(fullName);
  if (!want) return [];
  const rows = getSectionById("employees")?.rows || [];
  return rows.filter((r) => normalizePersonName(r[EMPLOYEE_COLUMNS.fullName]) === want);
}

function employeeHasTelegramBinding(row) {
  if (!row) return false;
  const connected = String(row[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен";
  const chat = String(row[EMPLOYEE_COLUMNS.chatId] || "").trim();
  return connected && Boolean(chat);
}

function collectTaskCloseApproverNames(taskRow) {
  const approvers = new Map();
  const add = (name, reason) => {
    const n = normalizePersonName(name);
    if (!n) return;
    if (!approvers.has(n)) approvers.set(n, new Set());
    approvers.get(n).add(reason);
  };

  // 1) Руководитель отдела исполнителя.
  parseTaskAssigneeNames(taskRow[TASK_COLUMNS.assignedResponsible]).forEach((assigneeName) => {
    const headName = getDepartmentHeadNameByEmployeeName(assigneeName);
    if (!headName) return;
    const headRows = getEmployeeRowsByDisplayName(headName);
    if (headRows.some((r) => employeeHasTelegramBinding(r))) {
      add(headName, "руководитель отдела");
    }
  });

  // 2) Админ и директор (по должности), если подключены в Telegram.
  const alwaysPositions = new Set(["Администратор", "Генеральный директор"]);
  const allEmployees = getSectionById("employees")?.rows || [];
  for (const row of allEmployees) {
    const pos = String(row[EMPLOYEE_COLUMNS.position] || "").trim();
    if (!alwaysPositions.has(pos)) continue;
    if (!employeeHasTelegramBinding(row)) continue;
    add(row[EMPLOYEE_COLUMNS.fullName], pos.toLowerCase());
  }

  // 3) Глобальные согласующие из настроек (пересечение: дубликаты + разрешено подтверждать).
  const dupIds = new Set(
    Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
      ? displaySettings.telegramGlobalDuplicateRecipientIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  const allowIds = new Set(
    Array.isArray(displaySettings.telegramCloseConfirmAllowedIds)
      ? displaySettings.telegramCloseConfirmAllowedIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  for (const row of allEmployees) {
    const id = String(row[EMPLOYEE_COLUMNS.id] ?? "").trim();
    if (!id || !dupIds.has(id) || !allowIds.has(id)) continue;
    if (!employeeHasTelegramBinding(row)) continue;
    add(row[EMPLOYEE_COLUMNS.fullName], "глобальный согласующий");
  }

  return Array.from(approvers.entries()).map(([name, reasons]) => ({
    name,
    reason: Array.from(reasons).join(", ")
  }));
}

/**
 * Имена получателей: исполнитель и контролирующий (все из списка ответственных по задаче получают сообщение).
 * Дублирование РП↔ЗРП только если в этих полях указан РП или ЗРП с карточки объекта задачи (остальные ФИО не вызывают эту пару).
 * Плюс глобальные получатели копий (по ID в настройках) — на все задачи.
 */
function collectTaskTelegramRecipientNames(taskRow) {
  const base = new Set();
  const addBase = (v) => {
    const n = normalizePersonName(v);
    if (n) base.add(n);
  };
  parseTaskAssigneeNames(taskRow[TASK_COLUMNS.assignedResponsible]).forEach((name) => addBase(name));
  addBase(taskRow[TASK_COLUMNS.responsible]);
  parseTaskAssigneeNames(taskRow[TASK_COLUMNS.assignedResponsible]).forEach((name) => {
    const departmentHead = getDepartmentHeadNameByEmployeeName(name);
    if (departmentHead) base.add(departmentHead);
  });

  const names = new Set(base);
  const oz = getObjectRpZrpForTask(taskRow);
  if (oz && oz.rp && oz.zrp) {
    for (const n of base) {
      if (n === oz.rp) names.add(oz.zrp);
      if (n === oz.zrp) names.add(oz.rp);
    }
  }

  addGlobalDuplicateRecipientNames(names);
  return Array.from(names);
}

/** Inline-кнопки у сообщения задачи (обрабатываются сервером /api/telegram/webhook). */
function buildTelegramInlineKeyboardForTask(taskNumber) {
  const n = String(taskNumber ?? "")
    .trim()
    .replace(/\|/g, "·")
    .slice(0, 48);
  const makeCb = (code) => {
    let s = `t|${n}|${code}`;
    if (s.length > 64) s = s.slice(0, 64);
    return s;
  };
  return {
    inline_keyboard: [
      [
        { text: "⌛️ Сменить статус", callback_data: makeCb("sm") },
        { text: "🗣 Комментарий", callback_data: makeCb("cm") }
      ],
      [{ text: "📸 Отправить фото", callback_data: makeCb("ph") }]
    ]
  };
}

function buildTelegramReadInlineKeyboardForTask(taskNumber) {
  const n = String(taskNumber ?? "")
    .trim()
    .replace(/\|/g, "·")
    .slice(0, 48);
  const makeCb = (code) => {
    let s = `t|${n}|${code}`;
    if (s.length > 64) s = s.slice(0, 64);
    return s;
  };
  return {
    inline_keyboard: [[{ text: "📖 Прочитать", callback_data: makeCb("rd") }]]
  };
}

/** Пример строки задачи для предпросмотра шаблонов в настройках (не влияет на данные). */
function buildDemoTaskRowForPreview(status) {
  const row = new Array(20).fill("");
  const st = String(status || "").trim() || "Новый";
  row[TASK_COLUMNS.number] = "42";
  row[TASK_COLUMNS.object] = "Офис Центр";
  row[TASK_COLUMNS.status] = st;
  row[TASK_COLUMNS.priority] = "Средний";
  row[TASK_COLUMNS.addedDate] = "16.04.2026";
  row[TASK_COLUMNS.phase] = "Реализация";
  row[TASK_COLUMNS.phaseSection] = "СМР";
  row[TASK_COLUMNS.phaseSubsection] = "Надземные конструкции";
  row[TASK_COLUMNS.assignedResponsible] = "Иван Петров";
  row[TASK_COLUMNS.task] = "Проверить освещение в коридоре";
  row[TASK_COLUMNS.responsible] = "Мария Волкова";
  row[TASK_COLUMNS.note] = "Согласовать с заказчиком";
  row[TASK_COLUMNS.plan] = "2 чел.-дня";
  row[TASK_COLUMNS.fact] = "1 чел.-день";
  row[TASK_COLUMNS.dueDate] = "20.04.2026";
  row[TASK_COLUMNS.closedDate] = st === "Закрыт" ? "16.04.2026" : "";
  row[TASK_COLUMNS.mediaBefore] = "";
  row[TASK_COLUMNS.mediaAfter] = "";
  row[TASK_COLUMNS.readState] = "Не прочитано\n—";
  row[TASK_COLUMNS.lastSentAt] = "—";
  return row;
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

function composeTaskReadState(isRead, whenText = "—") {
  return `${isRead ? "Прочитано" : "Не прочитано"}\n${String(whenText || "—").trim() || "—"}`;
}

function formatTelegramPreviewHtml(text) {
  const t = String(text ?? "").trim() || "—";
  return escapeHtmlText(t).replace(/\r\n|\r|\n/g, "<br />");
}

function findTaskMessageTemplateTextareaByStatus(status) {
  const want = String(status || "").trim();
  const nodes = document.querySelectorAll(".task-message-template-input");
  for (const n of nodes) {
    if (String(n.dataset.status || "").trim() === want) return n;
  }
  return null;
}

function attachTelegramTemplatePreviews() {
  const emulatorTask = document.getElementById("taskMsgTelegramEmulator");
  const emulatorRem = document.getElementById("reminderTelegramEmulator");
  const bubbleTask = document.getElementById("taskMsgTelegramBubbleText");
  const capTask = document.getElementById("taskMsgTelegramCaption");
  const kbTask = document.getElementById("taskMsgTelegramKeyboard");
  const bubbleRem = document.getElementById("reminderTelegramBubbleText");
  const capRem = document.getElementById("reminderTelegramCaption");

  let taskPreviewMode = "status";
  let taskPreviewStatus = STATUS_OPTIONS[0];
  let reminderPreviewStatus = STATUS_OPTIONS[0];

  const setTaskKeyboardVisible = (show) => {
    if (kbTask) kbTask.classList.toggle("telegram-emulator-keyboard--hidden", !show);
  };

  const refreshTaskPreview = () => {
    if (!bubbleTask || !emulatorTask) return;
    if (taskPreviewMode === "close") {
      const inp = document.getElementById("telegramCloseAcceptedInput");
      const tpl = inp ? String(inp.value || "") : "";
      const row = buildDemoTaskRowForPreview("Закрыт");
      bubbleTask.innerHTML = formatTelegramPreviewHtml(applyTaskMessageTemplate(tpl, row));
      if (capTask) capTask.textContent = "После подтверждения закрытия";
      setTaskKeyboardVisible(false);
      return;
    }
    const ta = findTaskMessageTemplateTextareaByStatus(taskPreviewStatus);
    const tpl = ta ? String(ta.value || "") : "";
    const row = buildDemoTaskRowForPreview(taskPreviewStatus);
    bubbleTask.innerHTML = formatTelegramPreviewHtml(applyTaskMessageTemplate(tpl, row));
    if (capTask) capTask.textContent = "Сообщение по задаче (с кнопками)";
    setTaskKeyboardVisible(true);
  };

  const refreshReminderPreview = () => {
    if (!bubbleRem || !emulatorRem) return;
    const inputs = document.querySelectorAll(".reminder-text-input");
    let inp = null;
    for (const el of inputs) {
      if (String(el.dataset.status || "").trim() === reminderPreviewStatus) {
        inp = el;
        break;
      }
    }
    if (!inp) inp = document.querySelector(".reminder-text-input");
    const tpl = inp ? String(inp.value || "") : "";
    const row = buildDemoTaskRowForPreview(reminderPreviewStatus);
    bubbleRem.innerHTML = formatTelegramPreviewHtml(applyTaskMessageTemplate(tpl, row));
    if (capRem) capRem.textContent = "Напоминание (текст без кнопок задачи)";
  };

  document.querySelectorAll(".task-message-template-input").forEach((el) => {
    el.addEventListener("focus", () => {
      taskPreviewMode = "status";
      taskPreviewStatus = String(el.dataset.status || STATUS_OPTIONS[0]).trim() || STATUS_OPTIONS[0];
      refreshTaskPreview();
    });
    el.addEventListener("input", () => {
      taskPreviewMode = "status";
      taskPreviewStatus = String(el.dataset.status || STATUS_OPTIONS[0]).trim() || STATUS_OPTIONS[0];
      refreshTaskPreview();
    });
  });

  const closeInp = document.getElementById("telegramCloseAcceptedInput");
  closeInp?.addEventListener("focus", () => {
    taskPreviewMode = "close";
    refreshTaskPreview();
  });
  closeInp?.addEventListener("input", () => {
    taskPreviewMode = "close";
    refreshTaskPreview();
  });

  document.querySelectorAll(".task-format-block").forEach((block) => {
    block.addEventListener("click", (e) => {
      if (e.target.closest("textarea")) return;
      const ta = block.querySelector(".task-message-template-input");
      if (!ta) return;
      taskPreviewMode = "status";
      taskPreviewStatus = String(ta.dataset.status || STATUS_OPTIONS[0]).trim() || STATUS_OPTIONS[0];
      ta.focus();
      refreshTaskPreview();
    });
  });

  document.querySelectorAll(".reminder-text-input").forEach((el) => {
    el.addEventListener("focus", () => {
      reminderPreviewStatus = String(el.dataset.status || STATUS_OPTIONS[0]).trim() || STATUS_OPTIONS[0];
      refreshReminderPreview();
    });
    el.addEventListener("input", () => {
      reminderPreviewStatus = String(el.dataset.status || STATUS_OPTIONS[0]).trim() || STATUS_OPTIONS[0];
      refreshReminderPreview();
    });
  });

  document.querySelectorAll(".reminder-item").forEach((item) => {
    item.addEventListener("click", (e) => {
      if (e.target.closest("input, textarea, select")) return;
      const inp = item.querySelector(".reminder-text-input");
      if (!inp) return;
      reminderPreviewStatus = String(inp.dataset.status || STATUS_OPTIONS[0]).trim() || STATUS_OPTIONS[0];
      inp.focus();
      refreshReminderPreview();
    });
  });

  refreshTaskPreview();
  refreshReminderPreview();

  window._mbcRefreshTaskFormatPreview = refreshTaskPreview;
  window._mbcRefreshReminderPreview = refreshReminderPreview;
}

function resolveEmployeeTelegramTargetsByFullNames(names) {
  const employeesRows = getSectionById("employees")?.rows || [];
  const want = new Set(names.map(normalizePersonName).filter(Boolean));
  const seenChat = new Set();
  const out = [];
  for (const row of employeesRows) {
    const fn = normalizePersonName(row[EMPLOYEE_COLUMNS.fullName]);
    if (!fn || !want.has(fn)) continue;
    const tg = String(row[EMPLOYEE_COLUMNS.telegram] || "");
    const chatId = String(row[EMPLOYEE_COLUMNS.chatId] || "").trim();
    if (tg !== "Подключен" || !chatId) continue;
    if (seenChat.has(chatId)) continue;
    seenChat.add(chatId);
    out.push({ fullName: fn, chatId });
  }
  return out;
}

function resolveTaskActionChatIds(taskRow) {
  const names = [];
  parseTaskAssigneeNames(taskRow?.[TASK_COLUMNS.assignedResponsible]).forEach((name) => names.push(name));
  let targets = resolveEmployeeTelegramTargetsByFullNames(names);
  if (!targets.length) {
    const responsible = normalizePersonName(taskRow?.[TASK_COLUMNS.responsible]);
    if (responsible) {
      targets = resolveEmployeeTelegramTargetsByFullNames([responsible]);
    }
  }
  return new Set(targets.map((t) => String(t.chatId || "").trim()).filter(Boolean));
}

function getTaskMediaItemsForTelegram(taskRow) {
  const before = getMediaItems(taskRow[TASK_COLUMNS.mediaBefore]);
  return [...before].map((x) => String(x || "").trim()).filter(Boolean);
}

function resolveTelegramSendablePhotoRefs(taskRow, token) {
  const items = getTaskMediaItemsForTelegram(taskRow);
  if (!items.length) return [];
  const toSendable = (raw) => {
    const directUrl = toAbsoluteMediaUrl(raw);
    if (directUrl) return directUrl;
    if (raw.startsWith("/")) {
      const clean = raw.replace(/^\/+/, "");
      return `${location.origin}/${clean}`;
    }
    if (raw.includes("/")) {
      const clean = raw.replace(/^\/+/, "");
      return `https://api.telegram.org/file/bot${token}/${clean}`;
    }
    return "";
  };
  const refs = [];
  for (const raw of items) {
    const displayName = getMediaDisplayName(raw);
    if (!mediaNameLooksLikeImage(displayName)) continue;
    const ref = toSendable(raw);
    if (ref) refs.push(ref);
  }
  return refs.slice(0, 5);
}

async function precheckPhotoRefForTelegram(photoRef) {
  const ref = String(photoRef || "").trim();
  if (!ref) return { ok: false, ref: "" };
  try {
    const u = new URL(ref, location.origin);
    if (u.origin !== location.origin) return { ok: true, ref };
    const r = await fetch(u.toString(), { method: "HEAD", cache: "no-store" });
    if (!r.ok) {
      return { ok: false, ref: "", reason: `HTTP ${r.status}` };
    }
    const ct = String(r.headers.get("content-type") || "").toLowerCase();
    if (ct && !ct.startsWith("image/")) {
      return { ok: false, ref: "", reason: `content-type: ${ct}` };
    }
    return { ok: true, ref: u.toString() };
  } catch (_) {
    return { ok: true, ref };
  }
}

async function sendTelegramPhotoViaServerProxy({ chatId, token, photoRef, caption, replyMarkup }) {
  if (!isHostedRuntime() || !getAuthToken()) {
    return { ok: false, description: "proxy_unavailable" };
  }
  try {
    const r = await fetch("/api/telegram/send-photo-proxy", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${getAuthToken()}`
      },
      body: JSON.stringify({
        chatId: String(chatId || "").trim(),
        token: String(token || "").trim(),
        photoRef: String(photoRef || "").trim(),
        caption: String(caption || ""),
        replyMarkup: replyMarkup || null
      })
    });
    const j = await r.json().catch(() => ({}));
    return {
      ok: r.ok && j?.ok === true,
      description: String(j?.error || j?.description || "")
    };
  } catch (e) {
    return { ok: false, description: String(e?.message || "proxy_failed") };
  }
}

async function sendTelegramMediaGroupViaServerProxy({ chatId, token, photoRefs }) {
  if (!isHostedRuntime() || !getAuthToken()) {
    return { ok: false, description: "proxy_unavailable", firstMessageId: null };
  }
  try {
    const r = await fetch("/api/telegram/send-media-group-proxy", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${getAuthToken()}`
      },
      body: JSON.stringify({
        chatId: String(chatId || "").trim(),
        token: String(token || "").trim(),
        photoRefs: Array.isArray(photoRefs) ? photoRefs.slice(0, 10) : []
      })
    });
    const j = await r.json().catch(() => ({}));
    return {
      ok: r.ok && j?.ok === true,
      description: String(j?.error || j?.description || ""),
      firstMessageId: Number(j?.firstMessageId) || null
    };
  } catch (e) {
    return { ok: false, description: String(e?.message || "proxy_group_failed"), firstMessageId: null };
  }
}

/**
 * Отправка текста задачи в Telegram по шаблону текущего статуса.
 * @returns {Promise<{ ok: boolean, reason?: string, okCount?: number, total?: number, missingNames?: string[] }>}
 */
async function sendTaskRowTelegramNotification(taskRow, options = {}) {
  const suppressAlerts = Boolean(options.suppressAlerts);
  const targetAssigneeNames = Array.isArray(options.targetAssigneeNames)
    ? options.targetAssigneeNames.map((x) => normalizePersonName(x)).filter(Boolean)
    : [];
  const notify = (message, type = "info") => {
    if (suppressAlerts) return;
    showStatusDialog({
      title: "Отправка в Telegram",
      message: String(message || ""),
      type
    });
  };
  const token = String(displaySettings.telegramBotToken || "").trim();
  cleanupTaskMultiStateForRow(taskRow);
  if (!token) {
    const msg = "Укажите токен Telegram-бота в прочих настройках.";
    notify(msg, "error");
    return { ok: false, reason: "no_token", message: msg };
  }

  const status = String(taskRow[TASK_COLUMNS.status] || "").trim();
  const templates = displaySettings.taskMessageTemplatesByStatus || {};
  const template = String(templates[status] || "");
  if (!template.trim()) {
    const msg = `Нет текста шаблона для статуса «${status}». Заполните в «Прочие настройки» → «Шаблон сообщений».`;
    notify(msg, "error");
    return { ok: false, reason: "no_template", message: msg };
  }

  const text = applyTaskMessageTemplate(template, taskRow);
  if (!String(text).trim()) {
    const msg = "После подстановки полей сообщение пустое.";
    notify(msg, "error");
    return { ok: false, reason: "empty_message", message: msg };
  }

  const recipientNames = targetAssigneeNames.length
    ? targetAssigneeNames
    : collectTaskTelegramRecipientNames(taskRow);
  if (!recipientNames.length) {
    const msg =
      "Нет получателей: укажите исполнителя и/или контролирующего ответственного в задаче либо добавьте глобальных получателей копий в настройках Telegram.";
    notify(msg, "error");
    return { ok: false, reason: "no_recipients", message: msg };
  }

  const targets = resolveEmployeeTelegramTargetsByFullNames(recipientNames);
  const actionChatIds = resolveTaskActionChatIds(taskRow);
  const resolvedNames = new Set(targets.map((t) => t.fullName));
  const missingNames = recipientNames.filter((n) => !resolvedNames.has(n));

  if (!targets.length) {
    const msg =
      "Не удалось найти сотрудников с Telegram: проверьте ФИО в задаче и в справочнике сотрудников (подключение Telegram и Chat ID).";
    notify(msg, "error");
    return { ok: false, reason: "no_targets", message: msg, missingNames };
  }

  const photoRefs = resolveTelegramSendablePhotoRefs(taskRow, token);
  const shortText = `У вас есть задача ID ${String(taskRow[TASK_COLUMNS.number] || "").trim() || "—"}.\nЧтобы прочитать полное содержание, нажмите «📖 Прочитать».`;
  const results = await Promise.all(
    targets.map(async (t) => {
      try {
        const isActionRecipient = actionChatIds.has(String(t.chatId || "").trim());
        const outgoingText = String(isActionRecipient ? shortText : text).replace(/\r\n?/g, "\n");
        const replyMarkup = !options.skipInlineKeyboard && isActionRecipient
          ? buildTelegramReadInlineKeyboardForTask(taskRow[TASK_COLUMNS.number])
          : undefined;
        if (photoRefs.length) {
          let photosOk = false;
          let photoDescription = "";
          let firstMessageId = null;

          if (photoRefs.length > 1 && isHostedRuntime() && getAuthToken()) {
            const group = await sendTelegramMediaGroupViaServerProxy({
              chatId: t.chatId,
              token,
              photoRefs
            });
            photosOk = group.ok === true;
            photoDescription = group.description || "";
            firstMessageId = group.firstMessageId || null;
          } else {
            const firstRef = photoRefs[0];
            if (isHostedRuntime() && getAuthToken()) {
              const proxy = await sendTelegramPhotoViaServerProxy({
                chatId: t.chatId,
                token,
                photoRef: firstRef,
                caption: ""
              });
              photosOk = proxy.ok === true;
              photoDescription = proxy.description || "";
            } else {
              const r = await fetch(`https://api.telegram.org/bot${token}/sendPhoto`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  chat_id: t.chatId,
                  photo: firstRef
                })
              });
              const j = await r.json().catch(() => ({}));
              photosOk = r.ok && j.ok === true;
              photoDescription = String(j.description || "");
              firstMessageId = Number(j?.result?.message_id) || null;
            }
          }

          if (!photosOk) {
            const fallbackBody = {
              chat_id: t.chatId,
              text: outgoingText,
              ...(replyMarkup ? { reply_markup: replyMarkup } : {})
            };
            const fallbackResponse = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(fallbackBody)
            });
            const fallbackJson = await fallbackResponse.json().catch(() => ({}));
            return {
              ...t,
              ok: fallbackResponse.ok && fallbackJson.ok === true,
              apiDescription: fallbackJson.description || photoDescription,
              photoFallback: true,
              photoError: photoDescription || "",
              photoMode: "server_proxy"
            };
          }

          const msgBody = {
            chat_id: t.chatId,
            text: outgoingText,
            ...(replyMarkup ? { reply_markup: replyMarkup } : {}),
            ...(firstMessageId ? { reply_to_message_id: firstMessageId, allow_sending_without_reply: true } : {})
          };
          const textResponse = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(msgBody)
          });
          const textJson = await textResponse.json().catch(() => ({}));
          return {
            ...t,
            ok: textResponse.ok && textJson.ok === true,
            apiDescription: textJson.description,
            photoFallback: false,
            photoMode: "server_proxy"
          };
        }

        const msgBody = {
          chat_id: t.chatId,
          text: outgoingText,
          ...(replyMarkup ? { reply_markup: replyMarkup } : {})
        };
        const response = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(msgBody)
        });
        const apiJson = await response.json().catch(() => ({}));
        return {
          ...t,
          ok: response.ok && apiJson.ok === true,
          apiDescription: apiJson.description,
          photoFallback: false,
          photoMode: "none"
        };
      } catch (_) {
        return { ...t, ok: false, photoFallback: false, photoMode: "none" };
      }
    })
  );

  const okCount = results.filter((r) => r.ok).length;
  const ok = okCount > 0;
  if (ok && taskRow) {
    markTaskAsSentImported(taskRow[TASK_COLUMNS.number], { save: false });
    taskRow[TASK_COLUMNS.readState] = composeTaskReadState(false, "—");
    taskRow[TASK_COLUMNS.lastSentAt] = formatTrashDate(Date.now());
    saveSectionsData();
    saveDisplaySettings();
    // Избегаем гонки: отложенный debounce-пуш может перезаписать серверное
    // «Прочитано», если сотрудник нажмёт «📖 Прочитать» сразу после отправки.
    clearTimeout(serverSyncTimer);
    if (isHostedRuntime() && getAuthToken()) {
      try {
        await pushAppToServerImmediate();
      } catch (_) {
        /* noop */
      }
    }
  }

  if (!suppressAlerts) {
    let msg = `Отправлено успешно: ${okCount} из ${results.length}.`;
    const fallbackWithPhotoError = results.find((r) => r.photoFallback === true);
    if (fallbackWithPhotoError) {
      const why = "Telegram отклонил фото при серверной отправке.";
      msg += `\n\nФото не прикрепилось: ${why}${fallbackWithPhotoError.photoError ? `\nПричина Telegram: ${fallbackWithPhotoError.photoError}` : ""}`;
    }
    if (missingNames.length) {
      msg += `\n\nБез доставки (нет подключения Telegram или Chat ID): ${missingNames.join(", ")}.`;
    }
    if (!ok) {
      const firstErr = results.find((r) => !r.ok && r.apiDescription);
      msg = firstErr?.apiDescription
        ? `Не удалось отправить: ${firstErr.apiDescription}\n\nУбедитесь, что в «Chat ID» указан числовой Telegram user ID (например из @userinfobot), а не цифры с номера телефона.`
        : "Не удалось отправить сообщение. Проверьте токен и Chat ID.";
    }
    notify(msg, ok ? "success" : "error");
  }

  return { ok, okCount, total: results.length, missingNames, results };
}

function updateGlobalDupRecipientSummary() {
  const el = document.getElementById("globalDupRecipientSummary");
  if (!el) return;
  const n = Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
    ? displaySettings.telegramGlobalDuplicateRecipientIds.length
    : 0;
  el.textContent = n ? `Выбрано получателей копий: ${n}` : "Получателей копий не выбрано";
}

function getGlobalDupEmployeeEntries() {
  const rows = getSectionById("employees")?.rows || [];
  return rows
    .map((r) => ({
      id: String(r[EMPLOYEE_COLUMNS.id] ?? "").trim(),
      name: normalizePersonName(r[EMPLOYEE_COLUMNS.fullName]),
      position: normalizePersonName(r[EMPLOYEE_COLUMNS.position])
    }))
    .filter((e) => e.id && e.name)
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
}

function employeeMatchesDupLeftFilters(entry, searchRaw, positionFilter) {
  const pf = String(positionFilter || "").trim();
  if (pf && entry.position !== pf) return false;
  const q = normalizePersonName(searchRaw).toLowerCase();
  if (!q) return true;
  const hay = `${entry.name} ${entry.position}`.toLowerCase();
  return hay.includes(q);
}

function dupShuttleRowHtml(e, side) {
  if (side === "left") {
    return `
      <label class="dup-shuttle-row dup-shuttle-row--left">
        <input type="checkbox" class="dup-cb-left" data-emp-id="${escapeHtmlAttr(e.id)}" />
        <span class="dup-shuttle-name">${escapeHtmlText(e.name)}</span>
        <span class="dup-shuttle-pos">${escapeHtmlText(e.position || "—")}</span>
      </label>`;
  }
  const allow = new Set(
    Array.isArray(displaySettings.telegramCloseConfirmAllowedIds)
      ? displaySettings.telegramCloseConfirmAllowedIds.map((x) => String(x).trim()).filter(Boolean)
      : []
  );
  const confirmChecked = allow.has(e.id) ? "checked" : "";
  return `
      <div class="dup-shuttle-row dup-shuttle-row--right">
        <label class="dup-shuttle-cell-cb" title="Выбор для переноса">
          <input type="checkbox" class="dup-cb-right" data-emp-id="${escapeHtmlAttr(e.id)}" />
        </label>
        <span class="dup-shuttle-name">${escapeHtmlText(e.name)}</span>
        <span class="dup-shuttle-pos">${escapeHtmlText(e.position || "—")}</span>
        <label class="dup-shuttle-cell-confirm" title="Может подтверждать закрытие задачи в Telegram">
          <input type="checkbox" class="dup-cb-confirm-close" data-emp-id="${escapeHtmlAttr(e.id)}" ${confirmChecked} />
        </label>
      </div>`;
}

function renderGlobalDupShuttleIntoDom() {
  const leftEl = document.getElementById("dupRecipientLeftScroll");
  const rightEl = document.getElementById("dupRecipientRightScroll");
  if (!leftEl || !rightEl) return;

  const searchEl = document.getElementById("dupRecipientSearchInput");
  const posEl = document.getElementById("dupRecipientPositionFilter");
  const searchRaw = searchEl ? String(searchEl.value || "") : "";
  const positionFilter = posEl ? String(posEl.value || "").trim() : "";

  const selected = new Set(
    (Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
      ? displaySettings.telegramGlobalDuplicateRecipientIds
      : []
    )
      .map((x) => String(x).trim())
      .filter(Boolean)
  );

  const all = getGlobalDupEmployeeEntries();
  let leftHtml = "";
  if (!all.length) {
    leftHtml = "";
  } else {
    leftHtml = all
      .filter((e) => !selected.has(e.id))
      .filter((e) => employeeMatchesDupLeftFilters(e, searchRaw, positionFilter))
      .map((e) => dupShuttleRowHtml(e, "left"))
      .join("");
  }

  const rightHtml = all
    .filter((e) => selected.has(e.id))
    .map((e) => dupShuttleRowHtml(e, "right"))
    .join("");

  leftEl.innerHTML = !all.length
    ? '<p class="dup-shuttle-empty">Нет сотрудников в справочнике</p>'
    : leftHtml || '<p class="dup-shuttle-empty">Нет записей по фильтру</p>';
  rightEl.innerHTML = rightHtml || '<p class="dup-shuttle-empty">Никого не выбрано</p>';
  updateGlobalDupRecipientSummary();
  updateDupSelectAllHeaderState();
}

function updateDupSelectAllHeaderState() {
  const leftMaster = document.getElementById("dupRecipientSelectAllLeft");
  const rightMaster = document.getElementById("dupRecipientSelectAllRight");
  const leftCbs = Array.from(document.querySelectorAll("#dupRecipientLeftScroll .dup-cb-left"));
  const rightCbs = Array.from(document.querySelectorAll("#dupRecipientRightScroll .dup-cb-right"));

  const sync = (master, list) => {
    if (!master) return;
    if (!list.length) {
      master.checked = false;
      master.indeterminate = false;
      master.disabled = true;
      return;
    }
    master.disabled = false;
    const checkedCount = list.filter((cb) => cb.checked).length;
    master.checked = checkedCount === list.length;
    master.indeterminate = checkedCount > 0 && checkedCount < list.length;
  };

  sync(leftMaster, leftCbs);
  sync(rightMaster, rightCbs);
}

function buildDupPositionFilterOptionsHtml() {
  const positions = getUniqueValues(getSectionById("employees")?.rows || [], EMPLOYEE_COLUMNS.position)
    .map((p) => normalizePersonName(p))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "ru"));
  return `<option value="">Все должности</option>${positions
    .map((p) => `<option value="${escapeHtmlAttr(p)}">${escapeHtmlText(p)}</option>`)
    .join("")}`;
}

function attachGlobalDuplicateRecipientsHandlers() {
  const addBtn = document.getElementById("dupRecipientAddBtn");
  const removeBtn = document.getElementById("dupRecipientRemoveBtn");
  const searchInput = document.getElementById("dupRecipientSearchInput");
  const posFilter = document.getElementById("dupRecipientPositionFilter");

  const commitIds = (nextSet) => {
    displaySettings.telegramGlobalDuplicateRecipientIds = Array.from(nextSet);
    const dup = new Set(displaySettings.telegramGlobalDuplicateRecipientIds.map((x) => String(x).trim()));
    if (!Array.isArray(displaySettings.telegramCloseConfirmAllowedIds)) {
      displaySettings.telegramCloseConfirmAllowedIds = [];
    }
    displaySettings.telegramCloseConfirmAllowedIds = displaySettings.telegramCloseConfirmAllowedIds
      .map((x) => String(x).trim())
      .filter((id) => dup.has(id));
    saveDisplaySettings();
    renderGlobalDupShuttleIntoDom();
  };

  addBtn?.addEventListener("click", () => {
    const selected = new Set(
      (Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
        ? displaySettings.telegramGlobalDuplicateRecipientIds
        : []
      )
        .map((x) => String(x).trim())
        .filter(Boolean)
    );
    document.querySelectorAll("#dupRecipientLeftScroll .dup-cb-left:checked").forEach((cb) => {
      const id = cb.getAttribute("data-emp-id");
      if (id) selected.add(id);
    });
    commitIds(selected);
  });

  removeBtn?.addEventListener("click", () => {
    const selected = new Set(
      (Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
        ? displaySettings.telegramGlobalDuplicateRecipientIds
        : []
      )
        .map((x) => String(x).trim())
        .filter(Boolean)
    );
    document.querySelectorAll("#dupRecipientRightScroll .dup-cb-right:checked").forEach((cb) => {
      const id = cb.getAttribute("data-emp-id");
      if (id) selected.delete(id);
    });
    commitIds(selected);
  });

  const scheduleLeftRefresh = () => {
    renderGlobalDupShuttleIntoDom();
  };

  searchInput?.addEventListener("input", scheduleLeftRefresh);
  posFilter?.addEventListener("change", scheduleLeftRefresh);

  const leftMaster = document.getElementById("dupRecipientSelectAllLeft");
  const rightMaster = document.getElementById("dupRecipientSelectAllRight");
  leftMaster?.addEventListener("change", () => {
    const on = Boolean(leftMaster.checked);
    document.querySelectorAll("#dupRecipientLeftScroll .dup-cb-left").forEach((cb) => {
      cb.checked = on;
    });
    leftMaster.indeterminate = false;
    updateDupSelectAllHeaderState();
  });
  rightMaster?.addEventListener("change", () => {
    const on = Boolean(rightMaster.checked);
    document.querySelectorAll("#dupRecipientRightScroll .dup-cb-right").forEach((cb) => {
      cb.checked = on;
    });
    rightMaster.indeterminate = false;
    updateDupSelectAllHeaderState();
  });

  document.getElementById("dupRecipientLeftScroll")?.addEventListener("change", (e) => {
    const t = e.target;
    if (t && t.classList && t.classList.contains("dup-cb-left")) {
      updateDupSelectAllHeaderState();
    }
  });
  document.getElementById("dupRecipientRightScroll")?.addEventListener("change", (e) => {
    const t = e.target;
    if (!t || !t.classList) return;
    if (t.classList.contains("dup-cb-right")) {
      updateDupSelectAllHeaderState();
      return;
    }
    if (t.classList.contains("dup-cb-confirm-close")) {
      const id = String(t.getAttribute("data-emp-id") || "").trim();
      if (!id) return;
      const set = new Set(
        Array.isArray(displaySettings.telegramCloseConfirmAllowedIds)
          ? displaySettings.telegramCloseConfirmAllowedIds.map((x) => String(x).trim()).filter(Boolean)
          : []
      );
      const dup = new Set(
        (Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)
          ? displaySettings.telegramGlobalDuplicateRecipientIds
          : []
        ).map((x) => String(x).trim()).filter(Boolean)
      );
      if (!dup.has(id)) return;
      if (t.checked) set.add(id);
      else set.delete(id);
      displaySettings.telegramCloseConfirmAllowedIds = Array.from(set);
      saveDisplaySettings();
    }
  });

  renderGlobalDupShuttleIntoDom();
}

function normalizeDateDisplayFormatId(value) {
  const id = String(value || "");
  return ["DMY_DOT", "ISO", "DMY_SLASH", "MDY_SLASH"].includes(id) ? id : "DMY_DOT";
}

function normalizeTimeDisplayFormatId(value) {
  return String(value) === "12" ? "12" : "24";
}

function normalizeGoogleSheetsInterval(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 30;
  return Math.min(1440, Math.max(1, Math.floor(n)));
}

function formatGoogleSheetsSyncStatusText() {
  const status = String(displaySettings.googleSheetsLastSyncStatus || "").trim();
  const atIso = String(displaySettings.googleSheetsLastSyncAt || "").trim();
  const msg = String(displaySettings.googleSheetsLastSyncMessage || "").trim();
  const mode = String(displaySettings.googleSheetsLastSyncMode || "").trim();
  const rows = Number(displaySettings.googleSheetsLastSyncRows) || 0;
  const atLabel = atIso ? formatTrashDate(new Date(atIso).getTime()) : "—";
  const modeLabel = mode === "auto" ? "Авто" : mode === "manual" ? "Ручной" : "—";
  if (!status) return `Статус: —\nПоследний запуск: ${atLabel}`;
  if (status === "ok") {
    return `Статус: Успешно\nПоследний запуск: ${atLabel}\nРежим: ${modeLabel}\nСтрок: ${rows}\n${msg || ""}`.trim();
  }
  if (status === "error") {
    return `Статус: Ошибка\nПоследний запуск: ${atLabel}\nРежим: ${modeLabel}\n${msg || "—"}`.trim();
  }
  return `Статус: ${status}\nПоследний запуск: ${atLabel}\n${msg || ""}`.trim();
}

function getGoogleSheetsBrandIconHtml() {
  return `
    <span class="status-dialog-gs-icon" aria-hidden="true">
      <svg viewBox="0 0 48 48" fill="none">
        <path d="M8 5h20l12 12v24a2 2 0 0 1-2 2H10a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2z" fill="#0F9D58"/>
        <path d="M28 5v10a2 2 0 0 0 2 2h10L28 5z" fill="#B7E1CD"/>
        <rect x="14" y="22" width="20" height="3" rx="1.5" fill="#fff"/>
        <rect x="14" y="28" width="20" height="3" rx="1.5" fill="#fff"/>
        <rect x="14" y="34" width="12" height="3" rx="1.5" fill="#fff"/>
      </svg>
    </span>
  `;
}

function showStatusDialog({ title = "Уведомление", message = "", type = "info", icon = "" } = {}) {
  const overlay = document.createElement("div");
  overlay.className = "status-dialog-overlay";
  const safeTitle = escapeHtmlText(String(title || "Уведомление"));
  const safeMessage = escapeHtmlText(String(message || "")).replace(/\r\n?|\n/g, "<br />");
  const typeClass = type === "success" ? "is-success" : type === "error" ? "is-error" : "is-info";
  const iconHtml = icon === "googleSheets" ? getGoogleSheetsBrandIconHtml() : "";
  overlay.innerHTML = `
    <div class="status-dialog-box ${typeClass}">
      <h4>${iconHtml}<span>${safeTitle}</span></h4>
      <div class="status-dialog-message">${safeMessage}</div>
      <div class="status-dialog-actions">
        <button type="button" class="confirm-btn confirm-close-btn">OK</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  const close = () => overlay.remove();
  overlay.querySelector(".confirm-close-btn")?.addEventListener("click", close);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
}

async function triggerGoogleSheetsManualSync() {
  if (!isHostedRuntime() || !getAuthToken()) {
    showStatusDialog({
      title: "Google Sheets",
      message: "Ручная синхронизация доступна после входа в серверный режим.",
      type: "error",
      icon: "googleSheets"
    });
    return false;
  }
  try {
    const r = await fetch("/api/google-sheets/sync", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${getAuthToken()}`
      },
      body: JSON.stringify({})
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || j?.ok !== true) {
      showStatusDialog({
        title: "Google Sheets",
        message: String(j?.error || "Ошибка синхронизации"),
        type: "error",
        icon: "googleSheets"
      });
      return false;
    }
    const durMs = Number(j?.durationMs) || 0;
    const durSec = durMs > 0 ? Math.max(0.1, durMs / 1000).toFixed(1) : "";
    const details = [];
    if (j?.message) details.push(String(j.message));
    if (durSec) details.push(`Время синхронизации: ${durSec} сек.`);
    showStatusDialog({
      title: "Google Sheets",
      message: details.join("\n") || "Синхронизация выполнена",
      type: "success",
      icon: "googleSheets"
    });
    await pullRemoteAppState({ rerender: true });
    return true;
  } catch (e) {
    showStatusDialog({
      title: "Google Sheets",
      message: String(e?.message || "Ошибка сети"),
      type: "error",
      icon: "googleSheets"
    });
    return false;
  }
}

function getServerTimezone() {
  const raw = displaySettings.serverTimezone;
  if (raw !== undefined && raw !== null && String(raw).trim() !== "") {
    return String(raw).trim();
  }
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
  } catch (_) {
    return "UTC";
  }
}

function getCalendarDatePartsInTimeZone(date, timeZone) {
  const d = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(d.getTime())) return null;
  const tz = timeZone || getServerTimezone();
  try {
    const fmt = new Intl.DateTimeFormat("en-CA", {
      timeZone: tz,
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    });
    const parts = fmt.formatToParts(d);
    const y = Number(parts.find((p) => p.type === "year")?.value);
    const m = Number(parts.find((p) => p.type === "month")?.value);
    const day = Number(parts.find((p) => p.type === "day")?.value);
    if (!y || !m || !day) return null;
    return { year: y, month: m, day };
  } catch (_) {
    return null;
  }
}

function parseRuDateStringToParts(value) {
  const s = String(value || "").trim();
  if (!s) return null;
  const [a, b, c] = s.split(".");
  if (!a || !b || !c) return null;
  const day = Number(a);
  const month = Number(b);
  const year = Number(c);
  if (!day || !month || !year) return null;
  if (day < 1 || day > 31 || month < 1 || month > 12) return null;
  return { day, month, year };
}

function formatDatePartsStorage(day, month, year) {
  return `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`;
}

function calendarDiffDays(fromParts, toParts) {
  const a = Date.UTC(fromParts.year, fromParts.month - 1, fromParts.day);
  const b = Date.UTC(toParts.year, toParts.month - 1, toParts.day);
  return Math.round((b - a) / 86400000);
}

function formatStoredDateForDisplay(stored) {
  const p = parseRuDateStringToParts(stored);
  if (!p) return String(stored || "");
  const mode = normalizeDateDisplayFormatId(displaySettings.dateDisplayFormat);
  const dd = String(p.day).padStart(2, "0");
  const mm = String(p.month).padStart(2, "0");
  const yyyy = String(p.year);
  if (mode === "ISO") return `${yyyy}-${mm}-${dd}`;
  if (mode === "DMY_SLASH") return `${dd}/${mm}/${yyyy}`;
  if (mode === "MDY_SLASH") return `${mm}/${dd}/${yyyy}`;
  return formatDatePartsStorage(p.day, p.month, p.year);
}

function parseRuDate(value) {
  const parts = parseRuDateStringToParts(value);
  if (!parts) return null;
  const date = new Date(parts.year, parts.month - 1, parts.day);
  return Number.isNaN(date.getTime()) ? null : date;
}
const PRIORITY_OPTIONS = ["Средний", "Высокий", "Критический"];
const TASKS_UNSENT_TAB_ID = "unsent_import";
const STATUS_TABS = [
  { id: "all", label: "Все статусы" },
  { id: "Новый", label: "Новый" },
  { id: "В процессе", label: "В процессе" },
  { id: "Закрыт", label: "Закрыт" },
  { id: TASKS_UNSENT_TAB_ID, label: "Не отправленные" },
  { id: "trash", label: "Корзина" }
];
const TASK_IMPORT_COLUMNS = [
  { key: "object", label: "Название объекта", aliases: ["Название объекта", "Объект"] },
  { key: "priority", label: "Приоритет", aliases: ["Приоритет"] },
  { key: "addedDate", label: "Дата постановки задачи", aliases: ["Дата постановки задачи", "Дата добавления"] },
  { key: "phase", label: "Фаза", aliases: ["Фаза"] },
  { key: "phaseSection", label: "Раздел", aliases: ["Раздел"] },
  { key: "phaseSubsection", label: "Подраздел", aliases: ["Подраздел"] },
  { key: "task", label: "Задача", aliases: ["Задача", "Название задачи"] },
  { key: "responsible", label: "Постановщик задачи", aliases: ["Постановщик задачи", "Контролирующий ответственный"] },
  { key: "assignedResponsible", label: "Ответственный", aliases: ["Ответственный", "Исполнитель"] },
  { key: "note", label: "Коментарии к задаче", aliases: ["Коментарии к задаче", "Комментарии к задаче", "Примичание", "Примечание"] },
  { key: "dueDate", label: "Плановый срок устранения", aliases: ["Плановый срок устранения", "Срок устранения", "Срок"] }
];
const SECTION_GROUPS = {
  reference: {
    id: "reference",
    title: "Справочник",
    icon: "database",
    sections: ["data", "phases", "phaseSections", "phaseSubsections"]
  },
  users: {
    id: "users",
    title: "Пользователи",
    icon: "users",
    sections: ["employees", "roles", "departments"]
  }
};

/** Сид «Объекты»; `mbc_objects_seed_version` в localStorage — при увеличении версии строки обновляются у всех, у кого сохранены старые данные */
const OBJECTS_SEED_VERSION_KEY = "mbc_objects_seed_version";
const OBJECTS_SEED_VERSION = 1;
const DEFAULT_OBJECTS_ROWS = [
  ["1", "Center one by MBC", "г. Ташкент, ул. Садика Азимова", "Активен", "Nasirov Karimjon", "Qosimov Sarvar", ""],
  ["2", "Test project", "1", "Активен", "Yeloyev Zaur", "Xamidov Doniyor", ""],
  ["3", "ЖК REGNUM PLAZA", "г. Ташкент, Мирузо Улугбекский район", "Активен", "Umarov Shaxzod", "Temirov Vaxob", ""],
  ["4", "ЖК SOY BO'YI", "г. Ташкент, Учтепинский район, ул. Юсуфа Саккокий", "Активен", "Shomaxsudov Shosaid", "Anvarov Sardor", ""],
  ["5", "ЖК Saadiyat", "г. Ташкент, Юнусабадский район", "Активен", "Yeloyev Zaur", "Xamidov Doniyor", ""]
];

let sections = [
  {
    id: "tasks",
    title: "Задачи",
    columns: [
      "ID",
      "Название объекта",
      "Статус",
      "Приоритет",
      "Дата постановки задачи",
      "Фаза",
      "Раздел",
      "Подраздел",
      "Задача",
      "Постановщик задачи",
      "Ответственный",
      "Коментарии к задаче",
      "Комментарии сотрудника (Результат)",
      "Коментарии администратора",
      "Плановый срок устранения",
      "Факт даты устранения",
      "Медиа до (5)",
      "Медиа после (5)",
      "Ознакомление",
      "Дата последней отправки"
    ],
    rows: [
      [
        "1",
        "Склад Север",
        "В процессе",
        "Высокий",
        "15.04.2026",
        "Инициация",
        "Первичный анализ ЗУ",
        "УЧРП",
        "Проверить состояние стеллажей",
        "Иван Петров",
        "Иван Петров",
        "Требуется фотофиксация",
        "Осмотр зоны А и Б, составить акт",
        "Осмотр зоны А выполнен",
        "20.04.2026",
        "",
        "",
        "",
        "Не прочитано\n—",
        "—"
      ],
      [
        "2",
        "Офис Центр",
        "Новый",
        "Средний",
        "16.04.2026",
        "Реализация",
        "СМР",
        "Надземные конструкции",
        "Обновить схему эвакуации",
        "Мария Волкова",
        "Мария Волкова",
        "Согласовать с отделом безопасности",
        "Подготовить макет, утвердить, разместить",
        "Макет в подготовке",
        "24.04.2026",
        "",
        "",
        "",
        "Не прочитано\n—",
        "—"
      ],
      [
        "3",
        "Архив Юг",
        "Закрыт",
        "Средний",
        "12.04.2026",
        "Завершение",
        "Тех.приемка УК",
        "ПНР",
        "Заменить замок архивной комнаты",
        "Антон Кузнецов",
        "Антон Кузнецов",
        "Работы выполнены подрядчиком",
        "Согласование и замена замка",
        "Исполнено полностью",
        "14.04.2026",
        "14.04.2026",
        "",
        "",
        "Не прочитано\n—"
      ]
    ]
  },
  {
    id: "employees",
    title: "Сотрудники",
    columns: ["ID", "ФИО", "Отдел", "Должность", "Телефон", "Telegram", "Chat ID", "Активность"],
    rows: [
      ["1", "Ольга Смирнова", "Проектная группа", "Frontend", "+998 90 111 11 11", "Подключен", "901111111", "Активен"],
      ["2", "Сергей Орлов", "ПТО", "Backend", "+998 90 222 22 22", "Не подключен", "", "Активен"],
      ["3", "Елена Белова", "Плановый отдел", "QA", "+998 90 333 33 33", "Подключен", "903333333", "В отпуске"]
    ]
  },
  {
    id: "departments",
    title: "Отдел",
    columns: ["ID", "Отдел", "Руководитель отдела", "Тип"],
    rows: SYSTEM_DEPARTMENTS.map((department, index) => [String(index + 1), department, "", "Системный"])
  },
  {
    id: "phases",
    title: "Фаза",
    columns: ["ID", "Фаза"],
    rows: [
      ["1", "Инициация"],
      ["2", "Выбор"],
      ["3", "Проработка"],
      ["4", "Реализация"],
      ["5", "Завершение"]
    ]
  },
  {
    id: "phaseSections",
    title: "Раздел",
    columns: ["ID", "Раздел"],
    rows: [
      ["1", "Первичный анализ ЗУ"],
      ["2", "Глубокий анализ ЗУ"],
      ["3", "Детальная проверка"],
      ["4", "Финансирование"],
      ["5", "Другое"],
      ["6", "Подбор и найм персонала"],
      ["7", "Эскизное проектирование"],
      ["8", "Административные согласования"],
      ["9", "Геология"],
      ["10", "ППР"],
      ["11", "Рабочее проектирование"],
      ["12", "Маркетинг"],
      ["13", "Офис продаж"],
      ["14", "Продажи"],
      ["15", "Поставки"],
      ["16", "Тендерные процедуры"],
      ["17", "СМР"],
      ["18", "Рабочая комиссия"],
      ["19", "Устранение замечаний"],
      ["20", "ЗоС"],
      ["21", "Акт ГК"],
      ["22", "Кадастр-Кучирма"],
      ["23", "Финансовый результат проекта"],
      ["24", "Тех.приемка УК"]
    ]
  },
  {
    id: "phaseSubsections",
    title: "Подраздел",
    columns: ["ID", "Подраздел"],
    rows: [
      ["1", "Альбом к Градсовету"],
      ["2", "Дизайн проект благоустройства"],
      ["3", "Дизайн проект + РД МОП"],
      ["4", "Другое"],
      ["5", "Градсовет"],
      ["6", "СТУ"],
      ["7", "УЧРП"],
      ["8", "Доработка РП"],
      ["9", "Экспертиза"],
      ["10", "РнС"],
      ["11", "Дизайн проект"],
      ["12", "Рабочий проект"],
      ["13", "Офис продаж"],
      ["14", "Мобилизация"],
      ["15", "Земляные работы"],
      ["16", "Шпунтовое ограждение"],
      ["17", "Фундаментные работы"],
      ["18", "Подземные конструкции"],
      ["19", "Надземные конструкции"],
      ["20", "Кровельные работы"],
      ["21", "Фасадные конструкции"],
      ["22", "СПК"],
      ["23", "Кладочные работы"],
      ["24", "Черновая отделка"],
      ["25", "Чистовая отделка"],
      ["26", "Инженерные сети"],
      ["27", "Лифтовое оборудование"],
      ["28", "ПНР"],
      ["29", "Внутриплощадочные сети"],
      ["30", "Благоустройство"],
      ["31", "Внеплощадочные сети"]
    ]
  },
  {
    id: "roles",
    title: "Должности",
    columns: ["ID", "Должность", "Тип"],
    rows: SYSTEM_ROLES.map((role, index) => [String(index + 1), role, "Системный"])
  },
  {
    id: "data",
    title: "Ответственные",
    columns: ["ID", "Фаза", "Раздел", "Подраздел", "Ответственные", "Описание"],
    rows: [
      ["1", "Инициация", "Первичный анализ ЗУ", "УЧРП", "Иван Петров, Мария Волкова", "Сбор исходных данных по ЗУ"],
      ["2", "Реализация", "СМР", "Фундаментные работы", "Мария Волкова, Сергей Орлов", "Контроль строительства фундамента"],
      ["3", "Завершение", "Тех.приемка УК", "ПНР", "Антон Кузнецов, Елена Белова", "Приёмка систем и сдача объекта"]
    ]
  },
  {
    id: "objects",
    title: "Объекты",
    columns: ["ID", "Наименование", "Адрес", "Статус", "РП", "ЗРП", "Фото объекта"],
    rows: JSON.parse(JSON.stringify(DEFAULT_OBJECTS_ROWS))
  }
];

/** Копия строк справочников фаз на случай пустых `rows` из БД/localStorage (иначе в задачах пустые списки). */
const PHASE_CATALOG_DEFAULT_ROWS = Object.fromEntries(
  sections
    .filter((s) => ["phases", "phaseSections", "phaseSubsections"].includes(s.id))
    .map((s) => [s.id, JSON.parse(JSON.stringify(s.rows))])
);

let activeSectionId = "tasks";
let isSettingsOpen = false;
const filtersBySection = {};
const selectedRowsBySection = {};
const filterPanelOpenBySection = {};
const statusTabBySection = {};
let isSidebarCollapsed = false;
const activeRowBySection = {};
const visibleColumnsBySection = {};
const columnOrderBySection = {};
const headerNumberingBySection = {};
const mediaPreviewStore = {};
/** Превью фото объекта: ключ obj-ph-{id строки объекта} */
const objectPhotoPreviewStore = {};
const trashBySection = {};
/** Экземпляры Chart.js на экране «Отчёт» */
let reportChartInstances = [];
/** null = выбраны все статусы */
let reportStatusFilter = null;
let reportDateFrom = "";
let reportDateTo = "";
/** Точное имя объекта из списка или ввода; пусто = все */
let reportFilterObject = "";
/** Исполнитель (assignedResponsible); пусто = все */
let reportFilterEmployee = "";
const REPORT_WEEK_ROWS_STEP = 5;
let reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
/** Подпись для задач без объекта (режим «по объектам») */
const TASKS_EMPTY_OBJECT_LABEL = "— без объекта —";
const TASKS_DEFAULT_LIST_PAGE_SIZE = 50;
/** Страница списка задач (режим пагинации) */
let tasksListPage = 1;
/** Сколько строк показано в режиме «порциями» */
let tasksListVisibleCount = TASKS_DEFAULT_LIST_PAGE_SIZE;
/** null — показан список объектов; иначе ключ объекта (как в ячейке) */
let tasksBrowseObjectKey = null;
/** Панель фильтров отчёта скрыта до открытия кнопкой */
let reportFiltersPanelOpen = false;
/** Во время drag-and-drop плиток отчёта */
let reportChartDragId = null;
/** Просмотр отчёта по временной ссылке (?share=…) */
let sharedReportMode = false;
let reportShareRowsOverride = null;
let sharedReportExpiresAt = 0;
let otherSettingsActiveTab = "general";
let currentAuthRole = "user";

let displaySettings = {
  highlightClosed: false,
  highlightNeedDecision: false,
  telegramBotToken: "",
  /** @type {string} username бота без @ — подставляется после setWebhook / getMe */
  telegramBotUsername: "",
  /** Имя (и фамилия) бота в Telegram из getMe — только для отображения */
  telegramBotDisplayName: "",
  /** Числовой chat_id админа: проверочное сообщение «Проверить бота» уходит только сюда */
  telegramAdminChatId: "",
  /** Пустая строка = часовой пояс браузера */
  serverTimezone: "",
  dateDisplayFormat: "DMY_DOT",
  timeDisplayFormat: "24",
  timeShowSeconds: false,
  overdueNotificationsEnabled: false,
  overdueNotificationsTime: "09:00",
  reminderSettings: Object.fromEntries(
    STATUS_OPTIONS.map((status) => [status, { days: "none", text: "" }])
  ),
  taskMessageTemplatesByStatus: Object.fromEntries(STATUS_OPTIONS.map((status) => [status, ""])),
  /** ID сотрудников (колонка ID): всегда получают копию по всем задачам при отправке в Telegram */
  telegramGlobalDuplicateRecipientIds: [],
  /** Подмножество ID из списка копий: могут подтверждать закрытие задачи в Telegram (кнопки уведомления) */
  telegramCloseConfirmAllowedIds: [],
  /** Сообщение исполнителю после подтверждения закрытия администратором (токены [ид_задачи], [название_задачи], [статус], [объект]) */
  telegramCloseAcceptedTemplate:
    "Задача [ид_задачи] ([название_задачи]): закрытие подтверждено.",
  googleSheetsEnabled: false,
  googleSheetsAutoSyncEnabled: false,
  googleSheetsSpreadsheetId: "",
  googleSheetsSummarySheetName: "Сводная",
  googleSheetsIncludeObjectSheets: true,
  googleSheetsSyncIntervalMinutes: 30,
  googleSheetsLastSyncStatus: "",
  googleSheetsLastSyncAt: "",
  googleSheetsLastSyncAtMs: 0,
  googleSheetsLastSyncMessage: "",
  googleSheetsLastSyncRows: 0,
  googleSheetsLastSyncMode: "",
  /** ID задач, импортированных без отправки в Telegram (или с ошибкой отправки). */
  pendingImportedTaskIds: [],
  /** pagination | chunks — страницы или «Показать ещё» */
  tasksListPagingMode: "pagination",
  /** Размер страницы / одной порции (5–500) */
  tasksListPageSize: 50,
  /** flat | byObject — таблица сразу или сначала выбор объекта */
  tasksListBrowseMode: "flat"
};
let taskMultiState = {};
const expandedTaskAssigneeRows = new Set();

function iconSvg(name) {
  const attrs = 'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide-icon" aria-hidden="true"';
  const icons = {
    "layout-dashboard": `<svg ${attrs}><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>`,
    user: `<svg ${attrs}><circle cx="12" cy="8" r="4"/><path d="M6 20c1.6-2.7 4-4 6-4s4.4 1.3 6 4"/></svg>`,
    panelLeft: `<svg ${attrs}><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18"/></svg>`,
    listChecks: `<svg ${attrs}><path d="M9 6h11"/><path d="M9 12h11"/><path d="M9 18h11"/><path d="m3 6 1 1 2-2"/><path d="m3 12 1 1 2-2"/><path d="m3 18 1 1 2-2"/></svg>`,
    users: `<svg ${attrs}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
    userCheck: `<svg ${attrs}><path d="m16 11 2 2 4-4"/><path d="M8 7a4 4 0 1 1 8 0 4 4 0 0 1-8 0"/><path d="M6 21v-2a6 6 0 0 1 9-5"/></svg>`,
    database: `<svg ${attrs}><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v6c0 1.7 4 3 9 3s9-1.3 9-3V5"/><path d="M3 11v6c0 1.7 4 3 9 3s9-1.3 9-3v-6"/></svg>`,
    building2: `<svg ${attrs}><path d="M6 22V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v18"/><path d="M2 22h20"/><path d="M10 6h4"/><path d="M10 10h4"/><path d="M10 14h4"/><path d="M10 18h4"/></svg>`,
    settings: `<svg ${attrs}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06A1.6 1.6 0 0 0 15 19.4a1.6 1.6 0 0 0-1 .6 1.6 1.6 0 0 0-.4 1V21a2 2 0 1 1-4 0v-.1a1.6 1.6 0 0 0-.4-1 1.6 1.6 0 0 0-1-.6 1.6 1.6 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.6 1.6 0 0 0 4.6 15a1.6 1.6 0 0 0-.6-1 1.6 1.6 0 0 0-1-.4H3a2 2 0 1 1 0-4h.1a1.6 1.6 0 0 0 1-.4 1.6 1.6 0 0 0 .6-1 1.6 1.6 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.6 1.6 0 0 0 9 4.6a1.6 1.6 0 0 0 1-.6 1.6 1.6 0 0 0 .4-1V3a2 2 0 1 1 4 0v.1a1.6 1.6 0 0 0 .4 1 1.6 1.6 0 0 0 1 .6 1.6 1.6 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.6 1.6 0 0 0 19.4 9c.28.3.47.67.54 1.08.07.41.02.83-.14 1.22.16.39.21.81.14 1.22A1.6 1.6 0 0 0 19.4 15Z"/></svg>`,
    briefcase: `<svg ${attrs}><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/><path d="M2 13h20"/></svg>`,
    barChart: `<svg ${attrs}><path d="M3 3v18h18"/><path d="M7 16V8"/><path d="M12 16v-5"/><path d="M17 16V5"/></svg>`
  };
  return icons[name] || "";
}

function withIcon(iconName, text) {
  return `<span class="icon-label">${iconSvg(iconName)}<span>${text}</span></span>`;
}

function withLucideIcon(iconName, text) {
  return `<span class="icon-label"><i data-lucide="${iconName}" class="lucide-icon" aria-hidden="true"></i><span>${text}</span></span>`;
}

function initLucideIcons() {
  if (window.lucide && typeof window.lucide.createIcons === "function") {
    window.lucide.createIcons();
  }
}

function getSectionIcon(sectionId) {
  const iconBySection = {
    tasks: "listChecks",
    employees: "users",
    departments: "users",
    phases: "database",
    phaseSections: "database",
    phaseSubsections: "database",
    roles: "briefcase",
    data: "database",
    objects: "building2"
  };
  return iconBySection[sectionId] || "layout-dashboard";
}

function selectSection(sectionId) {
  const allowSettings = canAccessSettingsMenu();
  if (!allowSettings && sectionId !== "tasks" && sectionId !== "report") {
    sectionId = "tasks";
  }
  startTurboLoader();
  if (sectionId === "report" && activeSectionId !== "report") {
    reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
  }
  activeSectionId = sectionId;
  saveActiveSection(sectionId);
  if (sectionId !== "tasks" && sectionId !== "report") {
    isSettingsOpen = true;
  }
  renderSidebarMenu();
  renderTable();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      finishTurboLoader();
    });
  });
}

function renderSidebarMenu() {
  const allowSettings = canAccessSettingsMenu();
  if (!allowSettings && activeSectionId !== "tasks" && activeSectionId !== "report") {
    activeSectionId = "tasks";
    saveActiveSection("tasks");
  }
  const sectionById = new Map(sections.map((section) => [section.id, section]));
  const isReferenceActive = SECTION_GROUPS.reference.sections.includes(activeSectionId);
  const isUsersActive = SECTION_GROUPS.users.sections.includes(activeSectionId);
  tabsRoot.innerHTML = "";
  tabsRoot.classList.toggle("is-collapsed", isSidebarCollapsed);

  const tasksButton = document.createElement("button");
  tasksButton.className = `tab-btn top-level ${activeSectionId === "tasks" ? "active" : ""}`;
  tasksButton.type = "button";
  tasksButton.innerHTML = withIcon("listChecks", isSidebarCollapsed ? "" : "Задачи");
  tasksButton.title = "Задачи";
  tasksButton.addEventListener("click", () => selectSection("tasks"));
  tabsRoot.appendChild(tasksButton);

  const reportButton = document.createElement("button");
  reportButton.className = `tab-btn top-level ${activeSectionId === "report" ? "active" : ""}`;
  reportButton.type = "button";
  reportButton.innerHTML = withIcon("barChart", isSidebarCollapsed ? "" : "Отчёт");
  reportButton.title = "Отчёт";
  reportButton.addEventListener("click", () => selectSection("report"));
  tabsRoot.appendChild(reportButton);

  if (allowSettings) {
    const settingsButton = document.createElement("button");
    settingsButton.className = "tab-btn top-level settings-toggle";
    settingsButton.type = "button";
    settingsButton.innerHTML = isSidebarCollapsed
      ? `<span class="icon-label">${iconSvg("panelLeft")}<span></span></span>`
      : `<span class="icon-label">${iconSvg("panelLeft")}<span>Настройки</span></span><span class="menu-caret">${isSettingsOpen ? "▾" : "▸"}</span>`;
    settingsButton.title = "Настройки";
    settingsButton.addEventListener("click", () => {
      isSettingsOpen = !isSettingsOpen;
      renderSidebarMenu();
    });
    tabsRoot.appendChild(settingsButton);
  }

  if (allowSettings && isSettingsOpen && !isSidebarCollapsed) {
    const referenceButton = document.createElement("button");
    referenceButton.className = `tab-btn submenu-item ${isReferenceActive ? "active" : ""}`;
    referenceButton.type = "button";
    referenceButton.innerHTML = withIcon(SECTION_GROUPS.reference.icon, SECTION_GROUPS.reference.title);
    referenceButton.addEventListener("click", () => {
      const next = SECTION_GROUPS.reference.sections.includes(activeSectionId)
        ? activeSectionId
        : SECTION_GROUPS.reference.sections[0];
      selectSection(next);
    });
    tabsRoot.appendChild(referenceButton);

    const usersButton = document.createElement("button");
    usersButton.className = `tab-btn submenu-item ${isUsersActive ? "active" : ""}`;
    usersButton.type = "button";
    usersButton.innerHTML = withIcon(SECTION_GROUPS.users.icon, SECTION_GROUPS.users.title);
    usersButton.addEventListener("click", () => {
      const next = SECTION_GROUPS.users.sections.includes(activeSectionId)
        ? activeSectionId
        : SECTION_GROUPS.users.sections[0];
      selectSection(next);
    });
    tabsRoot.appendChild(usersButton);

    const objectsSection = sectionById.get("objects");
    if (objectsSection) {
      const objectsButton = document.createElement("button");
      objectsButton.className = `tab-btn submenu-item ${objectsSection.id === activeSectionId ? "active" : ""}`;
      objectsButton.type = "button";
      objectsButton.innerHTML = withIcon(getSectionIcon(objectsSection.id), objectsSection.title);
      objectsButton.addEventListener("click", () => selectSection(objectsSection.id));
      tabsRoot.appendChild(objectsButton);
    }

    const otherSettingsButton = document.createElement("button");
    otherSettingsButton.className = `tab-btn submenu-item ${activeSectionId === "otherSettings" ? "active" : ""}`;
    otherSettingsButton.type = "button";
    otherSettingsButton.innerHTML = withIcon("settings", "Прочие настройки");
    otherSettingsButton.addEventListener("click", () => selectSection("otherSettings"));
    tabsRoot.appendChild(otherSettingsButton);
  }
}

function normalizeTasksListDisplaySettings() {
  const mode = displaySettings.tasksListPagingMode === "chunks" ? "chunks" : "pagination";
  const browse = displaySettings.tasksListBrowseMode === "byObject" ? "byObject" : "flat";
  displaySettings.tasksListPagingMode = mode;
  displaySettings.tasksListBrowseMode = browse;
  let n = Number(displaySettings.tasksListPageSize);
  if (!Number.isFinite(n)) n = TASKS_DEFAULT_LIST_PAGE_SIZE;
  displaySettings.tasksListPageSize = Math.min(500, Math.max(5, Math.floor(n)));
}

function getTasksListPageSize() {
  normalizeTasksListDisplaySettings();
  return displaySettings.tasksListPageSize;
}

function resetTasksListPagingWindow() {
  tasksListPage = 1;
  tasksListVisibleCount = getTasksListPageSize();
}

function taskObjectLabelFromRow(row) {
  return String(row[TASK_COLUMNS.object] || "").trim() || TASKS_EMPTY_OBJECT_LABEL;
}

function groupTaskEntriesByObject(entries) {
  const map = new Map();
  for (const entry of entries) {
    const key = taskObjectLabelFromRow(entry.row);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(entry);
  }
  return Array.from(map.entries())
    .map(([key, list]) => ({ key, count: list.length, entries: list }))
    .sort((a, b) => b.count - a.count || a.key.localeCompare(b.key, "ru"));
}

function filterTaskEntriesByObjectKey(entries, objectKey) {
  if (objectKey == null) return entries;
  return entries.filter((e) => taskObjectLabelFromRow(e.row) === objectKey);
}

function getObjectPhotoPreviewKeyForRow(objectRow) {
  const oid = String(objectRow?.[OBJECT_COLUMNS.id] ?? "").trim();
  return oid ? `obj-ph-${oid}` : "";
}

function loadObjectPhotoThumbsFromStorage() {
  const sec = getSectionById("objects");
  const validIds = new Set();
  if (sec) {
    for (let i = 0; i < sec.rows.length; i++) {
      const oid = String(sec.rows[i][OBJECT_COLUMNS.id] ?? i).trim();
      if (oid) validIds.add(`obj-ph-${oid}`);
    }
  }
  for (const k of Object.keys(objectPhotoPreviewStore)) {
    if (!validIds.has(k)) {
      const prev = objectPhotoPreviewStore[k];
      if (prev?.url && String(prev.url).startsWith("blob:")) URL.revokeObjectURL(prev.url);
      delete objectPhotoPreviewStore[k];
    }
  }
  try {
    const raw = localStorage.getItem(OBJECT_PHOTO_THUMBS_STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    if (typeof data !== "object" || !data) return;
    for (const [k, v] of Object.entries(data)) {
      if (!validIds.has(k)) continue;
      if (v && typeof v.dataUrl === "string" && v.dataUrl.startsWith("data:")) {
        objectPhotoPreviewStore[k] = {
          name: v.name || "",
          type: v.type || "image/jpeg",
          url: v.dataUrl
        };
      }
    }
  } catch (_) {
    /* noop */
  }
}

function saveObjectPhotoThumbsToStorage() {
  try {
    const sec = getSectionById("objects");
    const out = {};
    if (sec) {
      for (let i = 0; i < sec.rows.length; i++) {
        const oid = String(sec.rows[i][OBJECT_COLUMNS.id] ?? i).trim();
        const pkey = oid ? `obj-ph-${oid}` : "";
        const v = pkey ? objectPhotoPreviewStore[pkey] : null;
        if (pkey && v?.url && String(v.url).startsWith("data:")) {
          out[pkey] = { dataUrl: v.url, name: v.name || "", type: v.type || "" };
        }
      }
    }
    localStorage.setItem(OBJECT_PHOTO_THUMBS_STORAGE_KEY, JSON.stringify(out));
  } catch (e) {
    console.warn("saveObjectPhotoThumbsToStorage", e);
  }
}

/** Склонение: 1 задача, 2 задачи, 5 задач */
function pluralTasksRu(count) {
  const n = Math.abs(Number(count)) || 0;
  const mod100 = n % 100;
  const mod10 = n % 10;
  if (mod100 > 10 && mod100 < 20) return `${n} задач`;
  if (mod10 === 1) return `${n} задача`;
  if (mod10 >= 2 && mod10 <= 4) return `${n} задачи`;
  return `${n} задач`;
}

function findObjectRowByTaskObjectLabel(objectLabel) {
  const sec = getSectionById("objects");
  if (!sec) return null;
  const want = normalizePersonName(objectLabel);
  if (!want || objectLabel === TASKS_EMPTY_OBJECT_LABEL) return null;
  for (let i = 0; i < sec.rows.length; i++) {
    const row = sec.rows[i];
    const n = normalizePersonName(row[OBJECT_COLUMNS.name] || "");
    if (n === want) return { row, rowIndex: i };
  }
  return null;
}

function renderObjectCardThumb(objectLabel) {
  const found = findObjectRowByTaskObjectLabel(objectLabel);
  if (!found) {
    return `<div class="tasks-object-pick-photo tasks-object-pick-photo--empty" aria-hidden="true"><span class="tasks-object-pick-logo" role="img" aria-label=""></span></div>`;
  }
  const key = getObjectPhotoPreviewKeyForRow(found.row);
  const prev = key ? objectPhotoPreviewStore[key] : null;
  const storedFileName = String(found.row?.[OBJECT_COLUMNS.photo] || "").trim();
  const fallbackUrl = resolveObjectPhotoSourceUrl(storedFileName);
  const src = prev?.url || fallbackUrl;
  if (src) {
    return `<div class="tasks-object-pick-photo"><img src="${escapeHtmlAttr(src)}" alt="" loading="lazy" /></div>`;
  }
  return `<div class="tasks-object-pick-photo tasks-object-pick-photo--empty" aria-hidden="true"><span class="tasks-object-pick-logo" role="img" aria-label=""></span></div>`;
}

function renderTasksObjectPickerHtml(grouped) {
  if (!grouped.length) {
    return `<p class="hint tasks-object-picker-empty">Нет задач по текущему фильтру.</p>`;
  }
  return `
    <div class="tasks-object-picker-grid">
      ${grouped
        .map(
          (g) => `
        <button type="button" class="tasks-object-pick-card" data-object-key="${escapeHtmlAttr(g.key)}">
          ${renderObjectCardThumb(g.key)}
          <div class="tasks-object-pick-card-body">
            <span class="tasks-object-pick-name">${escapeHtmlText(g.key)}</span>
            <span class="tasks-object-pick-count">${pluralTasksRu(g.count)}</span>
          </div>
        </button>`
        )
        .join("")}
    </div>`;
}

function renderTasksListFooterHtml(pagingMode, pageSize, currentPage, total, renderedCount) {
  if (total <= 0) return "";
  if (pagingMode === "chunks") {
    const shown = Math.min(renderedCount, total);
    const canMore = shown < total;
    return `
      <div class="tasks-list-footer tasks-list-footer--chunks">
        <span class="tasks-list-footer-meta">Показано ${shown} из ${total}<span class="tasks-list-footer-hint"> · внизу таблицы подгрузка при прокрутке</span></span>
        <button type="button" class="tasks-load-more-btn" id="tasksLoadMoreBtn" ${canMore ? "" : "disabled"}>Показать ещё</button>
      </div>`;
  }
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  const safePage = Math.min(Math.max(1, currentPage), totalPages);
  const start = total === 0 ? 0 : (safePage - 1) * pageSize + 1;
  const end = Math.min((safePage - 1) * pageSize + renderedCount, total);
  return `
    <div class="tasks-list-footer tasks-list-footer--pagination">
      <span class="tasks-list-footer-meta">Строки ${start}–${end} из ${total}</span>
      <div class="tasks-list-pager">
        <button type="button" class="tasks-page-prev-btn" id="tasksPagePrevBtn" ${safePage <= 1 ? "disabled" : ""}>Назад</button>
        <span class="tasks-list-page-label">Стр. ${safePage} / ${totalPages}</span>
        <button type="button" class="tasks-page-next-btn" id="tasksPageNextBtn" ${safePage >= totalPages ? "disabled" : ""}>Вперёд</button>
      </div>
    </div>`;
}

function attachTasksListFooterHandlers(section) {
  if (section.id !== "tasks") return;
  const prev = document.getElementById("tasksPagePrevBtn");
  const next = document.getElementById("tasksPageNextBtn");
  const more = document.getElementById("tasksLoadMoreBtn");
  const pageSize = getTasksListPageSize();
  if (prev) {
    prev.addEventListener("click", () => {
      if (tasksListPage <= 1) return;
      tasksListPage -= 1;
      renderTablePreserveScroll();
    });
  }
  if (next) {
    next.addEventListener("click", () => {
      tasksListPage += 1;
      renderTablePreserveScroll();
    });
  }
  if (more) {
    more.addEventListener("click", () => {
      tasksListVisibleCount += pageSize;
      renderTablePreserveScroll();
    });
  }
  attachTasksChunksInfiniteScroll(section);
}

function attachTasksChunksInfiniteScroll(section) {
  if (section.id !== "tasks" || displaySettings.tasksListPagingMode !== "chunks") return;
  const sentinel = document.getElementById("tasksChunksScrollSentinel");
  const wrap = sentinel?.closest(".table-wrap");
  if (!sentinel || !wrap) return;
  const obs = new IntersectionObserver(
    (entries) => {
      const e = entries[0];
      if (!e?.isIntersecting) return;
      const el = document.getElementById("tasksChunksScrollSentinel");
      if (!el) return;
      const all = Number(el.dataset.all || 0);
      const shown = Number(el.dataset.shown || 0);
      if (!all || shown >= all) return;
      tasksListVisibleCount += getTasksListPageSize();
      renderTablePreserveScroll();
    },
    { root: wrap, rootMargin: "120px", threshold: 0 }
  );
  obs.observe(sentinel);
  window._mbcTaskChunksObserver = obs;
}

function attachTasksObjectPickerHandlers(section) {
  document.querySelectorAll(".tasks-object-pick-card").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-object-key");
      if (key == null) return;
      tasksBrowseObjectKey = key;
      resetTasksListPagingWindow();
      renderTable();
    });
  });
  const back = document.getElementById("tasksBackToObjectsBtn");
  if (back) {
    back.addEventListener("click", () => {
      tasksBrowseObjectKey = null;
      resetTasksListPagingWindow();
      renderTable();
    });
  }
  const refreshBtn = document.getElementById("tasksObjectPickerRefreshBtn");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", () => {
      refreshCurrentViewData();
    });
  }
  Array.from(document.querySelectorAll('input[name="tasksQuickBrowseMode"]')).forEach((r) => {
    r.addEventListener("change", () => {
      displaySettings.tasksListBrowseMode = r.value === "byObject" ? "byObject" : "flat";
      if (displaySettings.tasksListBrowseMode === "flat") tasksBrowseObjectKey = null;
      saveDisplaySettings();
      resetTasksListPagingWindow();
      renderTable();
    });
  });
}

function getClockPartsInTimeZone(date, timeZone) {
  const d = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(d.getTime())) return null;
  const tz = timeZone || getServerTimezone();
  try {
    const fmt = new Intl.DateTimeFormat("en-GB", {
      timeZone: tz,
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    });
    const parts = fmt.formatToParts(d);
    const h = Number(parts.find((p) => p.type === "hour")?.value);
    const m = Number(parts.find((p) => p.type === "minute")?.value);
    if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
    return { hour: h, minute: m };
  } catch (_) {
    return null;
  }
}

function normalizeOverdueNotifyTimeValue(value) {
  const s = String(value || "").trim();
  const m = /^(\d{1,2}):(\d{2})$/.exec(s);
  if (!m) return "09:00";
  const hh = Math.min(23, Math.max(0, Number(m[1])));
  const mm = Math.min(59, Math.max(0, Number(m[2])));
  return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
}

function loadOverdueNotifyRuntime() {
  try {
    const raw = localStorage.getItem(OVERDUE_NOTIFY_RUNTIME_KEY);
    if (!raw) return { lastRunDate: "" };
    const parsed = JSON.parse(raw);
    return {
      lastRunDate: String(parsed?.lastRunDate || "")
    };
  } catch (_) {
    return { lastRunDate: "" };
  }
}

function saveOverdueNotifyRuntime(runtime) {
  try {
    localStorage.setItem(OVERDUE_NOTIFY_RUNTIME_KEY, JSON.stringify({
      lastRunDate: String(runtime?.lastRunDate || "")
    }));
  } catch (_) {
    /* noop */
  }
}

function getTodayStorageDateInTimezone() {
  const p = getCalendarDatePartsInTimeZone(new Date(), getServerTimezone());
  if (!p) return getTodayRuDate();
  return formatDatePartsStorage(p.day, p.month, p.year);
}

function getCurrentMinutesInTimezone() {
  const p = getClockPartsInTimeZone(new Date(), getServerTimezone());
  if (!p) return 0;
  return p.hour * 60 + p.minute;
}

function buildTaskTelegramBodyForOverdue(row) {
  const status = String(row[TASK_COLUMNS.status] || "").trim();
  const templates = displaySettings.taskMessageTemplatesByStatus || {};
  const template = String(templates[status] || "").trim();
  if (!template) return "";
  return String(applyTaskMessageTemplate(template, row) || "").trim();
}

function buildDefaultFullTaskText(row) {
  const taskId = String(row[TASK_COLUMNS.number] || "").trim() || "—";
  const status = String(row[TASK_COLUMNS.status] || "").trim();
  const statusEmoji = TELEGRAM_STATUS_EMOJI[status] || "⚪";
  const lines = [
    `📝 Задача ID:${taskId}: ${String(row[TASK_COLUMNS.task] || "").trim() || "—"}`,
    `🏗️ Название объекта:${String(row[TASK_COLUMNS.object] || "").trim() || "—"}`,
    `📌 Статус:${statusEmoji} ${status || "—"}`,
    `⚡ Приоритет:${String(row[TASK_COLUMNS.priority] || "").trim() || "—"}`,
    `📅 Дата постановки задачи:${String(row[TASK_COLUMNS.addedDate] || "").trim() || "—"}`,
    `🧭 Фаза:${String(row[TASK_COLUMNS.phase] || "").trim() || "—"}`,
    `📂 Раздел:${String(row[TASK_COLUMNS.phaseSection] || "").trim() || "—"}`,
    `🗂️ Подраздел:${String(row[TASK_COLUMNS.phaseSubsection] || "").trim() || "—"}`,
    `👤 Ответственный:${String(row[TASK_COLUMNS.assignedResponsible] || "").trim() || "—"}`,
    `🕴️ Постановщик задачи:${String(row[TASK_COLUMNS.responsible] || "").trim() || "—"}`,
    `⏳ Плановый срок устранения:${String(row[TASK_COLUMNS.dueDate] || "").trim() || "—"}`,
    `💬 Коментарии к задаче:${String(row[TASK_COLUMNS.note] || "").trim() || "—"}`
  ];
  const plan = String(row[TASK_COLUMNS.plan] || "").trim();
  const fact = String(row[TASK_COLUMNS.fact] || "").trim();
  if (plan) lines.push(`🛠️ Комментарии сотрудника (Результат):${plan}`);
  if (fact) lines.push(`🧾 Коментарии администратора:${fact}`);
  return lines.join("\n");
}

function buildOverdueTaskTelegramText(row, overdueDays) {
  const fullBody = buildTaskTelegramBodyForOverdue(row) || buildDefaultFullTaskText(row);
  return [`⚠️ Уведомление о просрочке задачи`, "", fullBody, "", `Просрочено: ${overdueDays} дн.`].join("\n");
}

async function sendOverdueTaskNotification(row, overdueDays, token) {
  const recipientNames = collectTaskTelegramRecipientNames(row);
  if (!recipientNames.length) return { ok: false, reason: "no_recipients" };
  const targets = resolveEmployeeTelegramTargetsByFullNames(recipientNames);
  if (!targets.length) return { ok: false, reason: "no_targets" };
  const text = buildOverdueTaskTelegramText(row, overdueDays);
  const results = await Promise.all(
    targets.map(async (t) => {
      try {
        const response = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            chat_id: t.chatId,
            text
          })
        });
        const apiJson = await response.json().catch(() => ({}));
        return response.ok && apiJson.ok === true;
      } catch (_) {
        return false;
      }
    })
  );
  return { ok: results.some(Boolean), sentCount: results.filter(Boolean).length, total: results.length };
}

async function runOverdueTaskNotificationTick() {
  if (overdueNotifyInFlight) return;
  if (!getAuthToken()) return;
  if (!displaySettings.overdueNotificationsEnabled) return;
  const token = String(displaySettings.telegramBotToken || "").trim();
  if (!token) return;

  const nowMinutes = getCurrentMinutesInTimezone();
  const targetMinutes = (() => {
    const [hh, mm] = normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime).split(":").map(Number);
    return hh * 60 + mm;
  })();
  const today = getTodayStorageDateInTimezone();
  const runtime = loadOverdueNotifyRuntime();
  if (runtime.lastRunDate === today) return;
  if (nowMinutes < targetMinutes) return;

  overdueNotifyInFlight = true;
  try {
    const todayParts = parseRuDateStringToParts(today);
    const rows = getSectionById("tasks")?.rows || [];
    for (const row of rows) {
      const status = String(row[TASK_COLUMNS.status] || "").trim();
      if (status === "Закрыт") continue;
      const due = parseRuDateStringToParts(String(row[TASK_COLUMNS.dueDate] || "").trim());
      if (!due || !todayParts) continue;
      const diff = calendarDiffDays(due, todayParts);
      if (diff <= 0) continue;
      await sendOverdueTaskNotification(row, diff, token);
    }
    runtime.lastRunDate = today;
    saveOverdueNotifyRuntime(runtime);
  } finally {
    overdueNotifyInFlight = false;
  }
}

async function triggerOverdueTaskManualNotifications() {
  if (!getAuthToken()) {
    showStatusDialog({
      title: "Просроченные задачи",
      message: "Ручная отправка доступна после входа в систему.",
      type: "error"
    });
    return false;
  }
  const token = String(displaySettings.telegramBotToken || "").trim();
  if (!token) {
    showStatusDialog({
      title: "Просроченные задачи",
      message: "Укажите токен Telegram-бота в настройках Telegram.",
      type: "error"
    });
    return false;
  }
  const rows = getSectionById("tasks")?.rows || [];
  const today = getTodayStorageDateInTimezone();
  const todayParts = parseRuDateStringToParts(today);
  if (!todayParts) {
    showStatusDialog({
      title: "Просроченные задачи",
      message: "Не удалось определить текущую дату для проверки просрочки.",
      type: "error"
    });
    return false;
  }

  const overdueRows = rows
    .map((row) => {
      const status = String(row[TASK_COLUMNS.status] || "").trim();
      if (status === "Закрыт") return null;
      const due = parseRuDateStringToParts(String(row[TASK_COLUMNS.dueDate] || "").trim());
      if (!due) return null;
      const diff = calendarDiffDays(due, todayParts);
      if (diff <= 0) return null;
      return { row, diff };
    })
    .filter(Boolean);

  if (!overdueRows.length) {
    showStatusDialog({
      title: "Просроченные задачи",
      message: "На сегодня просроченных задач нет.",
      type: "info"
    });
    return true;
  }

  let sentTasks = 0;
  let noRecipients = 0;
  let failed = 0;
  for (const item of overdueRows) {
    const result = await sendOverdueTaskNotification(item.row, item.diff, token);
    if (result.ok) {
      sentTasks += 1;
      continue;
    }
    if (result.reason === "no_recipients" || result.reason === "no_targets") {
      noRecipients += 1;
    } else {
      failed += 1;
    }
  }

  const lines = [
    `Просроченных задач: ${overdueRows.length}.`,
    `Уведомления отправлены по задачам: ${sentTasks}.`
  ];
  if (noRecipients) lines.push(`Без получателей (нет Telegram/Chat ID): ${noRecipients}.`);
  if (failed) lines.push(`Ошибки отправки: ${failed}.`);

  showStatusDialog({
    title: "Просроченные задачи",
    message: lines.join("\n"),
    type: sentTasks > 0 ? "success" : "error"
  });
  return sentTasks > 0;
}

function startOverdueTaskNotificationsScheduler() {
  clearInterval(overdueNotifyTimer);
  overdueNotifyTimer = null;
  if (!isHostedRuntime() || !getAuthToken()) return;
  overdueNotifyTimer = setInterval(() => {
    if (document.hidden) return;
    runOverdueTaskNotificationTick().catch(() => {});
  }, 60000);
  runOverdueTaskNotificationTick().catch(() => {});
}

function stopOverdueTaskNotificationsScheduler() {
  clearInterval(overdueNotifyTimer);
  overdueNotifyTimer = null;
}

function detachTasksChunksObserver() {
  if (window._mbcTaskChunksObserver) {
    try {
      window._mbcTaskChunksObserver.disconnect();
    } catch (_) {
      /* noop */
    }
    window._mbcTaskChunksObserver = null;
  }
}

function renderTable() {
  detachTasksChunksObserver();
  destroyReportCharts();

  if (activeSectionId === "otherSettings") {
    tableContainer.innerHTML = renderOtherSettingsPanel();
    attachOtherSettingsHandlers();
    return;
  }

  if (activeSectionId === "report") {
    tableContainer.innerHTML = renderReportsPanel();
    requestAnimationFrame(() => {
      attachReportCharts();
      attachReportFilterHandlers();
      attachReportPhaseGroupLayoutHandlers();
      attachReportChartTileDragHandlers();
      attachReportExportAndShareHandlers();
      initLucideIcons();
    });
    return;
  }

  const section = sections.find((item) => item.id === activeSectionId);
  if (!section) return;
  cleanupExpiredTrash(section.id);

  const sectionFilters = filtersBySection[section.id] || {};
  let allFilteredEntries = getFilteredRows(section, sectionFilters);
  let tasksListFooterHtml = "";
  let showTasksBackBtn = false;

  if (section.id === "tasks") {
    normalizeTasksListDisplaySettings();
    const browseByObject = displaySettings.tasksListBrowseMode === "byObject";
    if (browseByObject) {
      showTasksBackBtn = tasksBrowseObjectKey !== null;
      if (tasksBrowseObjectKey === null) {
        const grouped = groupTaskEntriesByObject(allFilteredEntries);
        const pickerInner = renderTasksObjectPickerHtml(grouped);
        tableContainer.innerHTML = `
    <section class="table-card table-card--tasks-object-picker">
      <div class="table-header table-header--object-picker-only">
        <h3>${withIcon(getSectionIcon(section.id), section.title)}</h3>
        <div class="table-header-right">
          <button type="button" class="icon-action-btn refresh-section-btn" id="tasksObjectPickerRefreshBtn" title="Обновить">
            <i data-lucide="refresh-cw" class="lucide-icon" aria-hidden="true"></i>
          </button>
        </div>
      </div>
      <div class="tasks-object-picker-shell">
        ${renderTasksScreenModeSwitch(section)}
        <p class="tasks-object-picker-hint">Выберите объект — откроется таблица задач по этому объекту (фильтры и вкладки будут доступны в таблице).</p>
        ${pickerInner}
      </div>
    </section>
  `;
        attachTableActionHandlers(section, []);
        attachTasksObjectPickerHandlers(section);
        initLucideIcons();
        return;
      }
      allFilteredEntries = filterTaskEntriesByObjectKey(allFilteredEntries, tasksBrowseObjectKey);
    }
  }

  let entriesForTbody = allFilteredEntries;
  if (section.id === "tasks") {
    const pageSize = getTasksListPageSize();
    const total = allFilteredEntries.length;
    if (displaySettings.tasksListPagingMode === "chunks") {
      const vis = Math.min(tasksListVisibleCount, total);
      entriesForTbody = allFilteredEntries.slice(0, vis);
      tasksListFooterHtml = renderTasksListFooterHtml("chunks", pageSize, 1, total, entriesForTbody.length);
    } else {
      const totalPages = Math.max(1, Math.ceil(total / pageSize) || 1);
      if (tasksListPage > totalPages) tasksListPage = totalPages;
      if (tasksListPage < 1) tasksListPage = 1;
      const start = (tasksListPage - 1) * pageSize;
      entriesForTbody = allFilteredEntries.slice(start, start + pageSize);
      tasksListFooterHtml = renderTasksListFooterHtml("pagination", pageSize, tasksListPage, total, entriesForTbody.length);
    }
  }

  let tasksChunksSentinelHtml = "";
  if (
    section.id === "tasks" &&
    displaySettings.tasksListPagingMode === "chunks" &&
    entriesForTbody.length < allFilteredEntries.length
  ) {
    tasksChunksSentinelHtml = `<div class="tasks-list-scroll-sentinel" id="tasksChunksScrollSentinel" data-shown="${entriesForTbody.length}" data-all="${allFilteredEntries.length}" aria-hidden="true"></div>`;
  }

  const sectionGroupTabs = renderSectionGroupTabs(section.id);

  const visibleColumnIndexes = getVisibleColumnIndexes(section);
  const isTrashView = isTrashTab(section.id);
  const selectedRows = getSelectedRowsSet(getSelectionKey(section.id));
  const selectedCount = selectedRows.size;
  const isAllFilteredSelected = allFilteredEntries.length > 0
    && allFilteredEntries.every((entry) => selectedRows.has(entry.rowIndex));

  const showHeaderNumbers = headerNumberingBySection[section.id] !== false;
  const headRowspan = showHeaderNumbers ? ` rowspan="2"` : "";
  const trashHeadersMain = isTrashView
    ? `<th class="trash-meta-col"${headRowspan}>Удалено</th><th class="trash-meta-col"${headRowspan}>До удаления</th>`
    : "";
  const titleHeaderCells = visibleColumnIndexes.map((columnIndex, viewOrder) => {
    const column = section.columns[columnIndex];
    const numberClass = columnIndex === 0 && viewOrder === 0 ? "number-col" : "";
    const rolesNumberClass = section.id === "roles" && columnIndex === 0 && viewOrder === 0 ? "roles-number-col" : "";
    const statusClass = columnIndex === TASK_COLUMNS.status ? "status-col" : "";
    const objectClass = columnIndex === TASK_COLUMNS.object ? "object-col" : "";
    const mediaClass = isMediaColumn(columnIndex) ? "media-col" : "";
    const objectPhotoClass = section.id === "objects" && columnIndex === OBJECT_COLUMNS.photo ? "object-photo-col" : "";
    return `<th class="${numberClass} ${rolesNumberClass} ${statusClass} ${objectClass} ${mediaClass} ${objectPhotoClass}">
      <span class="table-th-title">${escapeHtmlText(column)}</span>
    </th>`;
  }).join("");
  const orderHeaderRow = showHeaderNumbers
    ? `<tr class="table-head-order-row">
        ${visibleColumnIndexes.map((_, viewOrder) => `<th class="table-order-cell"><span class="table-th-order">${viewOrder + 1}</span></th>`).join("")}
      </tr>`
    : "";

  const thead = `
    <thead class="${showHeaderNumbers ? "table-head-has-order" : ""}">
      <tr class="table-head-main-row">
        <th class="checkbox-col ${section.id === "roles" ? "roles-compact-col" : ""}"${headRowspan}>
          <input type="checkbox" id="selectAllRows" ${isAllFilteredSelected ? "checked" : ""} />
        </th>
        ${titleHeaderCells}
        ${trashHeadersMain}
        <th class="actions-col ${section.id === "roles" ? "roles-compact-actions-col" : ""}"${headRowspan}>Действие</th>
      </tr>
      ${orderHeaderRow}
    </thead>
  `;

  const tbody = `
    <tbody>
      ${entriesForTbody
        .map((entry) => {
          const rowCells = visibleColumnIndexes
            .map((colIndex, viewOrder) => {
              const cell = entry.row[colIndex];
              const stickyClass = colIndex === 0 && viewOrder === 0 ? "number-col" : "";
              const rolesNumberClass = section.id === "roles" && colIndex === 0 && viewOrder === 0 ? "roles-number-col" : "";
              const statusClass = colIndex === TASK_COLUMNS.status ? "status-col" : "";
              const objectClass = colIndex === TASK_COLUMNS.object ? "object-col" : "";
              const mediaClass = isMediaColumn(colIndex) ? "media-col" : "";
              const objectPhotoClass = section.id === "objects" && colIndex === OBJECT_COLUMNS.photo ? "object-photo-col" : "";
              const wideClass = getWideColumnClass(colIndex);
              const readonlyClass = isReadonlyColumn(section, colIndex) ? "readonly-cell" : "";
              return `<td class="editable-cell ${stickyClass} ${rolesNumberClass} ${statusClass} ${objectClass} ${mediaClass} ${objectPhotoClass} ${wideClass} ${readonlyClass}" data-row-index="${entry.rowIndex}" data-col-index="${colIndex}">${renderCellContent(section, entry.row, colIndex, cell, entry.rowIndex)}</td>`;
            })
            .join("");
          const rowFocusClass = activeRowBySection[section.id] === entry.rowIndex ? "focused-row" : "";
          const rowHighlightClass = getRowHighlightClass(section, entry.row);
          const trashMetaCells = isTrashView
            ? `<td class="trash-meta-col">${formatTrashDate(entry.deletedAt)}</td><td class="trash-meta-col">${formatTrashRemaining(entry.expiresAt)}</td>`
            : "";
          const accordionRow =
            section.id === "tasks" ? renderTaskAssigneesAccordionRows(entry.row, visibleColumnIndexes, isTrashView) : "";
          return `
            <tr class="${rowFocusClass} ${rowHighlightClass}">
              <td class="checkbox-col ${section.id === "roles" ? "roles-compact-col" : ""}">
                <input type="checkbox" class="row-checkbox" data-row-index="${entry.rowIndex}" ${selectedRows.has(entry.rowIndex) ? "checked" : ""} />
              </td>
              ${rowCells}
              ${trashMetaCells}
              <td class="actions-col ${section.id === "roles" ? "roles-compact-actions-col" : ""}">
                ${renderRowActions(section.id, isTrashView, entry.rowIndex, entry.row)}
              </td>
            </tr>
            ${accordionRow}
          `;
        })
        .join("") || `<tr><td colspan="${visibleColumnIndexes.length + (isTrashView ? 4 : 2)}" class="empty-state">Нет данных по выбранным фильтрам</td></tr>`}
    </tbody>
  `;

  const tableHeaderIconButtons = `
        <div class="table-header-right">
          <button type="button" class="icon-action-btn add-row-btn" id="addRowBtn" title="Добавить">
            <i data-lucide="plus" class="lucide-icon" aria-hidden="true"></i>
          </button>
          <button type="button" class="icon-action-btn filter-toggle-btn" id="toggleFiltersBtn" title="Фильтр">
            <i data-lucide="filter" class="lucide-icon" aria-hidden="true"></i>
          </button>
          ${section.id === "tasks" ? `
          <button type="button" class="icon-action-btn import-tasks-btn" id="openTaskImportModalBtn" title="Импорт задач">
            <i data-lucide="file-up" class="lucide-icon" aria-hidden="true"></i>
          </button>` : ""}
          ${section.id === "tasks" ? `
          <button type="button" class="icon-action-btn" id="googleSheetsSyncTasksBtn" title="Синхронизация">
            <span class="gs-mini-icon" aria-hidden="true">
              <svg viewBox="0 0 48 48" fill="none">
                <path d="M8 5h20l12 12v24a2 2 0 0 1-2 2H10a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2z" fill="#0F9D58"/>
                <path d="M28 5v10a2 2 0 0 0 2 2h10L28 5z" fill="#B7E1CD"/>
                <rect x="14" y="22" width="20" height="3" rx="1.5" fill="#fff"/>
                <rect x="14" y="28" width="20" height="3" rx="1.5" fill="#fff"/>
                <rect x="14" y="34" width="12" height="3" rx="1.5" fill="#fff"/>
              </svg>
            </span>
          </button>` : ""}
          ${section.id === "tasks" ? `
          <button type="button" class="icon-action-btn" id="sendOverdueTasksBtn" title="Отправить просроченные">
            <i data-lucide="bell-ring" class="lucide-icon" aria-hidden="true"></i>
          </button>` : ""}
          <button type="button" class="icon-action-btn refresh-section-btn" id="refreshSectionBtn" title="Обновить">
            <i data-lucide="refresh-cw" class="lucide-icon" aria-hidden="true"></i>
          </button>
          <button type="button" class="icon-action-btn export-open-btn" id="openExportModalBtn" title="Скачать">
            <i data-lucide="download" class="lucide-icon" aria-hidden="true"></i>
          </button>
          <button type="button" class="icon-action-btn settings-toggle-btn" id="toggleTableSettingsBtn" title="Настройки таблицы">
            <i data-lucide="settings" class="lucide-icon" aria-hidden="true"></i>
          </button>
        </div>`;

  const sectionTitleH3 = `<h3>${withIcon(getSectionIcon(section.id), section.title)}</h3>`;

  const tasksDrilldownHeader =
    section.id === "tasks" && showTasksBackBtn
      ? `
      <div class="table-header table-header--tasks-drilldown">
        <div class="table-header-drill-left">
          <button type="button" class="secondary tasks-back-objects-btn" id="tasksBackToObjectsBtn">← К объектам</button>
          ${sectionTitleH3}
        </div>
        ${tableHeaderIconButtons}
      </div>`
      : `
      <div class="table-header">
        ${sectionTitleH3}
        ${tableHeaderIconButtons}
      </div>`;

  tableContainer.innerHTML = `
    <section class="table-card${section.id === "tasks" && showTasksBackBtn ? " table-card--tasks-drilldown" : ""}">
      ${tasksDrilldownHeader}
      ${sectionGroupTabs}
      ${renderBulkActions(selectedCount, isTrashView, section.id)}
      ${renderStatusTabs(section)}
      ${renderTasksScreenModeSwitch(section)}
      ${renderFilters(section, sectionFilters, filterPanelOpenBySection[section.id] === true)}
      <div class="table-wrap">
        <table>
          ${thead}
          ${tbody}
        </table>
        ${tasksChunksSentinelHtml}
      </div>
      ${tasksListFooterHtml}
    </section>
  `;
  attachFilterHandlers(section);
  attachTableActionHandlers(section, allFilteredEntries);
  attachEditableCellHandlers(section);
  attachHeaderActionHandlers(section, allFilteredEntries);
  attachMediaSlotHandlers(section);
  attachObjectPhotoHandlers(section);
  if (section.id === "tasks") {
    attachTaskAccordionHandlers(section);
    attachTasksListFooterHandlers(section);
    attachTasksObjectPickerHandlers(section);
  }
  initLucideIcons();
  updateTableStickyHeaderOffsets();
}

function getSectionGroupBySectionId(sectionId) {
  const groups = Object.values(SECTION_GROUPS);
  return groups.find((group) => group.sections.includes(sectionId)) || null;
}

function renderSectionGroupTabs(sectionId) {
  const group = getSectionGroupBySectionId(sectionId);
  if (!group) return "";
  const tabsHtml = group.sections.map((id) => {
    const section = getSectionById(id);
    if (!section) return "";
    const isActive = id === sectionId;
    return `<button type="button" class="section-subtab-btn ${isActive ? "active" : ""}" data-section-tab="${id}">${section.title}</button>`;
  }).join("");
  return `<div class="section-subtabs-row">${tabsHtml}</div>`;
}

function updateTableStickyHeaderOffsets() {
  const table = document.querySelector(".table-wrap table");
  if (!table) return;
  const mainHeadRow = table.querySelector("thead.table-head-has-order .table-head-main-row");
  if (!mainHeadRow) {
    table.style.removeProperty("--table-sticky-main-row-h");
    return;
  }
  const height = Math.max(1, Math.ceil(mainHeadRow.getBoundingClientRect().height));
  table.style.setProperty("--table-sticky-main-row-h", `${height}px`);
}

function formatTrashDate(ts) {
  if (!ts) return "-";
  const d = new Date(Number(ts));
  if (Number.isNaN(d.getTime())) return "-";
  const tz = getServerTimezone();
  const use12 = normalizeTimeDisplayFormatId(displaySettings.timeDisplayFormat) === "12";
  const showSec = displaySettings.timeShowSeconds === true;
  try {
    return new Intl.DateTimeFormat("ru-RU", {
      timeZone: tz,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      ...(showSec ? { second: "2-digit" } : {}),
      hour12: use12
    }).format(d);
  } catch (_) {
    return d.toLocaleString("ru-RU");
  }
}

function formatTrashRemaining(expiresAt) {
  if (!expiresAt) return "-";
  const now = Date.now();
  const diff = Math.ceil((expiresAt - now) / 86400000);
  if (diff <= 0) return "истек";
  return `${diff} дн.`;
}

function renderBulkActions(selectedCount, isTrashView, sectionId) {
  if (selectedCount <= 1) return "";
  if (isTrashView) {
    return `
      <div class="bulk-actions-bar">
        <span class="bulk-actions-count">В корзине выбрано: ${selectedCount}</span>
        <button type="button" class="icon-action-btn bulk-restore-btn" id="bulkRestoreBtn" title="Восстановить выбранные">
          <i data-lucide="rotate-ccw" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  if (sectionId !== "tasks") {
    return `
      <div class="bulk-actions-bar">
        <span class="bulk-actions-count">Выбрано: ${selectedCount}</span>
        <button type="button" class="icon-action-btn danger-btn bulk-delete-btn" id="bulkDeleteBtn" title="Удалить выбранные">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  return `
    <div class="bulk-actions-bar">
      <span class="bulk-actions-count">Выбрано: ${selectedCount}</span>
      <button type="button" class="icon-action-btn bulk-send-btn" id="bulkSendBtn" title="Отправить выбранные">
        <i data-lucide="send" class="lucide-icon" aria-hidden="true"></i>
      </button>
      <button type="button" class="icon-action-btn danger-btn bulk-delete-btn" id="bulkDeleteBtn" title="Удалить выбранные">
        <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
      </button>
    </div>
  `;
}

function renderRowActions(sectionId, isTrashView, rowIndex, row) {
  if (isTrashView) {
    return `
      <div class="action-buttons">
        <button type="button" class="icon-action-btn restore-row-btn" title="Восстановить" data-row-index="${rowIndex}">
          <i data-lucide="rotate-ccw" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  if (sectionId === "employees") {
    return `
      <div class="action-buttons">
        <button type="button" class="icon-action-btn copy-employee-msg-btn" title="Скопировать сообщение сотруднику" data-row-index="${rowIndex}">
          <i data-lucide="copy" class="lucide-icon" aria-hidden="true"></i>
        </button>
        <button type="button" class="icon-action-btn danger-btn delete-row-btn" title="Удалить" data-row-index="${rowIndex}">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  if (["data", "phases", "phaseSections", "phaseSubsections"].includes(sectionId)) {
    return `
      <div class="action-buttons">
        <button type="button" class="icon-action-btn danger-btn delete-row-btn" title="Удалить" data-row-index="${rowIndex}">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  if (sectionId === "roles") {
    const roleName = String(row?.[1] || "").trim();
    const isProtected = SYSTEM_ROLES.includes(roleName);
    const disabledAttr = isProtected ? "disabled" : "";
    const title = isProtected ? "Системная должность (нельзя удалить)" : "Удалить";
    return `
      <div class="action-buttons">
        <button type="button" class="icon-action-btn danger-btn delete-row-btn" title="${title}" data-row-index="${rowIndex}" ${disabledAttr} data-protected="${isProtected ? "1" : "0"}">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  if (sectionId === "departments") {
    const departmentName = String(row?.[1] || "").trim();
    const isProtected = SYSTEM_DEPARTMENTS.includes(departmentName);
    const disabledAttr = isProtected ? "disabled" : "";
    const title = isProtected ? "Системный отдел (нельзя удалить)" : "Удалить";
    return `
      <div class="action-buttons">
        <button type="button" class="icon-action-btn danger-btn delete-row-btn" title="${title}" data-row-index="${rowIndex}" ${disabledAttr} data-protected="${isProtected ? "1" : "0"}">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  if (sectionId === "objects") {
    return `
      <div class="action-buttons">
        <button type="button" class="icon-action-btn danger-btn delete-row-btn" title="Удалить" data-row-index="${rowIndex}">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      </div>
    `;
  }
  return `
    <div class="action-buttons">
      <button type="button" class="icon-action-btn view-row-btn" title="Просмотр" data-row-index="${rowIndex}">
        <i data-lucide="eye" class="lucide-icon" aria-hidden="true"></i>
      </button>
      <button type="button" class="icon-action-btn send-row-btn" title="Отправить" data-row-index="${rowIndex}">
        <i data-lucide="send" class="lucide-icon" aria-hidden="true"></i>
      </button>
      <button type="button" class="icon-action-btn danger-btn delete-row-btn" title="Удалить" data-row-index="${rowIndex}">
        <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
      </button>
    </div>
  `;
}

function getUniqueValues(rows, index) {
  const values = new Set(rows.map((row) => String(row[index] || "").trim()).filter(Boolean));
  return Array.from(values);
}

function getAllTaskRows() {
  return getSectionById("tasks")?.rows || [];
}

function getReportDistinctObjects() {
  return getUniqueValues(getAllTaskRows(), TASK_COLUMNS.object).sort((a, b) => a.localeCompare(b, "ru"));
}

function getReportDistinctEmployees() {
  const rows = getAllTaskRows();
  const set = new Set();
  rows.forEach((row) => {
    const r = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
    set.add(r || "— не назначен —");
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, "ru"));
}

function sortedCountEntries(obj) {
  return Object.entries(obj).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0]), "ru"));
}

/** Пастельная палитра для не-статусных рядов */
const REPORT_CHART_COLORS = [
  "#b4c5f5",
  "#e8d4b0",
  "#c9c5eb",
  "#b8d9e8",
  "#f0c9a8",
  "#b8e0c8",
  "#e0b8c8",
  "#d4d0c4"
];

function parseHtmlDateValue(s) {
  const t = String(s || "").trim();
  if (!t) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  if (!year || !month || !day) return null;
  return { year, month, day };
}

function compareDateParts(a, b) {
  if (!a || !b) return 0;
  return a.year * 10000 + a.month * 100 + a.day - (b.year * 10000 + b.month * 100 + b.day);
}

function colorForStatusLabel(label) {
  return STATUS_CHART_COLORS[label] || STATUS_CHART_COLORS.Прочее;
}

function statusAccentForFilter(st) {
  return STATUS_CHART_COLORS_ACCENT[st] || STATUS_CHART_COLORS_ACCENT[REPORT_NO_STATUS_LABEL];
}

function statusBgForFilter(st) {
  return STATUS_CHART_COLORS[st] || STATUS_CHART_COLORS[REPORT_NO_STATUS_LABEL];
}

/** Высота области графика с вертикальной прокруткой при большом числе категорий */
function setReportScrollableChartHeight(wrapEl, barCount, opts = {}) {
  if (!wrapEl) return;
  const rowPx = opts.rowPx ?? 26;
  const padPx = opts.paddingPx ?? 88;
  const maxPx = opts.maxPx ?? 1200;
  const n = Math.max(1, Number(barCount) || 0);
  const contentPx = n * rowPx + padPx;
  const minPx = opts.minPx !== undefined ? opts.minPx : 260;
  const h = Math.min(maxPx, Math.max(minPx, contentPx));
  wrapEl.style.height = `${Math.round(h)}px`;
}

/** Пиксели высоты обёртки горизонтального bar (фазы / разделы / подразделы). */
function getReportPhaseHBarWrapHeightPx(barCount, opts = {}) {
  const rowPx = opts.rowPx ?? 14;
  const padPx = opts.paddingPx ?? 52;
  const maxPx = opts.maxPx ?? 900;
  const minPx = opts.minPx ?? 72;
  const n = Math.max(1, Number(barCount) || 0);
  return Math.min(maxPx, Math.max(minPx, n * rowPx + padPx));
}

/** Компактная высота под горизонтальные bar: высота ≈ число категорий × тонкая строка (без лишнего min). */
function setReportPhaseHBarWrapHeight(wrapEl, barCount, opts = {}) {
  if (!wrapEl) return;
  const h = getReportPhaseHBarWrapHeightPx(barCount, opts);
  wrapEl.style.height = `${Math.round(h)}px`;
}

/** В режиме «в одну строку» — одинаковая высота трёх графиков по max(категорий), иначе разная раскладка Chart.js. */
function applyReportPhaseTrioWrapHeights(phLen, psecLen, psubLen) {
  const wrapPh = document.getElementById("reportChartPhaseWrap");
  const wrapSec = document.getElementById("reportChartPhaseSectionWrap");
  const wrapSub = document.getElementById("reportChartPhaseSubsectionWrap");
  // Для stacked-графиков нужна большая высота: место под легенду + читабельные подписи категорий.
  const opts = { maxPx: 1100, minPx: 210, rowPx: 24, paddingPx: 108 };
  if (loadReportPhaseGroupLayout() === "row") {
    const nMax = Math.max(phLen, psecLen, psubLen, 1);
    const h = Math.round(getReportPhaseHBarWrapHeightPx(nMax, opts));
    [wrapPh, wrapSec, wrapSub].forEach((el) => {
      if (el) el.style.height = `${h}px`;
    });
  } else {
    setReportPhaseHBarWrapHeight(wrapPh, phLen, opts);
    setReportPhaseHBarWrapHeight(wrapSec, psecLen, opts);
    setReportPhaseHBarWrapHeight(wrapSub, psubLen, opts);
  }
}

function reportHexToRgba(hex, alpha) {
  const h = String(hex || "")
    .replace("#", "")
    .trim();
  if (h.length !== 6) return `rgba(100, 116, 139, ${alpha})`;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

/** Вертикальные столбцы: низ прозрачнее, верх насыщеннее */
function reportBarGradientVertical(chart, baseColor) {
  const { ctx, chartArea } = chart;
  if (!chartArea) return baseColor;
  const g = ctx.createLinearGradient(0, chartArea.bottom, 0, chartArea.top);
  g.addColorStop(0, reportHexToRgba(baseColor, 0.06));
  g.addColorStop(0.55, reportHexToRgba(baseColor, 0.72));
  g.addColorStop(1, baseColor);
  return g;
}

/** Общие опции тонких горизонтальных bar (фазы / разделы / подразделы), много категорий */
const REPORT_HBAR_OPTIONS_THIN = {
  datasets: {
    bar: {
      maxBarThickness: 9,
      categoryPercentage: 0.88,
      barPercentage: 0.62
    }
  },
  /** Отступ справа — подписи datalabels у макс. значения не обрезаются контейнером */
  layout: {
    padding: {
      right: 36,
      left: 2,
      top: 4,
      bottom: 4
    }
  }
};

const REPORT_HBAR_SCALE_X = {
  beginAtZero: true,
  ticks: { precision: 0 },
  grace: "10%"
};

/** Запас по оси значений (вертикальные столбцы, линии по месяцам) — подписи сверху не режутся */
const REPORT_VALUE_AXIS_GRACE = { grace: "10%" };

/** Общие отступы области графика для всех отчётных чартов с datalabels */
const REPORT_CHART_LABEL_SAFE_LAYOUT = {
  layout: {
    padding: {
      top: 12,
      right: 36,
      bottom: 10,
      left: 8
    }
  }
};

/** Горизонтальные столбцы: у оси прозрачнее, к концу полосы насыщеннее */
function reportBarGradientHorizontal(chart, baseColor) {
  const { ctx, chartArea } = chart;
  if (!chartArea) return baseColor;
  const g = ctx.createLinearGradient(chartArea.left, 0, chartArea.right, 0);
  g.addColorStop(0, reportHexToRgba(baseColor, 0.06));
  g.addColorStop(0.55, reportHexToRgba(baseColor, 0.72));
  g.addColorStop(1, baseColor);
  return g;
}

function getDefaultReportChartOrder() {
  return Object.keys(REPORT_CHART_TILE_META);
}

function loadReportChartOrder() {
  try {
    const raw = localStorage.getItem(REPORT_CHART_ORDER_STORAGE_KEY);
    if (!raw) return getDefaultReportChartOrder();
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return getDefaultReportChartOrder();
    const allowed = new Set(getDefaultReportChartOrder());
    const seen = new Set();
    const out = [];
    for (const id of parsed) {
      if (allowed.has(id) && !seen.has(id)) {
        seen.add(id);
        out.push(id);
      }
    }
    for (const id of getDefaultReportChartOrder()) {
      if (!seen.has(id)) out.push(id);
    }
    return out;
  } catch (_) {
    return getDefaultReportChartOrder();
  }
}

function saveReportChartOrder(order) {
  try {
    localStorage.setItem(REPORT_CHART_ORDER_STORAGE_KEY, JSON.stringify(order));
    scheduleServerSync();
  } catch (_) {
    /* noop */
  }
}

function loadReportPhaseGroupLayout() {
  try {
    const v = localStorage.getItem(REPORT_PHASE_GROUP_LAYOUT_KEY);
    if (v === "row" || v === "separate") return v;
  } catch (_) {
    /* noop */
  }
  return "separate";
}

function saveReportPhaseGroupLayout(mode) {
  try {
    if (mode === "row" || mode === "separate") {
      localStorage.setItem(REPORT_PHASE_GROUP_LAYOUT_KEY, mode);
      scheduleServerSync();
    }
  } catch (_) {
    /* noop */
  }
}

function reorderReportChartOrder(order, dragId, targetId) {
  const o = [...order];
  const from = o.indexOf(dragId);
  const to = o.indexOf(targetId);
  if (from === -1 || to === -1 || from === to) return order;
  o.splice(from, 1);
  const insertAt = o.indexOf(targetId);
  if (insertAt === -1) return order;
  o.splice(insertAt, 0, dragId);
  return o;
}

function renderReportChartTileFragment(id, opts = {}) {
  const meta = REPORT_CHART_TILE_META[id];
  if (!meta) return "";
  const wide = opts.inPhaseRow ? "" : meta.wide ? " report-tile-wide" : "";
  const handle = sharedReportMode
    ? ""
    : `<button type="button" class="report-chart-drag-handle" draggable="true" data-drag-chart-id="${escapeHtmlAttr(id)}" title="Переместить график" aria-label="Перетащите, чтобы изменить порядок графиков"><i data-lucide="move" class="lucide-icon" aria-hidden="true"></i></button>`;
  const bodies = {
    status: `<h4>По статусам</h4><div class="report-canvas-wrap report-canvas-wrap--donut"><canvas id="reportChartStatus"></canvas></div>`,
    priority: `<h4>По приоритету</h4><div class="report-canvas-wrap report-canvas-wrap--donut"><canvas id="reportChartPriority"></canvas></div>`,
    priorityDonut: `<h4>Приоритеты</h4><div class="report-canvas-wrap report-canvas-wrap--donut"><canvas id="reportChartPriorityDonut"></canvas></div>`,
    months: `<h4>Добавлено и закрыто по месяцам <span class="report-tile-year">${new Date().getFullYear()}</span></h4><div class="report-canvas-wrap report-canvas-tall"><canvas id="reportChartMonths"></canvas></div>`,
    overdue: `<h4>Просроченные задачи <span class="report-tile-note">(по объектам, топ)</span></h4><div class="report-canvas-wrap report-canvas-scroll" id="reportChartOverdueWrap"><canvas id="reportChartOverdue"></canvas></div>`,
    phase: `<h4>Топ фаз</h4><div class="report-canvas-wrap report-canvas-wrap--phase-h" id="reportChartPhaseWrap"><canvas id="reportChartPhase"></canvas></div>`,
    phaseDonut: `<h4>Фазы</h4><div class="report-canvas-wrap report-canvas-wrap--donut"><canvas id="reportChartPhaseDonut"></canvas></div>`,
    phaseSection: `<h4>Разделы</h4><div class="report-canvas-wrap report-canvas-wrap--phase-h report-canvas-scroll" id="reportChartPhaseSectionWrap"><canvas id="reportChartPhaseSection"></canvas></div>`,
    phaseSectionDonut: `<h4>Разделы</h4><div class="report-canvas-wrap report-canvas-wrap--donut"><canvas id="reportChartPhaseSectionDonut"></canvas></div>`,
    phaseSubsection: `<h4>Подразделы</h4><div class="report-canvas-wrap report-canvas-wrap--phase-h report-canvas-scroll" id="reportChartPhaseSubsectionWrap"><canvas id="reportChartPhaseSubsection"></canvas></div>`,
    phaseSubsectionDonut: `<h4>Подразделы</h4><div class="report-canvas-wrap report-canvas-wrap--donut"><canvas id="reportChartPhaseSubsectionDonut"></canvas></div>`,
    object: `<h4>Объекты <span class="report-tile-note">(все, по убыванию; прокрутка)</span></h4><div class="report-canvas-wrap report-canvas-scroll" id="reportChartObjectWrap"><canvas id="reportChartObject"></canvas></div>`,
    department: `<h4>По отделам <span class="report-tile-note">(исполнитель → отдел из справочника)</span></h4><div class="report-canvas-wrap report-canvas-scroll" id="reportChartDepartmentWrap"><canvas id="reportChartDepartment"></canvas></div>`,
    responsible: `<h4>Исполнители <span class="report-tile-note">(по статусам, топ)</span></h4><div class="report-canvas-wrap report-canvas-scroll" id="reportChartResponsibleWrap"><canvas id="reportChartResponsible"></canvas></div>`
  };
  const inner = bodies[id];
  if (!inner) return "";
  return `<div class="report-tile${wide}" data-report-chart="${escapeHtmlAttr(id)}">${handle}${inner}</div>`;
}

function renderReportChartsGridHtml() {
  const order = loadReportChartOrder();
  const topTrio = REPORT_TOP_TRIO_IDS.filter((id) => order.includes(id));
  const phaseBars = REPORT_PHASE_GROUP_IDS.filter((id) => order.includes(id));
  const phaseDonuts = REPORT_PHASE_DONUT_GROUP_IDS.filter((id) => order.includes(id));

  const groupedIds = new Set([...topTrio, ...phaseBars, ...phaseDonuts]);
  const rest = order.filter((id) => !groupedIds.has(id));

  const parts = [];
  if (topTrio.length) {
    parts.push(
      `<div class="report-phase-group-row report-triple-group-row">${topTrio.map((id) =>
        renderReportChartTileFragment(id, { inPhaseRow: true })
      ).join("")}</div>`
    );
  }
  if (phaseBars.length) {
    parts.push(
      `<div class="report-phase-group-row report-phase-bars-row">${phaseBars.map((id) =>
        renderReportChartTileFragment(id, { inPhaseRow: true })
      ).join("")}</div>`
    );
  }
  if (phaseDonuts.length) {
    parts.push(
      `<div class="report-phase-group-row report-phase-donuts-row">${phaseDonuts.map((id) =>
        renderReportChartTileFragment(id, { inPhaseRow: true })
      ).join("")}</div>`
    );
  }

  parts.push(...rest.map((id) => renderReportChartTileFragment(id)));
  return parts.join("");
}

function attachReportChartTileDragHandlers() {
  if (isSharedReportView()) return;
  const grid = tableContainer?.querySelector(".report-grid");
  if (!grid) return;

  const highlightTarget = (targetEl) => {
    grid.querySelectorAll(".report-tile--drop-target").forEach((el) => el.classList.remove("report-tile--drop-target"));
    const tile = targetEl?.closest?.(".report-tile[data-report-chart]");
    if (!tile) return;
    const tid = String(tile.getAttribute("data-report-chart") || "");
    if (reportChartDragId && tid && reportChartDragId !== tid) {
      tile.classList.add("report-tile--drop-target");
    }
  };

  grid.querySelectorAll(".report-chart-drag-handle").forEach((btn) => {
    btn.addEventListener("dragstart", (e) => {
      const id = String(btn.getAttribute("data-drag-chart-id") || "");
      reportChartDragId = id;
      e.dataTransfer.setData("application/x-mbc-chart", id);
      e.dataTransfer.setData("text/plain", id);
      e.dataTransfer.effectAllowed = "move";
      btn.closest(".report-tile")?.classList.add("report-tile--dragging");
    });
    btn.addEventListener("dragend", () => {
      reportChartDragId = null;
      grid.querySelectorAll(".report-tile").forEach((t) => {
        t.classList.remove("report-tile--dragging", "report-tile--drop-target");
      });
    });
  });

  grid.addEventListener("dragover", (e) => {
    if (!reportChartDragId) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    highlightTarget(e.target);
  });

  grid.addEventListener("drop", (e) => {
    e.preventDefault();
    const tile = e.target.closest?.(".report-tile[data-report-chart]");
    grid.querySelectorAll(".report-tile--drop-target").forEach((el) => el.classList.remove("report-tile--drop-target"));
    const fromId = String(e.dataTransfer.getData("application/x-mbc-chart") || reportChartDragId || "");
    const toId = tile ? String(tile.getAttribute("data-report-chart") || "") : "";
    if (!fromId || !toId || fromId === toId) return;
    const cur = loadReportChartOrder();
    const next = reorderReportChartOrder(cur, fromId, toId);
    saveReportChartOrder(next);
    refreshReportView();
  });
}

function getReportFilteredRows() {
  if (reportShareRowsOverride) {
    return reportShareRowsOverride;
  }
  const rows = getSectionById("tasks")?.rows || [];
  const allowed =
    reportStatusFilter === null ? new Set(REPORT_FILTER_ALL_STATUSES) : reportStatusFilter;
  const allStatusesSelected =
    reportStatusFilter === null ||
    (reportStatusFilter.size === STATUS_OPTIONS.length && STATUS_OPTIONS.every((s) => reportStatusFilter.has(s)));
  const fromParts = parseHtmlDateValue(reportDateFrom);
  const toParts = parseHtmlDateValue(reportDateTo);
  const useDate = !!(fromParts || toParts);

  return rows.filter((row) => {
    const stRaw = String(row[TASK_COLUMNS.status] || "").trim();
    if (!stRaw) {
      if (!allStatusesSelected) return false;
    } else if (!allowed.has(stRaw)) {
      return false;
    }
    if (reportFilterObject) {
      const ob = String(row[TASK_COLUMNS.object] || "").trim();
      if (ob !== reportFilterObject) return false;
    }
    if (reportFilterEmployee) {
      const resp = String(row[TASK_COLUMNS.assignedResponsible] || "").trim() || "— не назначен —";
      if (resp !== reportFilterEmployee) return false;
    }
    if (!useDate) return true;
    const addD = String(row[TASK_COLUMNS.addedDate] || "").trim();
    const rp = parseRuDateStringToParts(addD);
    if (!rp) return false;
    if (fromParts && compareDateParts(rp, fromParts) < 0) return false;
    if (toParts && compareDateParts(rp, toParts) > 0) return false;
    return true;
  });
}

function sortStatusEntriesForChart(statusCounts) {
  const order = [...STATUS_OPTIONS, REPORT_NO_STATUS_LABEL];
  return Object.entries(statusCounts).sort((a, b) => {
    const ia = order.indexOf(a[0]);
    const ib = order.indexOf(b[0]);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return String(a[0]).localeCompare(String(b[0]), "ru");
  });
}

/** ФИО (нормализованное) → отдел из справочника сотрудников; при дубликате ФИО — последняя строка. */
function buildEmployeeNameToDepartmentMap() {
  const map = new Map();
  const empRows = getSectionById("employees")?.rows || [];
  for (const er of empRows) {
    const fn = normalizePersonName(er[EMPLOYEE_COLUMNS.fullName]);
    if (!fn) continue;
    const dept = String(er[EMPLOYEE_COLUMNS.department] || "").trim() || "— без отдела —";
    map.set(fn, dept);
  }
  return map;
}

function buildResponsibleStatusRows(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const name = String(row[TASK_COLUMNS.assignedResponsible] || "").trim() || "— не назначен —";
    const stRaw = String(row[TASK_COLUMNS.status] || "").trim();
    if (!STATUS_OPTIONS.includes(stRaw)) return;
    if (!map.has(name)) {
      map.set(name, { counts: {} });
    }
    const rec = map.get(name);
    rec.counts[stRaw] = (rec.counts[stRaw] || 0) + 1;
  });
  return Array.from(map.entries())
    .map(([name, r]) => {
      const total = STATUS_OPTIONS.reduce((s, c) => s + (r.counts[c] || 0), 0);
      return { name, counts: r.counts, total };
    })
    .filter((r) => r.total > 0)
    .sort((a, b) => b.total - a.total);
}

function getReportYearMonthChartSeries(monthCounts, year) {
  const y = typeof year === "number" ? year : new Date().getFullYear();
  const data = [];
  for (let m = 1; m <= 12; m++) {
    const key = `${y}-${String(m).padStart(2, "0")}`;
    data.push(monthCounts[key] || 0);
  }
  return { year: y, labels: [...REPORT_MONTH_LABELS_RU], data };
}

/** Просрочка: не закрыта и срок раньше сегодняшней даты (по календарным дням). */
function isTaskOverdueForReport(row) {
  const st = String(row[TASK_COLUMNS.status] || "").trim();
  if (st === "Закрыт") return false;
  const due = parseRuDateStringToParts(String(row[TASK_COLUMNS.dueDate] || "").trim());
  if (!due) return false;
  const today = new Date();
  const todayParts = { year: today.getFullYear(), month: today.getMonth() + 1, day: today.getDate() };
  return compareDateParts(due, todayParts) < 0;
}

function ensureReportChartPlugins() {
  if (typeof Chart === "undefined") return false;
  const dl = typeof ChartDataLabels !== "undefined" ? ChartDataLabels : null;
  if (!dl) return false;
  if (!window._mbcReportDataLabelsRegistered) {
    Chart.register(dl);
    window._mbcReportDataLabelsRegistered = true;
  }
  return true;
}

function refreshReportView() {
  if (activeSectionId !== "report") return;
  destroyReportCharts();
  tableContainer.innerHTML = renderReportsPanel();
  requestAnimationFrame(() => {
    attachReportCharts();
    attachReportFilterHandlers();
    attachReportPhaseGroupLayoutHandlers();
    attachReportChartTileDragHandlers();
    attachReportExportAndShareHandlers();
    initLucideIcons();
  });
}

function attachReportPhaseGroupLayoutHandlers() {
  if (isSharedReportView()) return;
  const root = tableContainer;
  if (!root) return;
  root.querySelectorAll("[data-report-phase-layout]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const mode = String(btn.getAttribute("data-report-phase-layout") || "");
      if (mode !== "row" && mode !== "separate") return;
      saveReportPhaseGroupLayout(mode);
      refreshReportView();
    });
  });
}

function attachReportFilterHandlers() {
  if (isSharedReportView()) return;
  const root = tableContainer;
  if (!root) return;
  const toggleBtn = root.querySelector("#reportFiltersToggle");
  const panel = root.querySelector("#reportFiltersPanel");
  if (toggleBtn && panel) {
    toggleBtn.addEventListener("click", () => {
      reportFiltersPanelOpen = !reportFiltersPanelOpen;
      panel.classList.toggle("report-filters-panel--hidden", !reportFiltersPanelOpen);
      toggleBtn.setAttribute("aria-expanded", reportFiltersPanelOpen ? "true" : "false");
      toggleBtn.classList.toggle("is-active", reportFiltersPanelOpen);
    });
  }
  root.querySelector("#reportFilterObjectBtn")?.addEventListener("click", () => {
    openReportFilterPickerModal("Объект", "Все объекты", getReportDistinctObjects(), reportFilterObject, (v) => {
      reportFilterObject = String(v || "").trim();
      reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
      refreshReportView();
    });
  });
  root.querySelector("#reportFilterEmployeeBtn")?.addEventListener("click", () => {
    openReportFilterPickerModal("Сотрудник", "Все сотрудники", getReportDistinctEmployees(), reportFilterEmployee, (v) => {
      reportFilterEmployee = String(v || "").trim();
      reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
      refreshReportView();
    });
  });
  const applyBtn = root.querySelector("#reportFilterApply");
  const resetBtn = root.querySelector("#reportFilterReset");
  if (applyBtn) {
    applyBtn.addEventListener("click", () => {
      const boxes = root.querySelectorAll(".report-filter-status:checked");
      const picked = new Set(Array.from(boxes).map((el) => String(el.dataset.status || "")));
      if (picked.size === 0) {
        reportStatusFilter = new Set();
      } else if (picked.size === REPORT_FILTER_ALL_STATUSES.length) {
        reportStatusFilter = null;
      } else {
        reportStatusFilter = picked;
      }
      reportDateFrom = root.querySelector("#reportDateFrom")?.value?.trim() || "";
      reportDateTo = root.querySelector("#reportDateTo")?.value?.trim() || "";
      reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
      refreshReportView();
    });
  }
  if (resetBtn) {
    resetBtn.addEventListener("click", () => {
      reportStatusFilter = null;
      reportDateFrom = "";
      reportDateTo = "";
      reportFilterObject = "";
      reportFilterEmployee = "";
      reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
      refreshReportView();
    });
  }
  root.querySelector("#reportWeekShowMoreBtn")?.addEventListener("click", () => {
    reportWeekRowsVisible += REPORT_WEEK_ROWS_STEP;
    refreshReportView();
  });
}

function isSharedReportView() {
  return sharedReportMode === true;
}

function loadReportShares() {
  try {
    const raw = localStorage.getItem(REPORT_SHARE_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (_) {
    return {};
  }
}

function saveReportShares(map) {
  try {
    localStorage.setItem(REPORT_SHARE_STORAGE_KEY, JSON.stringify(map));
    scheduleServerSync();
  } catch (_) {
    /* noop */
  }
}

function randomReportShareId() {
  const buf = new Uint8Array(16);
  crypto.getRandomValues(buf);
  return Array.from(buf, (b) => b.toString(16).padStart(2, "0")).join("");
}

function randomSharePin4() {
  return String(1000 + Math.floor(Math.random() * 9000));
}

function toDatetimeLocalValue(date) {
  const d = date instanceof Date ? date : new Date(date);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatShareExpiryLabel(ts) {
  try {
    return new Date(ts).toLocaleString("ru-RU", { dateStyle: "medium", timeStyle: "short" });
  } catch (_) {
    return "";
  }
}

function buildReportShareSnapshot() {
  const rows = getReportFilteredRows();
  return {
    rows: JSON.parse(JSON.stringify(rows))
  };
}

function printReportDashboardPdf() {
  const source = tableContainer?.querySelector(".report-dashboard-card");
  if (!source) return;
  const clone = source.cloneNode(true);

  clone.querySelector(".shared-report-banner")?.remove();
  clone.querySelector(".table-header-right")?.remove();
  clone.querySelector(".report-filters-outer")?.remove();
  clone.querySelector(".report-phase-layout-bar")?.remove();
  clone.querySelectorAll(".report-chart-drag-handle").forEach((el) => el.remove());

  const mainTitle = clone.querySelector(".report-dashboard-header h3");
  if (mainTitle) {
    mainTitle.textContent = "Отчёт: аналитика по задачам";
  }
  clone.querySelectorAll(".report-week-tasks-title").forEach((el) => {
    el.textContent = "Задачи за последние 7 дней";
  });

  const canvases = source.querySelectorAll("canvas");
  const cloneCanvases = clone.querySelectorAll("canvas");
  canvases.forEach((c, i) => {
    const slot = cloneCanvases[i];
    if (!slot) return;
    try {
      const url = c.toDataURL("image/png");
      const img = document.createElement("img");
      img.src = url;
      img.alt = "";
      img.className = "report-print-chart-img";
      img.style.display = "block";
      img.style.margin = "6px auto 0";
      img.style.maxWidth = "100%";
      img.style.width = "auto";
      img.style.height = "auto";
      img.style.maxHeight = "220px";
      slot.replaceWith(img);
    } catch (_) {
      /* tainted canvas etc. */
    }
  });

  const stylesHref = new URL("styles.css", location.href).href;
  const printCss = `
    @page { size: A4 portrait; margin: 12mm; }
    html, body { margin: 0; background: #fff !important; color: #1f2a37; }
    .report-print-root {
      font-family: "Segoe UI", system-ui, Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.35;
      padding: 0;
      max-width: 100%;
    }
    .report-print-root * { box-sizing: border-box; }
    .report-print-root .table-card,
    .report-print-root .report-dashboard-card {
      border: none !important;
      box-shadow: none !important;
      padding: 0 !important;
    }
    .report-print-root .report-dashboard-header {
      margin-bottom: 10px;
    }
    .report-print-root .report-dashboard-header h3 {
      font-size: 14pt;
      font-weight: 700;
      margin: 0;
    }
    .report-print-root svg {
      width: 16px !important;
      height: 16px !important;
      max-width: 20px !important;
      max-height: 20px !important;
    }
    .report-print-root .report-grid,
    .report-print-root .report-phase-group-row {
      display: block !important;
      width: 100% !important;
    }
    .report-print-root .report-tile,
    .report-print-root .report-tile-wide {
      width: 100% !important;
      max-width: 100% !important;
      min-height: 0 !important;
      margin: 0 0 12px 0 !important;
      break-inside: avoid;
      page-break-inside: avoid;
    }
    .report-print-root .report-tile h4 {
      font-size: 10.5pt;
      padding-right: 0 !important;
    }
    .report-print-root .report-canvas-wrap,
    .report-print-root .report-canvas-scroll {
      min-height: 0 !important;
      overflow: visible !important;
    }
    .report-print-root .report-print-chart-img {
      max-height: 220px !important;
      max-width: 100% !important;
      width: auto !important;
      height: auto !important;
    }
    .report-print-root .report-week-tasks-scroll,
    .report-print-root .report-matrix-scroll {
      max-height: none !important;
      overflow: visible !important;
    }
    .report-print-root .report-week-table,
    .report-print-root .report-matrix-table {
      font-size: 9.5pt;
      width: 100%;
      border-collapse: collapse;
    }
    .report-print-root .report-week-table th,
    .report-print-root .report-week-table td,
    .report-print-root .report-matrix-table th,
    .report-print-root .report-matrix-table td {
      border: 1px solid #cbd5e1;
      padding: 5px 6px !important;
      vertical-align: top;
    }
    .report-print-root .report-matrix-title,
    .report-print-root .report-week-tasks-title {
      font-size: 11pt;
      margin: 12px 0 6px 0;
    }
    @media print {
      .report-print-root .hint.report-empty-hint { font-size: 10pt; }
    }
  `;

  const win = window.open("", "_blank");
  if (!win) return;
  win.document.write(`
    <!DOCTYPE html>
    <html lang="ru">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <title>Отчёт</title>
      <link rel="stylesheet" href="${stylesHref}" />
      <style>${printCss}</style>
    </head>
    <body class="report-print-root">${clone.outerHTML}</body>
    </html>
  `);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 250);
}

function exportCurrentReportToXls() {
  const taskSection = getSectionById("tasks");
  if (!taskSection) return;
  const rows = getReportFilteredRows();
  const filteredEntries = rows.map((row, rowIndex) => ({ row, rowIndex }));
  exportRowsToCsv(taskSection, filteredEntries, "Отчёт.xls");
}

function attachReportExportAndShareHandlers() {
  const root = tableContainer;
  if (!root) return;

  root.querySelector("#reportRefreshBtn")?.addEventListener("click", () => {
    reportWeekRowsVisible = REPORT_WEEK_ROWS_STEP;
    refreshCurrentViewData();
  });
  root.querySelector("#reportExportPdfBtn")?.addEventListener("click", () => {
    printReportDashboardPdf();
  });
  root.querySelector("#reportExportXlsBtn")?.addEventListener("click", () => {
    exportCurrentReportToXls();
  });
  root.querySelector("#reportShareBtn")?.addEventListener("click", () => {
    openReportShareModal();
  });
  root.querySelector("#sharedReportExitBtn")?.addEventListener("click", () => {
    exitSharedReportView();
  });
}

function openReportShareModal() {
  const snapshot = buildReportShareSnapshot();
  const overlay = document.createElement("div");
  overlay.className = "report-share-modal-overlay";
  const defaultExpiry = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  overlay.innerHTML = `
    <div class="report-share-modal card">
      <h3>Поделиться отчётом</h3>
      <p class="hint">Выберите дату и время окончания действия ссылки. Код из 4 цифр покажем после создания — передайте его вместе со ссылкой.</p>
      <label class="report-share-field">
        <span>Действует до</span>
        <input type="datetime-local" id="reportShareExpiresInput" required />
      </label>
      <div class="report-share-modal-actions">
        <button type="button" class="secondary" id="reportShareCancelBtn">Отмена</button>
        <button type="button" id="reportShareCreateBtn">Создать ссылку</button>
      </div>
      <div id="reportShareResult" class="hidden"></div>
    </div>
  `;
  document.body.appendChild(overlay);
  const expiresInput = overlay.querySelector("#reportShareExpiresInput");
  if (expiresInput instanceof HTMLInputElement) {
    expiresInput.min = toDatetimeLocalValue(new Date());
    expiresInput.value = toDatetimeLocalValue(defaultExpiry);
  }

  const close = () => overlay.remove();

  overlay.querySelector("#reportShareCancelBtn")?.addEventListener("click", close);
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) close();
  });

  overlay.querySelector("#reportShareCreateBtn")?.addEventListener("click", async () => {
    const raw = expiresInput instanceof HTMLInputElement ? expiresInput.value : "";
    const expiresAt = new Date(raw).getTime();
    if (!raw || Number.isNaN(expiresAt)) {
      alert("Укажите корректную дату и время окончания.");
      return;
    }
    if (expiresAt <= Date.now()) {
      alert("Время окончания должно быть позже текущего момента.");
      return;
    }
    const pin = randomSharePin4();
    let id;
    let linkStr;
    if (isHostedRuntime() && getAuthToken()) {
      try {
        const r = await fetch("/api/share", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${getAuthToken()}`
          },
          body: JSON.stringify({ pin, expiresAt, rows: snapshot.rows })
        });
        if (r.status === 401) {
          setAuthToken("");
          throw new Error("auth");
        }
        if (!r.ok) throw new Error("api");
        const j = await r.json();
        id = j.id;
        const map = loadReportShares();
        map[id] = { pin, expiresAt, rows: snapshot.rows };
        saveReportShares(map);
        const url = new URL(location.href);
        url.searchParams.set("share", id);
        linkStr = url.toString();
      } catch (_) {
        alert("Не удалось создать ссылку на сервере. Проверьте вход или сеть.");
        return;
      }
    } else {
      id = randomReportShareId();
      const map = loadReportShares();
      map[id] = {
        pin,
        expiresAt,
        rows: snapshot.rows
      };
      saveReportShares(map);
      const url = new URL(location.href);
      url.searchParams.set("share", id);
      linkStr = url.toString();
    }
    const resultEl = overlay.querySelector("#reportShareResult");
    if (resultEl) {
      resultEl.classList.remove("hidden");
      resultEl.innerHTML = `
        <div class="report-share-result-block">
          <label>Ссылка</label>
          <div class="report-share-copy-row">
            <input type="text" readonly class="report-share-link-input" value="${escapeHtmlAttr(linkStr)}" id="reportShareLinkField" />
            <button type="button" class="icon-action-btn" id="reportShareCopyLinkBtn" title="Копировать ссылку"><i data-lucide="copy" class="lucide-icon" aria-hidden="true"></i></button>
          </div>
          <label>Код доступа (4 цифры)</label>
          <div class="report-share-pin-row">
            <span class="report-share-pin" id="reportSharePinDisplay">${escapeHtmlText(pin)}</span>
            <button type="button" class="icon-action-btn" id="reportShareCopyPinBtn" title="Копировать код"><i data-lucide="copy" class="lucide-icon" aria-hidden="true"></i></button>
          </div>
        </div>
      `;
      initLucideIcons();
      resultEl.querySelector("#reportShareCopyLinkBtn")?.addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(linkStr);
        } catch (_) {
          /* noop */
        }
      });
      resultEl.querySelector("#reportShareCopyPinBtn")?.addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(pin);
        } catch (_) {
          /* noop */
        }
      });
    }
  });
}

function mountReportShareGate(shareId) {
  (async () => {
    let serverExpires = null;
    if (isHostedRuntime()) {
      try {
        const mr = await fetch(`/api/share/${encodeURIComponent(shareId)}/meta`);
        if (mr.ok) {
          const j = await mr.json();
          serverExpires = Number(j.expiresAt);
        }
      } catch (_) {
        /* noop */
      }
    }
    const mapPre = loadReportShares();
    const recPre = mapPre[shareId];
    const hasServer = serverExpires != null && !Number.isNaN(serverExpires);
    const missingOnly = !recPre && !hasServer;
    const expiredOnly =
      (recPre && Date.now() > recPre.expiresAt) || (hasServer && Date.now() > serverExpires);

    const overlay = document.createElement("div");
    overlay.className = "report-share-gate-overlay";
    overlay.innerHTML = `
    <div class="report-share-gate card">
      <h3>Доступ к отчёту</h3>
      ${missingOnly ? '<p class="error">Ссылка недействительна.</p>' : ""}
      ${expiredOnly ? '<p class="error">Срок действия ссылки истёк.</p>' : ""}
      <p class="hint">Введите 4-значный код, который вам передали.</p>
      <label for="reportSharePinInput">Код</label>
      <input type="text" inputmode="numeric" pattern="[0-9]*" maxlength="4" autocomplete="one-time-code" id="reportSharePinInput" class="report-share-pin-input" placeholder="0000" />
      <p id="reportShareGateError" class="error hidden">Неверный код или ссылка недействительна.</p>
      <div class="report-share-gate-actions">
        <button type="button" class="secondary" id="reportShareGateCancelBtn">На страницу входа</button>
        <button type="button" id="reportShareGateSubmitBtn">Открыть</button>
      </div>
    </div>
  `;
    document.body.appendChild(overlay);
    hideBootLoaderAfterRender();
    const pinInput = overlay.querySelector("#reportSharePinInput");
    const errEl = overlay.querySelector("#reportShareGateError");
    const submitBtn = overlay.querySelector("#reportShareGateSubmitBtn");
    if (missingOnly || expiredOnly) {
      pinInput?.setAttribute("disabled", "disabled");
      if (submitBtn) submitBtn.disabled = true;
    }

    pinInput?.addEventListener("input", (e) => {
      const t = e.target;
      if (t instanceof HTMLInputElement) {
        t.value = String(t.value).replace(/\D/g, "").slice(0, 4);
      }
    });

    const tryEnter = async () => {
      const pin = String(pinInput?.value || "").replace(/\D/g, "").slice(0, 4);
      errEl?.classList.add("hidden");

      if (recPre) {
        if (Date.now() > recPre.expiresAt) {
          errEl?.classList.remove("hidden");
          if (errEl) errEl.textContent = "Срок действия ссылки истёк.";
          return;
        }
        if (pin.length !== 4) {
          errEl?.classList.remove("hidden");
          if (errEl) errEl.textContent = "Введите 4 цифры кода.";
          return;
        }
        if (pin !== String(recPre.pin || "")) {
          errEl?.classList.remove("hidden");
          if (errEl) errEl.textContent = "Неверный код.";
          return;
        }
        overlay.remove();
        sharedReportExpiresAt = recPre.expiresAt;
        enterSharedReportView(recPre);
        return;
      }

      if (hasServer) {
        if (Date.now() > serverExpires) {
          errEl?.classList.remove("hidden");
          if (errEl) errEl.textContent = "Срок действия ссылки истёк.";
          return;
        }
        if (pin.length !== 4) {
          errEl?.classList.remove("hidden");
          if (errEl) errEl.textContent = "Введите 4 цифры кода.";
          return;
        }
        try {
          const r = await fetch(`/api/share/${encodeURIComponent(shareId)}/verify`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ pin })
          });
          if (r.status === 404) {
            errEl?.classList.remove("hidden");
            if (errEl) errEl.textContent = "Ссылка недействительна.";
            return;
          }
          if (r.status === 410) {
            errEl?.classList.remove("hidden");
            if (errEl) errEl.textContent = "Срок действия ссылки истёк.";
            return;
          }
          if (r.status === 403) {
            errEl?.classList.remove("hidden");
            if (errEl) errEl.textContent = "Неверный код.";
            return;
          }
          if (!r.ok) {
            errEl?.classList.remove("hidden");
            if (errEl) errEl.textContent = "Не удалось открыть ссылку.";
            return;
          }
          const j = await r.json();
          overlay.remove();
          sharedReportExpiresAt = j.expiresAt;
          enterSharedReportView({ rows: j.rows });
        } catch (_) {
          errEl?.classList.remove("hidden");
          if (errEl) errEl.textContent = "Ошибка сети.";
        }
        return;
      }

      errEl?.classList.remove("hidden");
      if (errEl) errEl.textContent = "Ссылка недействительна или устарела.";
    };

    overlay.querySelector("#reportShareGateSubmitBtn")?.addEventListener("click", () => {
      tryEnter();
    });
    pinInput?.addEventListener("keydown", (e) => {
      if (e.key === "Enter") tryEnter();
    });
    overlay.querySelector("#reportShareGateCancelBtn")?.addEventListener("click", () => {
      overlay.remove();
      const url = new URL(location.href);
      url.searchParams.delete("share");
      history.replaceState({}, "", url.pathname + url.search + url.hash);
      showLogin();
    });
    setTimeout(() => pinInput?.focus(), 0);
  })();
}

function enterSharedReportView(rec) {
  sharedReportMode = true;
  reportShareRowsOverride = rec.rows || [];
  document.body.classList.remove("login-mode");
  document.body.classList.add("shared-report-mode");
  loginSection.classList.add("hidden");
  appSection.classList.remove("hidden");
  activeSectionId = "report";
  renderSidebarMenu();
  renderTable();
  stopSessionIdleWatcher();
  hideBootLoaderAfterRender();
}

function exitSharedReportView() {
  sharedReportMode = false;
  reportShareRowsOverride = null;
  sharedReportExpiresAt = 0;
  document.body.classList.remove("shared-report-mode");
  const url = new URL(location.href);
  url.searchParams.delete("share");
  history.replaceState({}, "", url.pathname + url.search + url.hash);
  showLogin();
}

function destroyReportCharts() {
  reportChartInstances.forEach((ch) => {
    try {
      ch.destroy();
    } catch (_) {
      /* noop */
    }
  });
  reportChartInstances = [];
}

function buildTaskReportStats(rows) {
  const statusCounts = {};
  const priorityCounts = {};
  const phaseCounts = {};
  const objectCounts = {};
  const phaseSectionCounts = {};
  const phaseSubsectionCounts = {};
  const responsibleCounts = {};
  const monthCounts = {};
  const monthClosedCounts = {};
  const overdueByObject = {};
  const nameToDept = buildEmployeeNameToDepartmentMap();
  const departmentStatusMap = new Map();
  const responsibleStatusMap = new Map();
  const objectStatusMap = new Map();
  const phaseStatusMap = new Map();
  const phaseSectionStatusMap = new Map();
  const phaseSubsectionStatusMap = new Map();

  rows.forEach((row) => {
    const st = String(row[TASK_COLUMNS.status] || "").trim() || REPORT_NO_STATUS_LABEL;
    statusCounts[st] = (statusCounts[st] || 0) + 1;

    const pr = String(row[TASK_COLUMNS.priority] || "").trim() || "(не указан)";
    priorityCounts[pr] = (priorityCounts[pr] || 0) + 1;

    const ph = String(row[TASK_COLUMNS.phase] || "").trim() || "— нет —";
    phaseCounts[ph] = (phaseCounts[ph] || 0) + 1;

    const ob = String(row[TASK_COLUMNS.object] || "").trim() || "— нет —";
    objectCounts[ob] = (objectCounts[ob] || 0) + 1;

    const psec = String(row[TASK_COLUMNS.phaseSection] || "").trim() || "— нет —";
    phaseSectionCounts[psec] = (phaseSectionCounts[psec] || 0) + 1;

    const psub = String(row[TASK_COLUMNS.phaseSubsection] || "").trim() || "— нет —";
    phaseSubsectionCounts[psub] = (phaseSubsectionCounts[psub] || 0) + 1;

    const resp = String(row[TASK_COLUMNS.assignedResponsible] || "").trim() || "— не назначен —";
    responsibleCounts[resp] = (responsibleCounts[resp] || 0) + 1;

    const addD = String(row[TASK_COLUMNS.addedDate] || "").trim();
    const parts = parseRuDateStringToParts(addD);
    if (parts) {
      const key = `${parts.year}-${String(parts.month).padStart(2, "0")}`;
      monthCounts[key] = (monthCounts[key] || 0) + 1;
    }

    const closedD = String(row[TASK_COLUMNS.closedDate] || "").trim();
    const closedParts = parseRuDateStringToParts(closedD);
    if (closedParts) {
      const ckey = `${closedParts.year}-${String(closedParts.month).padStart(2, "0")}`;
      monthClosedCounts[ckey] = (monthClosedCounts[ckey] || 0) + 1;
    }

    if (isTaskOverdueForReport(row)) {
      const obO = String(row[TASK_COLUMNS.object] || "").trim() || "— нет —";
      overdueByObject[obO] = (overdueByObject[obO] || 0) + 1;
    }

    if (STATUS_OPTIONS.includes(st)) {
      const respRaw = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
      const respLabel = respRaw || "— не назначен —";
      if (!responsibleStatusMap.has(respLabel)) {
        const o = {};
        for (const st of STATUS_OPTIONS) o[st] = 0;
        responsibleStatusMap.set(respLabel, o);
      }
      responsibleStatusMap.get(respLabel)[st] += 1;

      let deptLabel;
      if (!respRaw) deptLabel = "— не назначен —";
      else {
        const key = normalizePersonName(respRaw);
        deptLabel = nameToDept.get(key) || "Не в справочнике";
      }
      if (!departmentStatusMap.has(deptLabel)) {
        const o = {};
        for (const st of STATUS_OPTIONS) o[st] = 0;
        departmentStatusMap.set(deptLabel, o);
      }
      departmentStatusMap.get(deptLabel)[st] += 1;

      if (!objectStatusMap.has(ob)) {
        const o = {};
        for (const st of STATUS_OPTIONS) o[st] = 0;
        objectStatusMap.set(ob, o);
      }
      objectStatusMap.get(ob)[st] += 1;

      if (!phaseStatusMap.has(ph)) {
        const o = {};
        for (const st of STATUS_OPTIONS) o[st] = 0;
        phaseStatusMap.set(ph, o);
      }
      phaseStatusMap.get(ph)[st] += 1;

      if (!phaseSectionStatusMap.has(psec)) {
        const o = {};
        for (const st of STATUS_OPTIONS) o[st] = 0;
        phaseSectionStatusMap.set(psec, o);
      }
      phaseSectionStatusMap.get(psec)[st] += 1;

      if (!phaseSubsectionStatusMap.has(psub)) {
        const o = {};
        for (const st of STATUS_OPTIONS) o[st] = 0;
        phaseSubsectionStatusMap.set(psub, o);
      }
      phaseSubsectionStatusMap.get(psub)[st] += 1;
    }
  });

  const topEntries = (obj, n) =>
    Object.entries(obj)
      .sort((a, b) => b[1] - a[1])
      .slice(0, n);

  const monthKeysSorted = Object.keys(monthCounts).sort();

  const buildStatusStacked = (statusMap, stackId, maxItems = null) => {
    const labelsSorted = Array.from(statusMap.keys()).sort((a, b) => {
      const sum = (key) => STATUS_OPTIONS.reduce((s, st) => s + (statusMap.get(key)[st] || 0), 0);
      const ta = sum(a);
      const tb = sum(b);
      if (tb !== ta) return tb - ta;
      return String(a).localeCompare(String(b), "ru");
    });
    const labels = maxItems ? labelsSorted.slice(0, maxItems) : labelsSorted;
    return labels.length > 0
      ? {
          labels,
          datasets: STATUS_OPTIONS.map((st) => ({
            label: st,
            data: labels.map((lab) => statusMap.get(lab)[st] || 0),
            backgroundColor: colorForStatusLabel(st),
            stack: stackId
          }))
        }
      : {
          labels: ["Нет данных"],
          datasets: STATUS_OPTIONS.map((st) => ({
            label: st,
            data: [0],
            backgroundColor: colorForStatusLabel(st),
            stack: stackId
          }))
        };
  };

  const departmentStacked = buildStatusStacked(departmentStatusMap, "dept");
  const responsibleStacked = buildStatusStacked(responsibleStatusMap, "resp", 10);
  const objectStacked = buildStatusStacked(objectStatusMap, "obj");
  const phaseStacked = buildStatusStacked(phaseStatusMap, "ph", 10);
  const phaseSectionStacked = buildStatusStacked(phaseSectionStatusMap, "phsec");
  const phaseSubsectionStacked = buildStatusStacked(phaseSubsectionStatusMap, "phsub");

  return {
    statusCounts,
    priorityCounts,
    phaseTop: topEntries(phaseCounts, 10),
    objectAll: sortedCountEntries(objectCounts),
    phaseSectionAll: sortedCountEntries(phaseSectionCounts),
    phaseSubsectionAll: sortedCountEntries(phaseSubsectionCounts),
    responsibleTop: topEntries(responsibleCounts, 10),
    monthCounts,
    monthClosedCounts,
    monthKeysSorted,
    overdueTop: topEntries(overdueByObject, 12),
    departmentStacked,
    responsibleStacked,
    objectStacked,
    phaseStacked,
    phaseSectionStacked,
    phaseSubsectionStacked,
    total: rows.length
  };
}

function renderResponsibleStatusTable(rsRows) {
  if (!rsRows.length) return "";
  const cols = STATUS_OPTIONS;
  const colgroup = `<colgroup>
    <col class="report-matrix-col-resp" />
    ${cols.map(() => `<col class="report-matrix-col-st" />`).join("")}
    <col class="report-matrix-col-total" />
  </colgroup>`;
  const head = `<tr><th>Ответственный</th>${cols.map((c) => `<th>${escapeHtmlText(c)}</th>`).join("")}<th>Всего</th></tr>`;
  const body = rsRows
    .map((row) => {
      const cells = cols.map((c) => `<td class="report-matrix-num">${row.counts[c] || 0}</td>`).join("");
      return `<tr><td>${escapeHtmlText(row.name)}</td>${cells}<td class="report-matrix-num"><strong>${row.total}</strong></td></tr>`;
    })
    .join("");
  return `
    <div class="report-matrix-wrap report-matrix-wrap--after-charts">
      <h4 class="report-matrix-title">Задачи по исполнителям и статусам</h4>
      <div class="report-matrix-scroll">
        <table class="report-matrix-table">
          ${colgroup}
          <thead>${head}</thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </div>
  `;
}

function getTaskRowsAddedInLast7Days() {
  const rows = getAllTaskRows();
  const end = new Date();
  const start = new Date(end);
  start.setDate(start.getDate() - 6);
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);
  const sp = { year: start.getFullYear(), month: start.getMonth() + 1, day: start.getDate() };
  const ep = { year: end.getFullYear(), month: end.getMonth() + 1, day: end.getDate() };
  return rows
    .filter((row) => {
      const rp = parseRuDateStringToParts(String(row[TASK_COLUMNS.addedDate] || "").trim());
      if (!rp) return false;
      return compareDateParts(rp, sp) >= 0 && compareDateParts(rp, ep) <= 0;
    })
    .sort((a, b) => {
      const pa = parseRuDateStringToParts(String(a[TASK_COLUMNS.addedDate] || ""));
      const pb = parseRuDateStringToParts(String(b[TASK_COLUMNS.addedDate] || ""));
      if (!pa || !pb) return 0;
      return compareDateParts(pb, pa);
    });
}

function formatReportDeadlinesCell(row) {
  const plan = String(row[TASK_COLUMNS.plan] || "").trim();
  const due = String(row[TASK_COLUMNS.dueDate] || "").trim();
  const dueDisp = due ? formatStoredDateForDisplay(due) : "";
  const parts = [];
  if (plan) parts.push(`план: ${plan}`);
  if (dueDisp) parts.push(`срок: ${dueDisp}`);
  return parts.length ? parts.join(" · ") : "—";
}

function formatReportResponsiblesCell(row) {
  const a = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
  const r = String(row[TASK_COLUMNS.responsible] || "").trim();
  const bits = [];
  if (a) bits.push(a);
  if (r) bits.push(`контр.: ${r}`);
  return bits.length ? bits.join(" · ") : "—";
}

function renderReportWeekTasksTable() {
  const weekRows = getTaskRowsAddedInLast7Days();
  const visibleLimit = Math.min(Math.max(REPORT_WEEK_ROWS_STEP, reportWeekRowsVisible || REPORT_WEEK_ROWS_STEP), weekRows.length || REPORT_WEEK_ROWS_STEP);
  const shownRows = weekRows.slice(0, visibleLimit);
  if (!weekRows.length) {
    return `
    <div class="report-week-tasks-wrap">
      <h4 class="report-week-tasks-title">${withLucideIcon("calendar-days", "Задачи за последние 7 дней")}</h4>
      <p class="hint report-week-tasks-empty">Нет задач с датой добавления за последние 7 дней.</p>
    </div>`;
  }
  const head = `<tr>
    <th>ID</th>
    <th>Объект</th>
    <th>Дата постановки задачи</th>
    <th>Сроки</th>
    <th>Название задачи</th>
    <th>Фаза</th>
    <th>Раздел</th>
    <th>Подраздел</th>
    <th>Ответственные</th>
  </tr>`;
  const body = shownRows
    .map((row) => {
      const id = escapeHtmlText(String(row[TASK_COLUMNS.number] ?? ""));
      const ob = escapeHtmlText(String(row[TASK_COLUMNS.object] ?? ""));
      const addD = escapeHtmlText(formatStoredDateForDisplay(row[TASK_COLUMNS.addedDate]));
      const dl = escapeHtmlText(formatReportDeadlinesCell(row));
      const task = escapeHtmlText(String(row[TASK_COLUMNS.task] ?? ""));
      const ph = escapeHtmlText(String(row[TASK_COLUMNS.phase] ?? ""));
      const sec = escapeHtmlText(String(row[TASK_COLUMNS.phaseSection] ?? ""));
      const sub = escapeHtmlText(String(row[TASK_COLUMNS.phaseSubsection] ?? ""));
      const resp = escapeHtmlText(formatReportResponsiblesCell(row));
      return `<tr>
        <td class="report-week-num">${id}</td>
        <td>${ob}</td>
        <td class="report-week-nowrap">${addD}</td>
        <td>${dl}</td>
        <td>${task}</td>
        <td>${ph}</td>
        <td>${sec}</td>
        <td>${sub}</td>
        <td>${resp}</td>
      </tr>`;
    })
    .join("");
  const hasMore = weekRows.length > shownRows.length;
  const moreBtn = hasMore
    ? `<div class="report-week-more-wrap"><button type="button" class="secondary report-week-more-btn" id="reportWeekShowMoreBtn">Показать ещё</button></div>`
    : "";
  return `
    <div class="report-week-tasks-wrap">
      <h4 class="report-week-tasks-title">${withLucideIcon("calendar-days", "Задачи за последние 7 дней")}</h4>
      <div class="report-week-tasks-scroll">
        <table class="report-week-table">
          <thead>${head}</thead>
          <tbody>${body}</tbody>
        </table>
      </div>
      ${moreBtn}
    </div>`;
}

function renderReportsPanel() {
  const rows = getReportFilteredRows();
  const stats = buildTaskReportStats(rows);
  const rsRows = buildResponsibleStatusRows(rows);
  const shareBanner = sharedReportMode
    ? `<div class="shared-report-banner" role="status">
        <span>Временный просмотр по ссылке. Доступ до <strong>${escapeHtmlText(formatShareExpiryLabel(sharedReportExpiresAt))}</strong>.</span>
        <button type="button" class="secondary" id="sharedReportExitBtn">Закрыть</button>
      </div>`
    : "";
  const exportShareBtns = `
      <button type="button" class="icon-action-btn" id="reportRefreshBtn" title="Обновить" aria-label="Обновить">
        <i data-lucide="refresh-cw" class="lucide-icon" aria-hidden="true"></i>
      </button>
      <button type="button" class="icon-action-btn" id="reportExportPdfBtn" title="Скачать PDF" aria-label="Скачать PDF">
        <i data-lucide="file-text" class="lucide-icon" aria-hidden="true"></i>
      </button>
      <button type="button" class="icon-action-btn" id="reportExportXlsBtn" title="Скачать XLS" aria-label="Скачать XLS">
        <i data-lucide="sheet" class="lucide-icon" aria-hidden="true"></i>
      </button>
      ${sharedReportMode ? "" : `<button type="button" class="icon-action-btn" id="reportShareBtn" title="Поделиться" aria-label="Поделиться">
        <i data-lucide="share-2" class="lucide-icon" aria-hidden="true"></i>
      </button>`}
    `;
  const statusChecks = REPORT_FILTER_ALL_STATUSES.map((st) => {
    const on = reportStatusFilter === null || reportStatusFilter.has(st);
    const bg = statusBgForFilter(st);
    const ac = statusAccentForFilter(st);
    return `<label class="report-filter-check report-filter-status-pill report-filter-status-pill--compact" style="--st-bg:${bg};--st-border:${ac}"><input type="checkbox" class="report-filter-status" data-status="${escapeHtmlAttr(st)}" ${on ? "checked" : ""} /><span>${escapeHtmlText(st)}</span></label>`;
  }).join("");
  const objectBtnLabel = reportFilterObject ? escapeHtmlText(reportFilterObject) : "Объект";
  const employeeBtnLabel = reportFilterEmployee ? escapeHtmlText(reportFilterEmployee) : "Сотрудник";
  const filtersPanelClass = reportFiltersPanelOpen ? "" : "report-filters-panel--hidden";
  const filterBtnActive = reportFiltersPanelOpen ? " is-active" : "";
  const phaseLayout = loadReportPhaseGroupLayout();
  const phaseLayoutRowActive = phaseLayout === "row" ? " report-phase-layout-btn--active" : "";
  const phaseLayoutSepActive = phaseLayout === "separate" ? " report-phase-layout-btn--active" : "";
  return `
    <section class="table-card report-dashboard-card">
      ${shareBanner}
      <div class="table-header report-dashboard-header">
        <h3>${withIcon("barChart", "Отчёт: аналитика по задачам")}</h3>
        <div class="table-header-right">
          ${exportShareBtns}
          <button type="button" class="icon-action-btn filter-toggle-btn${filterBtnActive}${sharedReportMode ? " hidden" : ""}" id="reportFiltersToggle" title="Показать/скрыть фильтры" aria-label="Показать или скрыть фильтры" aria-expanded="${reportFiltersPanelOpen ? "true" : "false"}">
            <i data-lucide="filter" class="lucide-icon" aria-hidden="true"></i>
          </button>
        </div>
      </div>
      <div class="report-filters-outer${sharedReportMode ? " hidden" : ""}">
        <div id="reportFiltersPanel" class="report-filters-panel ${filtersPanelClass}">
          <div class="report-filters-one-row">
            <div class="report-filters-statuses report-filters-statuses--compact">${statusChecks}</div>
            <button type="button" class="report-filter-picker report-filter-input--compact" id="reportFilterObjectBtn" title="Выбрать объект">
              <span class="report-filter-picker-text">${objectBtnLabel}</span>
            </button>
            <button type="button" class="report-filter-picker report-filter-select--compact" id="reportFilterEmployeeBtn" title="Выбрать сотрудника">
              <span class="report-filter-picker-text">${employeeBtnLabel}</span>
            </button>
            <input type="date" class="report-filter-date report-filter-date--compact" id="reportDateFrom" value="${escapeHtmlAttr(reportDateFrom)}" />
            <span class="report-date-sep report-date-sep--inline">по</span>
            <input type="date" class="report-filter-date report-filter-date--compact" id="reportDateTo" value="${escapeHtmlAttr(reportDateTo)}" />
            <button type="button" class="report-filter-icon-btn" id="reportFilterApply" title="Применить" aria-label="Применить фильтр">
              <i data-lucide="check" class="lucide-icon" aria-hidden="true"></i>
            </button>
            <button type="button" class="report-filter-icon-btn report-filter-icon-btn--reset" id="reportFilterReset" title="Сбросить" aria-label="Сбросить фильтр">
              <i data-lucide="rotate-ccw" class="lucide-icon" aria-hidden="true"></i>
            </button>
          </div>
        </div>
      </div>
      ${renderReportWeekTasksTable()}
      ${stats.total === 0 ? '<p class="hint report-empty-hint">Нет задач по текущему фильтру — измените условия или добавьте записи в «Задачи».</p>' : ""}
      <div class="report-phase-layout-bar${sharedReportMode ? " hidden" : ""}">
        <span class="report-phase-layout-label">Топ фаз · Разделы · Подразделы</span>
        <div class="report-phase-layout-btns" role="group" aria-label="Расположение похожих графиков">
          <button type="button" class="report-phase-layout-btn${phaseLayoutRowActive}" data-report-phase-layout="row" title="Три графика в одной строке">В одну строку</button>
          <button type="button" class="report-phase-layout-btn${phaseLayoutSepActive}" data-report-phase-layout="separate" title="Каждый график на отдельной строке">По отдельности</button>
        </div>
      </div>
      <div class="report-grid report-grid--phase-${escapeHtmlAttr(phaseLayout)}">
        ${renderReportChartsGridHtml()}
      </div>
      ${stats.total > 0 ? renderResponsibleStatusTable(rsRows) : ""}
    </section>
  `;
}

function attachReportCharts() {
  if (typeof Chart === "undefined") {
    return;
  }
  const hasDl = ensureReportChartPlugins();
  const rows = getReportFilteredRows();
  const s = buildTaskReportStats(rows);
  const common = {
    responsive: true,
    maintainAspectRatio: false
  };
  const colorRot = (i) => REPORT_CHART_COLORS[i % REPORT_CHART_COLORS.length];
  const dlBarV = hasDl
    ? {
        anchor: "end",
        align: "top",
        offset: 2,
        formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
        color: "#475569",
        font: { weight: "600", size: 11 },
        clamp: true,
        clip: false
      }
    : { display: false };
  const dlBarH = hasDl
    ? {
        anchor: "end",
        align: "end",
        offset: 4,
        formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
        color: "#475569",
        font: { weight: "600", size: 11 },
        clamp: true,
        clip: false
      }
    : { display: false };
  const makeDlBarHLong = (n) =>
    hasDl
      ? {
          anchor: "end",
          align: "end",
          offset: 4,
          formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
          color: "#475569",
          font: { weight: "600", size: n > 50 ? 9 : 11 },
          clamp: true,
          clip: false,
          display: (ctx) => {
            const c = n || 0;
            if (c <= 45) return true;
            const step = Math.ceil(c / 45);
            return ctx.dataIndex % step === 0;
          }
        }
      : { display: false };
  const dlLine = hasDl
    ? {
        align: "top",
        anchor: "end",
        offset: 4,
        formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
        color: "#475569",
        font: { weight: "600", size: 11 },
        clamp: true,
        clip: false
      }
    : { display: false };
  const dlLineDual = hasDl
    ? {
        align: "top",
        anchor: "end",
        offset: 4,
        formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
        color: (ctx) => (ctx.datasetIndex === 0 ? "#475569" : "#166534"),
        font: { weight: "600", size: 11 },
        clamp: true,
        clip: false
      }
    : { display: false };
  const dlDonut = hasDl
    ? {
        formatter: (value) => (Number(value) > 0 ? String(Math.round(Number(value))) : ""),
        color: "#334155",
        font: { weight: "700", size: 10 },
        textAlign: "center",
        anchor: "end",
        align: "end",
        offset: 4,
        clamp: true,
        clip: false
      }
    : { display: false };
  const dlDonutStatus = hasDl
    ? {
        ...dlDonut,
        color: (ctx) => {
          const lab = ctx.chart.data.labels[ctx.dataIndex];
          if (lab === "Нет данных") return "#94a3b8";
          return STATUS_CHART_COLORS_ACCENT[lab] || "#475569";
        }
      }
    : { display: false };

  const stSorted = sortStatusEntriesForChart(s.statusCounts).filter((x) => x[1] > 0);
  const stLabelsFinal = stSorted.length ? stSorted.map((x) => x[0]) : ["Нет данных"];
  const stDataFinal = stSorted.length ? stSorted.map((x) => x[1]) : [0];
  const stColors = stLabelsFinal.map((lab) => colorForStatusLabel(lab));
  const totalTasksInSystem = getAllTaskRows().length;
  const donutCenterTotalPlugin = {
    id: "reportDonutCenterTotal",
    afterDraw(chart) {
      const meta = chart.getDatasetMeta(0);
      const arc = meta?.data?.[0];
      if (!arc) return;
      let x = arc.x;
      let y = arc.y;
      if (typeof x !== "number" || typeof y !== "number") {
        const pos = typeof arc.tooltipPosition === "function" ? arc.tooltipPosition() : null;
        if (pos && typeof pos.x === "number" && typeof pos.y === "number") {
          x = pos.x;
          y = pos.y;
        } else {
          const { left, top, width, height } = chart.chartArea;
          x = left + width * 0.36;
          y = top + height / 2;
        }
      }
      const c = chart.ctx;
      c.save();
      c.textAlign = "center";
      c.textBaseline = "middle";
      c.font = "600 18px Inter, system-ui, 'Segoe UI', sans-serif";
      c.fillStyle = "#1e293b";
      c.fillText(String(totalTasksInSystem), x, y - 8);
      c.font = "11px Inter, system-ui, 'Segoe UI', sans-serif";
      c.fillStyle = "#64748b";
      c.fillText("задач в системе", x, y + 10);
      c.restore();
    }
  };

  const ctx1 = document.getElementById("reportChartStatus");
  if (ctx1) {
    reportChartInstances.push(
      new Chart(ctx1, {
        type: "doughnut",
        plugins: [donutCenterTotalPlugin],
        data: {
          labels: stLabelsFinal,
          datasets: [
            {
              data: stDataFinal,
              backgroundColor: stColors,
              borderColor: "#f8fafc",
              borderWidth: 2
            }
          ]
        },
        options: {
          ...common,
          cutout: "52%",
          layout: {
            padding: { left: 14, right: 40, top: 16, bottom: 16 }
          },
          plugins: {
            legend: {
              position: "left",
              align: "center",
              labels: {
                boxWidth: 10,
                boxHeight: 10,
                padding: 10,
                usePointStyle: true,
                font: { size: 11 },
                generateLabels: (chart) => {
                  const data = chart.data;
                  const ds = data.datasets[0];
                  if (!data.labels?.length || !ds) return [];
                  return data.labels.map((label, i) => ({
                    text: `${label}: ${ds.data[i]}`,
                    fillStyle: Array.isArray(ds.backgroundColor) ? ds.backgroundColor[i] : ds.backgroundColor,
                    hidden: false,
                    index: i
                  }));
                }
              }
            },
            datalabels: dlDonutStatus
          }
        }
      })
    );
  }

  const prLabels = Object.keys(s.priorityCounts);
  const prData = prLabels.map((k) => s.priorityCounts[k]);
  const prLabelsFinal = prLabels.length ? prLabels : ["Нет данных"];
  const prDataFinal = prLabels.length ? prData : [0];
  const ctx2 = document.getElementById("reportChartPriority");
  if (ctx2) {
    reportChartInstances.push(
      new Chart(ctx2, {
        type: "bar",
        data: {
          labels: prLabelsFinal,
          datasets: [
            {
              label: "Задач",
              data: prDataFinal,
              backgroundColor: (context) => {
                const base = colorRot(context.dataIndex + 2);
                return reportBarGradientVertical(context.chart, base);
              }
            }
          ]
        },
        options: {
          ...common,
          ...REPORT_CHART_LABEL_SAFE_LAYOUT,
          scales: {
            y: { beginAtZero: true, ticks: { precision: 0 }, ...REPORT_VALUE_AXIS_GRACE }
          },
          plugins: {
            legend: { display: false },
            datalabels: hasDl ? { ...dlBarV, display: true } : { display: false }
          }
        }
      })
    );
  }

  const ym = getReportYearMonthChartSeries(s.monthCounts);
  const ymClosed = getReportYearMonthChartSeries(s.monthClosedCounts, ym.year);
  const mKeys = ym.labels;
  const mData = ym.data;
  const mDataClosed = ymClosed.data;
  const ctx3 = document.getElementById("reportChartMonths");
  if (ctx3) {
    reportChartInstances.push(
      new Chart(ctx3, {
        type: "line",
        data: {
          labels: mKeys,
          datasets: [
            {
              label: "Добавлено",
              data: mData,
              borderColor: "#8b9fd9",
              backgroundColor: (context) => {
                const chart = context.chart;
                const { ctx, chartArea } = chart;
                if (!chartArea) return "rgba(168, 180, 240, 0.25)";
                const g = ctx.createLinearGradient(0, chartArea.bottom, 0, chartArea.top);
                g.addColorStop(0, "rgba(139, 159, 217, 0)");
                g.addColorStop(0.55, "rgba(168, 180, 240, 0.35)");
                g.addColorStop(1, "rgba(139, 159, 217, 0.45)");
                return g;
              },
              fill: true,
              tension: 0.25
            },
            {
              label: "Закрыто",
              data: mDataClosed,
              borderColor: "#34d399",
              backgroundColor: (context) => {
                const chart = context.chart;
                const { ctx, chartArea } = chart;
                if (!chartArea) return "rgba(52, 211, 153, 0.2)";
                const g = ctx.createLinearGradient(0, chartArea.bottom, 0, chartArea.top);
                g.addColorStop(0, "rgba(52, 211, 153, 0)");
                g.addColorStop(0.55, "rgba(52, 211, 153, 0.28)");
                g.addColorStop(1, "rgba(16, 185, 129, 0.4)");
                return g;
              },
              fill: true,
              tension: 0.25
            }
          ]
        },
        options: {
          ...common,
          ...REPORT_CHART_LABEL_SAFE_LAYOUT,
          scales: {
            x: {
              ticks: {
                maxRotation: 45,
                minRotation: 0,
                autoSkip: false
              }
            },
            y: { beginAtZero: true, ticks: { precision: 0 }, ...REPORT_VALUE_AXIS_GRACE }
          },
          plugins: {
            legend: {
              display: true,
              position: "top",
              labels: { usePointStyle: true, padding: 14, font: { size: 11 } }
            },
            datalabels: dlLineDual
          }
        }
      })
    );
  }

  const ov = s.overdueTop.length ? s.overdueTop : [["—", 0]];
  const wrapOv = document.getElementById("reportChartOverdueWrap");
  setReportScrollableChartHeight(wrapOv, ov.length, { maxPx: 900, minPx: 240 });
  const ctxOverdue = document.getElementById("reportChartOverdue");
  if (ctxOverdue) {
    const nOv = ov.length;
    reportChartInstances.push(
      new Chart(ctxOverdue, {
        type: "bar",
        data: {
          labels: ov.map((x) => x[0]),
          datasets: [
            {
              label: "Просрочено",
              data: ov.map((x) => x[1]),
              backgroundColor: (context) => reportBarGradientHorizontal(context.chart, "#f97316")
            }
          ]
        },
        options: {
          ...common,
          ...REPORT_HBAR_OPTIONS_THIN,
          indexAxis: "y",
          scales: {
            x: { ...REPORT_HBAR_SCALE_X },
            y: {
              ticks: {
                autoSkip: false,
                font: { size: nOv > 40 ? 9 : 11 }
              }
            }
          },
          plugins: {
            legend: { display: false },
            datalabels: hasDl ? makeDlBarHLong(nOv) : { display: false }
          }
        }
      })
    );
  }

  const ph = s.phaseTop.length ? s.phaseTop : [["—", 0]];
  const psec = s.phaseSectionAll.length ? s.phaseSectionAll : [["—", 0]];
  const psub = s.phaseSubsectionAll.length ? s.phaseSubsectionAll : [["—", 0]];
  applyReportPhaseTrioWrapHeights(ph.length, psec.length, psub.length);

  const renderStatusStackedHBar = (ctx, stacked, n, fontRule) => {
    if (!ctx || !stacked?.datasets?.length) return;
    const labelsCount = n || stacked.labels?.length || 1;
    const dlStacked = hasDl
      ? {
          anchor: "center",
          align: "center",
          formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
          color: "#1e293b",
          font: { weight: "600", size: 10 },
          clamp: true,
          clip: false,
          display: (c) => Number(c.dataset.data[c.dataIndex]) > 0
        }
      : { display: false };
    reportChartInstances.push(
      new Chart(ctx, {
        type: "bar",
        data: {
          labels: stacked.labels,
          datasets: stacked.datasets
        },
        options: {
          ...common,
          ...REPORT_CHART_LABEL_SAFE_LAYOUT,
          indexAxis: "y",
          datasets: {
            bar: {
              maxBarThickness: 28,
              categoryPercentage: 0.82,
              barPercentage: 0.9
            }
          },
          scales: {
            x: {
              stacked: true,
              ...REPORT_HBAR_SCALE_X
            },
            y: {
              stacked: true,
              ticks: {
                autoSkip: false,
                font: { size: typeof fontRule === "function" ? fontRule(labelsCount) : 11 }
              }
            }
          },
          plugins: {
            legend: {
              display: true,
              position: "top",
              labels: { usePointStyle: true, padding: 12, font: { size: 11 } }
            },
            datalabels: dlStacked
          }
        }
      })
    );
  };

  const makeDonutSeries = (entries, colorOffset = 0) => {
    const positive = entries.filter((x) => Number(x[1]) > 0);
    const labels = positive.length ? positive.map((x) => x[0]) : ["Нет данных"];
    const data = positive.length ? positive.map((x) => x[1]) : [0];
    const colors = labels.map((_, i) => colorRot(i + colorOffset));
    return { labels, data, colors };
  };

  const phaseDonutSeries = makeDonutSeries(s.phaseTop, 0);
  const ctxPhaseDonut = document.getElementById("reportChartPhaseDonut");
  if (ctxPhaseDonut) {
    reportChartInstances.push(
      new Chart(ctxPhaseDonut, {
        type: "doughnut",
        data: {
          labels: phaseDonutSeries.labels,
          datasets: [
            {
              data: phaseDonutSeries.data,
              backgroundColor: phaseDonutSeries.colors,
              borderColor: "#f8fafc",
              borderWidth: 2
            }
          ]
        },
        options: {
          ...common,
          cutout: "52%",
          plugins: {
            legend: {
              position: "bottom",
              labels: { usePointStyle: true, padding: 10, font: { size: 11 } }
            },
            datalabels: dlDonut
          }
        }
      })
    );
  }

  const phaseSectionDonutSeries = makeDonutSeries(s.phaseSectionAll, 1);
  const ctxPhaseSectionDonut = document.getElementById("reportChartPhaseSectionDonut");
  if (ctxPhaseSectionDonut) {
    reportChartInstances.push(
      new Chart(ctxPhaseSectionDonut, {
        type: "doughnut",
        data: {
          labels: phaseSectionDonutSeries.labels,
          datasets: [
            {
              data: phaseSectionDonutSeries.data,
              backgroundColor: phaseSectionDonutSeries.colors,
              borderColor: "#f8fafc",
              borderWidth: 2
            }
          ]
        },
        options: {
          ...common,
          cutout: "52%",
          plugins: {
            legend: {
              position: "bottom",
              labels: { usePointStyle: true, padding: 10, font: { size: 11 } }
            },
            datalabels: dlDonut
          }
        }
      })
    );
  }

  const phaseSubsectionDonutSeries = makeDonutSeries(s.phaseSubsectionAll, 2);
  const ctxPhaseSubsectionDonut = document.getElementById("reportChartPhaseSubsectionDonut");
  if (ctxPhaseSubsectionDonut) {
    reportChartInstances.push(
      new Chart(ctxPhaseSubsectionDonut, {
        type: "doughnut",
        data: {
          labels: phaseSubsectionDonutSeries.labels,
          datasets: [
            {
              data: phaseSubsectionDonutSeries.data,
              backgroundColor: phaseSubsectionDonutSeries.colors,
              borderColor: "#f8fafc",
              borderWidth: 2
            }
          ]
        },
        options: {
          ...common,
          cutout: "52%",
          plugins: {
            legend: {
              position: "bottom",
              labels: { usePointStyle: true, padding: 10, font: { size: 11 } }
            },
            datalabels: dlDonut
          }
        }
      })
    );
  }

  const phaseStack = s.phaseStacked;
  const ctx4 = document.getElementById("reportChartPhase");
  renderStatusStackedHBar(ctx4, phaseStack, phaseStack?.labels?.length || 1, (n) => (n > 20 ? 10 : 11));

  const phaseSectionStack = s.phaseSectionStacked;
  const ctxSec = document.getElementById("reportChartPhaseSection");
  const nSec = phaseSectionStack?.labels?.length || 1;
  renderStatusStackedHBar(ctxSec, phaseSectionStack, nSec, (n) => (n > 40 ? 9 : 11));

  const phaseSubsectionStack = s.phaseSubsectionStacked;
  const ctxSub = document.getElementById("reportChartPhaseSubsection");
  const nSub = phaseSubsectionStack?.labels?.length || 1;
  renderStatusStackedHBar(ctxSub, phaseSubsectionStack, nSub, (n) => (n > 40 ? 9 : 11));

  const objectStack = s.objectStacked;
  const wrapObj = document.getElementById("reportChartObjectWrap");
  const nObj = objectStack?.labels?.length || 1;
  setReportScrollableChartHeight(wrapObj, nObj, { maxPx: 1200, minPx: 280, rowPx: 24 });
  const ctx5 = document.getElementById("reportChartObject");
  renderStatusStackedHBar(ctx5, objectStack, nObj, (n) => (n > 45 ? 8 : n > 25 ? 10 : 11));

  const deptStack = s.departmentStacked;
  const wrapDept = document.getElementById("reportChartDepartmentWrap");
  const nDept = deptStack?.labels?.length || 1;
  if (wrapDept) {
    setReportScrollableChartHeight(wrapDept, nDept, { maxPx: 1200, minPx: 260, rowPx: 24 });
  }
  const ctxDept = document.getElementById("reportChartDepartment");
  if (ctxDept && deptStack?.datasets?.length) {
    const dlDeptStacked = hasDl
      ? {
          anchor: "center",
          align: "center",
          formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
          color: "#1e293b",
          font: { weight: "600", size: 10 },
          clamp: true,
          clip: false,
          display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0
        }
      : { display: false };
    reportChartInstances.push(
      new Chart(ctxDept, {
        type: "bar",
        data: {
          labels: deptStack.labels,
          datasets: deptStack.datasets
        },
        options: {
          ...common,
          ...REPORT_CHART_LABEL_SAFE_LAYOUT,
          indexAxis: "y",
          datasets: {
            bar: {
              maxBarThickness: 28,
              categoryPercentage: 0.82,
              barPercentage: 0.9
            }
          },
          scales: {
            x: {
              stacked: true,
              ...REPORT_HBAR_SCALE_X
            },
            y: {
              stacked: true,
              ticks: {
                autoSkip: false,
                font: { size: nDept > 45 ? 8 : nDept > 25 ? 10 : 11 }
              }
            }
          },
          plugins: {
            legend: {
              display: true,
              position: "top",
              labels: { usePointStyle: true, padding: 12, font: { size: 11 } }
            },
            datalabels: dlDeptStacked
          }
        }
      })
    );
  }

  const respStack = s.responsibleStacked;
  const wrapResp = document.getElementById("reportChartResponsibleWrap");
  const nResp = respStack?.labels?.length || 1;
  if (wrapResp) {
    setReportScrollableChartHeight(wrapResp, nResp, { maxPx: 900, minPx: 240, rowPx: 24 });
  }
  const ctx6 = document.getElementById("reportChartResponsible");
  if (ctx6 && respStack?.datasets?.length) {
    const dlRespStacked = hasDl
      ? {
          anchor: "center",
          align: "center",
          formatter: (v) => (Number(v) > 0 ? String(Math.round(Number(v))) : ""),
          color: "#1e293b",
          font: { weight: "600", size: 10 },
          clamp: true,
          clip: false,
          display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0
        }
      : { display: false };
    reportChartInstances.push(
      new Chart(ctx6, {
        type: "bar",
        data: {
          labels: respStack.labels,
          datasets: respStack.datasets
        },
        options: {
          ...common,
          ...REPORT_CHART_LABEL_SAFE_LAYOUT,
          indexAxis: "y",
          datasets: {
            bar: {
              maxBarThickness: 28,
              categoryPercentage: 0.82,
              barPercentage: 0.9
            }
          },
          scales: {
            x: {
              stacked: true,
              ...REPORT_HBAR_SCALE_X
            },
            y: {
              stacked: true,
              ticks: {
                autoSkip: false,
                font: { size: nResp > 35 ? 9 : 11 }
              }
            }
          },
          plugins: {
            legend: {
              display: true,
              position: "top",
              labels: { usePointStyle: true, padding: 12, font: { size: 11 } }
            },
            datalabels: dlRespStacked
          }
        }
      })
    );
  }

  const ctxPriorityDonut = document.getElementById("reportChartPriorityDonut");
  if (ctxPriorityDonut) {
    reportChartInstances.push(
      new Chart(ctxPriorityDonut, {
        type: "doughnut",
        data: {
          labels: prLabelsFinal,
          datasets: [
            {
              data: prDataFinal,
              backgroundColor: prLabelsFinal.map((_, i) => colorRot(i + 2)),
              borderColor: "#f8fafc",
              borderWidth: 2
            }
          ]
        },
        options: {
          ...common,
          cutout: "52%",
          plugins: {
            legend: {
              position: "bottom",
              labels: { usePointStyle: true, padding: 10, font: { size: 11 } }
            },
            datalabels: dlDonut
          }
        }
      })
    );
  }
}

function getSectionById(sectionId) {
  return sections.find((item) => item.id === sectionId);
}

function getEmployeesList() {
  const employees = getSectionById("employees");
  if (!employees) return [];
  return getUniqueValues(employees.rows, EMPLOYEE_COLUMNS.fullName);
}

function normalizeUzPhone(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw) return DEFAULT_PHONE_PREFIX;
  let s = raw.replace(/[^\d+]/g, "");
  if (s.startsWith("00")) s = `+${s.slice(2)}`;
  if (!s.startsWith("+")) s = `+${s.replace(/\D/g, "")}`;
  const digits = s.slice(1).replace(/\D/g, "").slice(0, PHONE_MAX_DIGITS);
  return `+${digits}`;
}

function sanitizePhoneInputValue(rawValue) {
  const n = normalizeUzPhone(rawValue);
  if (n === "+") return DEFAULT_PHONE_PREFIX;
  const digits = n.replace(/\D/g, "");
  const dial = detectDialCodeByPhone(n);
  const rule = getPhoneRuleByDial(dial);
  return `+${digits.slice(0, rule.max)}`;
}

function enforcePhoneKeyInput(event) {
  const key = String(event.key || "");
  if (!key) return;
  if (event.ctrlKey || event.metaKey || event.altKey) return;
  const allowedControl = new Set(["Backspace", "Delete", "ArrowLeft", "ArrowRight", "Home", "End", "Tab"]);
  if (allowedControl.has(key)) return;
  if (/^\d$/.test(key)) return;
  if (key === "+") {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    const start = Number(input.selectionStart ?? 0);
    if (start === 0 && !input.value.includes("+")) return;
  }
  event.preventDefault();
}

function attachStrictPhoneInputBehavior(input, onAfterSanitize, onMaxReached) {
  if (!(input instanceof HTMLInputElement)) return;
  const syncMaxLength = () => {
    const dial = detectDialCodeByPhone(input.value);
    const rule = getPhoneRuleByDial(dial);
    input.maxLength = rule.max + 1; // + для символа "+"
    return rule.max;
  };
  let prevDigitsLen = getPhoneDigitsCount(input.value);
  syncMaxLength();
  input.addEventListener("keydown", enforcePhoneKeyInput);
  input.addEventListener("input", () => {
    input.value = sanitizePhoneInputValue(input.value);
    const maxDigits = syncMaxLength();
    onAfterSanitize?.();
    const digitsLen = getPhoneDigitsCount(input.value);
    if (digitsLen >= maxDigits && prevDigitsLen < maxDigits) {
      onMaxReached?.();
    }
    prevDigitsLen = digitsLen;
  });
  input.addEventListener("paste", () => {
    requestAnimationFrame(() => {
      input.value = sanitizePhoneInputValue(input.value);
      const maxDigits = syncMaxLength();
      onAfterSanitize?.();
      const digitsLen = getPhoneDigitsCount(input.value);
      if (digitsLen >= maxDigits && prevDigitsLen < maxDigits) {
        onMaxReached?.();
      }
      prevDigitsLen = digitsLen;
    });
  });
}

function formatUzPhoneDisplay(normalizedPhone) {
  const n = normalizeUzPhone(normalizedPhone);
  return n === "+" ? DEFAULT_PHONE_PREFIX : n;
}

/** Номер считается пригодным для сравнения, когда в нём есть хотя бы 8 цифр. */
function employeePhoneLocalCompleteNormalized(normalizedPhone) {
  const digits = String(normalizedPhone || "").replace(/\D/g, "");
  return digits.length >= 8;
}

const DIAL_TO_ISO = [
  ["998", "UZ"], ["7", "RU"], ["380", "UA"], ["375", "BY"], ["1", "US"],
  ["90", "TR"], ["971", "AE"], ["966", "SA"], ["44", "GB"], ["49", "DE"],
  ["33", "FR"], ["39", "IT"], ["34", "ES"], ["48", "PL"], ["995", "GE"],
  ["994", "AZ"], ["996", "KG"], ["992", "TJ"], ["993", "TM"], ["20", "EG"],
  ["91", "IN"], ["92", "PK"], ["86", "CN"], ["81", "JP"], ["82", "KR"],
  ["61", "AU"], ["55", "BR"], ["52", "MX"], ["62", "ID"], ["63", "PH"],
  ["65", "SG"], ["60", "MY"], ["66", "TH"], ["84", "VN"], ["98", "IR"]
];

const COUNTRY_NAME_BY_ISO = {
  UZ: "Узбекистан", RU: "Россия", UA: "Украина", BY: "Беларусь", US: "США",
  TR: "Турция", AE: "ОАЭ", SA: "Саудовская Аравия", GB: "Великобритания", DE: "Германия",
  FR: "Франция", IT: "Италия", ES: "Испания", PL: "Польша", GE: "Грузия",
  AZ: "Азербайджан", KG: "Кыргызстан", TJ: "Таджикистан", TM: "Туркменистан", EG: "Египет",
  IN: "Индия", PK: "Пакистан", CN: "Китай", JP: "Япония", KR: "Южная Корея",
  AU: "Австралия", BR: "Бразилия", MX: "Мексика", ID: "Индонезия", PH: "Филиппины",
  SG: "Сингапур", MY: "Малайзия", TH: "Таиланд", VN: "Вьетнам", IR: "Иран"
};

/** Ограничения по длине в формате E.164 (общее количество цифр без "+"). */
const PHONE_TOTAL_DIGITS_BY_DIAL = {
  "998": { min: 12, max: 12 }, // UZ: 998 + 9
  "7": { min: 11, max: 11 },   // RU/KZ: 7 + 10
  "380": { min: 12, max: 12 },
  "375": { min: 12, max: 12 },
  "1": { min: 11, max: 11 },   // US/CA NANP
  "90": { min: 12, max: 12 },
  "971": { min: 12, max: 12 },
  "966": { min: 12, max: 12 },
  "44": { min: 12, max: 12 },
  "49": { min: 11, max: 13 },
  "33": { min: 11, max: 11 },
  "39": { min: 11, max: 13 },
  "34": { min: 11, max: 11 },
  "48": { min: 11, max: 11 },
  "995": { min: 12, max: 12 },
  "994": { min: 12, max: 12 },
  "996": { min: 12, max: 12 },
  "992": { min: 12, max: 12 },
  "993": { min: 11, max: 11 },
  "20": { min: 12, max: 12 },
  "91": { min: 12, max: 12 },
  "92": { min: 12, max: 12 },
  "86": { min: 13, max: 13 },
  "81": { min: 11, max: 12 },
  "82": { min: 11, max: 12 },
  "61": { min: 11, max: 11 },
  "55": { min: 12, max: 13 },
  "52": { min: 12, max: 12 },
  "62": { min: 10, max: 13 },
  "63": { min: 12, max: 12 },
  "65": { min: 10, max: 10 },
  "60": { min: 10, max: 12 },
  "66": { min: 11, max: 11 },
  "84": { min: 11, max: 11 },
  "98": { min: 12, max: 12 }
};

function flagEmojiFromIso2(iso2) {
  const code = String(iso2 || "").trim().toUpperCase();
  if (!/^[A-Z]{2}$/.test(code)) return "🌐";
  return String.fromCodePoint(...Array.from(code).map((c) => 127397 + c.charCodeAt(0)));
}

function flagSvgUrlByIso(iso2) {
  const code = String(iso2 || "").trim().toLowerCase();
  if (!/^[a-z]{2}$/.test(code)) return "";
  return `https://flagcdn.com/${code}.svg`;
}

function globeSvgDataUrl() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="14" viewBox="0 0 18 14"><rect width="18" height="14" rx="2" fill="#eef3f9"/><circle cx="9" cy="7" r="4.2" fill="none" stroke="#6b7c93" stroke-width="1"/><path d="M4.8 7h8.4M9 2.8c1.6 1.1 1.6 7.3 0 8.4M9 2.8c-1.6 1.1-1.6 7.3 0 8.4" stroke="#6b7c93" stroke-width="1" fill="none"/></svg>`;
  return `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;
}

function detectCountryIsoByPhone(rawPhone) {
  const digits = normalizeUzPhone(rawPhone).replace(/\D/g, "");
  if (!digits) return "";
  let best = "";
  let bestIso = "";
  for (const [dial, iso] of DIAL_TO_ISO) {
    if (digits.startsWith(dial) && dial.length > best.length) {
      best = dial;
      bestIso = iso;
    }
  }
  return bestIso;
}

function phoneFlagByValue(rawPhone) {
  const iso = detectCountryIsoByPhone(rawPhone);
  return flagEmojiFromIso2(iso);
}

function getPhoneDigitsCount(rawPhone) {
  return normalizeUzPhone(rawPhone).replace(/\D/g, "").length;
}

function detectDialCodeByPhone(rawPhone) {
  const digits = normalizeUzPhone(rawPhone).replace(/\D/g, "");
  if (!digits) return "";
  let best = "";
  for (const [dial] of DIAL_TO_ISO) {
    if (digits.startsWith(dial) && dial.length > best.length) {
      best = dial;
    }
  }
  return best;
}

function getPhoneRuleByDial(dial) {
  const d = String(dial || "").trim();
  return PHONE_TOTAL_DIGITS_BY_DIAL[d] || { min: PHONE_MIN_DIGITS, max: PHONE_MAX_DIGITS };
}

function getPhoneLengthHint(rawPhone) {
  const dial = detectDialCodeByPhone(rawPhone);
  const rule = getPhoneRuleByDial(dial);
  return rule.min === rule.max ? `${rule.max}` : `${rule.min}-${rule.max}`;
}

function isPhoneLengthValid(rawPhone) {
  const normalized = normalizeUzPhone(rawPhone);
  const len = normalized.replace(/\D/g, "").length;
  const dial = detectDialCodeByPhone(normalized);
  const rule = getPhoneRuleByDial(dial);
  return len >= rule.min && len <= rule.max;
}

function buildCountryPhoneOptions() {
  return DIAL_TO_ISO
    .map(([dial, iso]) => ({
      dial,
      iso,
      flagUrl: flagSvgUrlByIso(iso),
      name: COUNTRY_NAME_BY_ISO[iso] || iso
    }))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
}

function applyCountryDialToPhone(rawPhone, selectedDial) {
  const normalized = normalizeUzPhone(rawPhone);
  const digits = normalized.replace(/\D/g, "");
  const prevDial = detectDialCodeByPhone(normalized);
  const national = prevDial && digits.startsWith(prevDial) ? digits.slice(prevDial.length) : digits;
  return normalizeUzPhone(`+${selectedDial}${national}`);
}

function setCaretAfterDialCode(input) {
  if (!(input instanceof HTMLInputElement)) return;
  const dial = detectDialCodeByPhone(input.value);
  const pos = Math.max(1, 1 + String(dial || "").length);
  requestAnimationFrame(() => {
    try {
      input.setSelectionRange(pos, pos);
    } catch (_) {
      /* noop */
    }
  });
}

function openLoginCountryPickerModal() {
  if (!phoneInput) return;
  const options = buildCountryPhoneOptions();
  const currentDial = detectDialCodeByPhone(phoneInput.value);
  let selectedDial = currentDial || "998";

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>Выбор страны</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск страны или кода..." />
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const render = () => {
    const q = String(search?.value || "").trim().toLowerCase();
    const filtered = options.filter((opt) => {
      if (!q) return true;
      return opt.name.toLowerCase().includes(q) || opt.iso.toLowerCase().includes(q) || `+${opt.dial}`.includes(q);
    });
      list.innerHTML = filtered.length
        ? filtered.map((opt) => `
        <label class="responsible-option-item">
          <input type="radio" name="countryDial" value="${opt.dial}" ${opt.dial === selectedDial ? "checked" : ""} />
          <span class="responsible-option-name"><img class="country-flag-svg" src="${escapeHtmlAttr(opt.flagUrl || globeSvgDataUrl())}" alt="" />${escapeHtmlText(opt.name)}</span>
          <span class="responsible-option-role">+${opt.dial}</span>
        </label>
      `).join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "radio") return;
    selectedDial = String(target.value || "").trim();
  });
  search?.addEventListener("input", render);
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    phoneInput.focus();
  });
  applyBtn?.addEventListener("click", () => {
    phoneInput.value = applyCountryDialToPhone(phoneInput.value, selectedDial);
    enforceUzPhonePrefix();
    overlay.remove();
    phoneInput.focus();
    setCaretAfterDialCode(phoneInput);
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      phoneInput.focus();
    }
  });

  render();
  search?.focus();
}

function updateLoginPhoneFlag() {
  if (!loginPhoneFlag || !phoneInput) return;
  const iso = detectCountryIsoByPhone(phoneInput.value);
  const svg = flagSvgUrlByIso(iso);
  loginPhoneFlag.src = svg || globeSvgDataUrl();
  loginPhoneFlag.alt = iso ? String(COUNTRY_NAME_BY_ISO[iso] || iso) : "Страна";
  loginPhoneFlag.title = iso ? `Страна: ${String(COUNTRY_NAME_BY_ISO[iso] || iso)}` : "Страна не определена";
}

function firstRowIndexWithSameCompletePhone(rows, phoneNormalized) {
  if (!employeePhoneLocalCompleteNormalized(phoneNormalized)) return -1;
  for (let i = 0; i < rows.length; i++) {
    const p = normalizeUzPhone(rows[i][EMPLOYEE_COLUMNS.phone] || "");
    if (employeePhoneLocalCompleteNormalized(p) && p === phoneNormalized) {
      return i;
    }
  }
  return -1;
}

function firstRowIndexWithSameNormalizedName(rows, fullNameRaw) {
  const name = normalizePersonName(fullNameRaw || "");
  if (!name) return -1;
  const key = name.toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    const n = normalizePersonName(rows[i][EMPLOYEE_COLUMNS.fullName] || "");
    if (n && n.toLowerCase() === key) {
      return i;
    }
  }
  return -1;
}

/** Оставляет первую строку с данным телефоном/ФИО, у остальных очищает поле (при загрузке/синхронизации). */
function dedupeEmployeesInPlace(employeesSection) {
  const rows = employeesSection.rows;
  rows.forEach((row, rowIndex) => {
    const p = normalizeUzPhone(row[EMPLOYEE_COLUMNS.phone] || "");
    if (employeePhoneLocalCompleteNormalized(p)) {
      const first = firstRowIndexWithSameCompletePhone(rows, p);
      if (first !== -1 && first !== rowIndex) {
        row[EMPLOYEE_COLUMNS.phone] = DEFAULT_PHONE_PREFIX;
        row[EMPLOYEE_COLUMNS.chatId] = "";
        row[EMPLOYEE_COLUMNS.activity] = row[EMPLOYEE_COLUMNS.telegram] === "Подключен" ? "Активен" : "Не активен";
      }
    }
    const name = normalizePersonName(row[EMPLOYEE_COLUMNS.fullName] || "");
    if (name) {
      const first = firstRowIndexWithSameNormalizedName(rows, name);
      if (first !== -1 && first !== rowIndex) {
        row[EMPLOYEE_COLUMNS.fullName] = "";
      }
    }
  });
}

/** При ручном вводе: дубликат телефона или ФИО — предупреждение и сброс поля (канон — первая строка). */
function enforceEmployeeUniquenessAfterEdit(section, rowIndex) {
  if (section.id !== "employees") return false;
  const rows = section.rows;
  const row = rows[rowIndex];
  if (!row) return false;
  let changed = false;

  const p = normalizeUzPhone(row[EMPLOYEE_COLUMNS.phone] || "");
  if (employeePhoneLocalCompleteNormalized(p)) {
    const first = firstRowIndexWithSameCompletePhone(rows, p);
    if (first !== -1 && first !== rowIndex) {
      window.alert("Этот номер телефона уже указан у другого сотрудника. Поле очищено.");
      row[EMPLOYEE_COLUMNS.phone] = DEFAULT_PHONE_PREFIX;
      row[EMPLOYEE_COLUMNS.chatId] = "";
      row[EMPLOYEE_COLUMNS.activity] = row[EMPLOYEE_COLUMNS.telegram] === "Подключен" ? "Активен" : "Не активен";
      changed = true;
    }
  }

  const name = normalizePersonName(row[EMPLOYEE_COLUMNS.fullName] || "");
  if (name) {
    const first = firstRowIndexWithSameNormalizedName(rows, name);
    if (first !== -1 && first !== rowIndex) {
      window.alert("Такое ФИО уже указано у другого сотрудника. Поле очищено.");
      row[EMPLOYEE_COLUMNS.fullName] = "";
      changed = true;
    }
  }

  return changed;
}

function getEmployeeByPhone(phoneValue) {
  const employeesSection = getSectionById("employees");
  if (!employeesSection) return null;
  const normalized = normalizeUzPhone(phoneValue);
  return employeesSection.rows.find((row) => normalizeUzPhone(row[4]) === normalized) || null;
}

function maskEmployeePhoneForMessage(phoneRaw) {
  const normalized = normalizeUzPhone(phoneRaw);
  const digits = normalized.replace(/\D/g, "");
  if (!digits) return "—";
  const head = digits.slice(0, 2);
  return `${head}${"*".repeat(Math.max(0, digits.length - 2))}`;
}

function buildEmployeeOnboardingMessage(row) {
  const botUsername = String(displaySettings.telegramBotUsername || "").trim();
  const bot = botUsername ? `@${botUsername}` : "не задан";
  const fio = String(row[EMPLOYEE_COLUMNS.fullName] || "").trim() || "—";
  const phoneMasked = maskEmployeePhoneForMessage(row[EMPLOYEE_COLUMNS.phone]);
  const pass = normalizeUzPhone(row[EMPLOYEE_COLUMNS.phone]).replace(/\D/g, "").slice(-4) || "****";
  const site = String(location.origin || "").trim() || "—";
  return [`Бот: ${bot}`, `Сайт: ${site}`, `ФИО: ${fio}`, `Ваш номер: ${phoneMasked}`, `Пароль: ${pass}`].join("\n");
}

function parseEmployeesCell(value) {
  return String(value || "")
    .split(",")
    .map((name) => name.trim())
    .filter(Boolean);
}

function parseTaskAssigneeNames(value) {
  const out = [];
  const seen = new Set();
  parseEmployeesCell(value).forEach((nameRaw) => {
    const name = normalizePersonName(nameRaw);
    if (!name) return;
    const key = name.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    out.push(name);
  });
  return out;
}

function getTaskIdForMultiState(taskRow) {
  return String(taskRow?.[TASK_COLUMNS.number] ?? "").trim();
}

function getTaskMultiAssigneeMap(taskId, { create = false } = {}) {
  const id = String(taskId || "").trim();
  if (!id) return null;
  if (!taskMultiState || typeof taskMultiState !== "object") {
    taskMultiState = {};
  }
  if (!taskMultiState[id] || typeof taskMultiState[id] !== "object" || Array.isArray(taskMultiState[id])) {
    if (!create) return null;
    taskMultiState[id] = {};
  }
  return taskMultiState[id];
}

function cleanupTaskMultiStateForRow(taskRow) {
  const taskId = getTaskIdForMultiState(taskRow);
  if (!taskId) return;
  const assignees = parseTaskAssigneeNames(taskRow[TASK_COLUMNS.assignedResponsible]);
  if (assignees.length <= 1) {
    delete taskMultiState[taskId];
    return;
  }
  const map = getTaskMultiAssigneeMap(taskId, { create: true });
  const known = new Set(assignees.map((x) => x.toLowerCase()));
  assignees.forEach((name) => {
    if (!map[name] || typeof map[name] !== "object" || Array.isArray(map[name])) {
      map[name] = {};
    }
    if (!String(map[name].status || "").trim()) {
      const baseStatus = String(taskRow[TASK_COLUMNS.status] || "").trim();
      map[name].status = baseStatus || "Новый";
    }
    if (!String(map[name].updatedAt || "").trim()) {
      map[name].updatedAt = String(taskRow[TASK_COLUMNS.lastSentAt] || "").trim() || "—";
    }
  });
  Object.keys(map).forEach((name) => {
    if (!known.has(String(name).toLowerCase())) delete map[name];
  });
}

function normalizeTaskMultiStateStore() {
  if (!taskMultiState || typeof taskMultiState !== "object" || Array.isArray(taskMultiState)) {
    taskMultiState = {};
  }
  const tasksRows = getSectionById("tasks")?.rows || [];
  const existingTaskIds = new Set();
  tasksRows.forEach((row) => {
    const id = getTaskIdForMultiState(row);
    if (!id) return;
    existingTaskIds.add(id);
    cleanupTaskMultiStateForRow(row);
  });
  Object.keys(taskMultiState).forEach((taskId) => {
    if (!existingTaskIds.has(taskId)) delete taskMultiState[taskId];
  });
}

function normalizeTaskIdValue(value) {
  return String(value ?? "").trim();
}

function readNumericTaskId(value) {
  const n = Number(normalizeTaskIdValue(value));
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : null;
}

function getMaxTaskIdAcrossSystem() {
  let maxId = 0;
  const taskRows = getSectionById("tasks")?.rows || [];
  taskRows.forEach((row) => {
    const n = readNumericTaskId(row?.[TASK_COLUMNS.number]);
    if (n && n > maxId) maxId = n;
  });
  const trashRows = getTrashRows("tasks");
  trashRows.forEach((item) => {
    const n = readNumericTaskId(item?.row?.[TASK_COLUMNS.number]);
    if (n && n > maxId) maxId = n;
  });
  return maxId;
}

function ensureTaskIdCounter() {
  const maxId = getMaxTaskIdAcrossSystem();
  let counter = Number(displaySettings.taskIdCounter);
  if (!Number.isFinite(counter) || counter < 1) {
    counter = maxId + 1;
  }
  if (counter <= maxId) {
    counter = maxId + 1;
  }
  displaySettings.taskIdCounter = counter;
  return counter;
}

function allocateNextTaskId() {
  const next = ensureTaskIdCounter();
  displaySettings.taskIdCounter = next + 1;
  saveDisplaySettings();
  return String(next);
}

function getPendingImportedTaskIdsSet() {
  if (!Array.isArray(displaySettings.pendingImportedTaskIds)) {
    displaySettings.pendingImportedTaskIds = [];
  }
  return new Set(displaySettings.pendingImportedTaskIds.map((id) => normalizeTaskIdValue(id)).filter(Boolean));
}

function syncPendingImportedTaskIds(taskIdSet, { save = true } = {}) {
  displaySettings.pendingImportedTaskIds = Array.from(taskIdSet)
    .map((id) => normalizeTaskIdValue(id))
    .filter(Boolean);
  if (save) saveDisplaySettings();
}

function cleanupPendingImportedTaskIds({ save = true } = {}) {
  const tasks = getSectionById("tasks")?.rows || [];
  const actual = new Set(tasks.map((row) => normalizeTaskIdValue(row?.[TASK_COLUMNS.number])).filter(Boolean));
  const pending = getPendingImportedTaskIdsSet();
  let changed = false;
  for (const id of Array.from(pending)) {
    if (!actual.has(id)) {
      pending.delete(id);
      changed = true;
    }
  }
  if (changed) syncPendingImportedTaskIds(pending, { save });
  return pending;
}

function markTaskAsPendingImported(taskId, { save = true } = {}) {
  const id = normalizeTaskIdValue(taskId);
  if (!id) return;
  const pending = cleanupPendingImportedTaskIds({ save: false });
  if (!pending.has(id)) {
    pending.add(id);
    syncPendingImportedTaskIds(pending, { save });
  }
}

function markTaskAsSentImported(taskId, { save = true } = {}) {
  const id = normalizeTaskIdValue(taskId);
  if (!id) return;
  const pending = cleanupPendingImportedTaskIds({ save: false });
  if (pending.delete(id)) {
    syncPendingImportedTaskIds(pending, { save });
  }
}

function getEmployeeNameSet() {
  const rows = getSectionById("employees")?.rows || [];
  return new Set(rows.map((row) => normalizePersonName(row?.[EMPLOYEE_COLUMNS.fullName])).filter(Boolean));
}

function hasSystemEmployeeName(fullName, employeeNameSet = null) {
  const name = normalizePersonName(fullName);
  if (!name) return false;
  const set = employeeNameSet || getEmployeeNameSet();
  return set.has(name);
}

function normalizeTaskImportPriority(raw) {
  const v = normalizeTaskPriorityValue(raw);
  if (!v) return "Средний";
  if (PRIORITY_OPTIONS.includes(v)) return v;
  const n = String(v).trim().toLowerCase();
  if (n.includes("крит")) return "Критический";
  if (n.includes("выс")) return "Высокий";
  if (n.includes("сред") || n.includes("низ")) return "Средний";
  return "Средний";
}

function normalizeTaskImportDate(raw) {
  const text = String(raw || "").trim();
  if (!text) return "";
  const ru = parseRuDateStringToParts(text);
  if (ru) return formatDatePartsStorage(ru.day, ru.month, ru.year);
  const iso = parseHtmlDateValue(text);
  if (iso) return formatDatePartsStorage(iso.day, iso.month, iso.year);
  const slash = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (slash) {
    const day = Number(slash[1]);
    const month = Number(slash[2]);
    const year = Number(slash[3]);
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return formatDatePartsStorage(day, month, year);
    }
  }
  return text;
}

function normalizeTaskImportCellValue(raw) {
  return String(raw ?? "")
    .replace(/\r\n?/g, "\n")
    .replace(/\u00a0/g, " ")
    .trim();
}

function parseTabularText(rawText) {
  const text = String(rawText || "").replace(/\r\n?/g, "\n");
  const rows = text
    .split("\n")
    .map((line) => line.split("\t").map((cell) => normalizeTaskImportCellValue(cell)));
  return rows.filter((row) => row.some((cell) => cell !== ""));
}

function mapTaskImportColumnsFromHeader(headerRow, { allowIndexFallback = false } = {}) {
  const out = {};
  const byNorm = new Map();
  (headerRow || []).forEach((label, index) => {
    const key = normalizeTaskColumnLabel(label);
    if (!key || byNorm.has(key)) return;
    byNorm.set(key, index);
  });
  let matched = 0;
  TASK_IMPORT_COLUMNS.forEach((col, fallbackIndex) => {
    let idx = -1;
    for (const alias of col.aliases) {
      const probe = byNorm.get(normalizeTaskColumnLabel(alias));
      if (Number.isInteger(probe)) {
        idx = probe;
        break;
      }
    }
    if (allowIndexFallback && idx < 0 && fallbackIndex < (headerRow || []).length) idx = fallbackIndex;
    if (idx >= 0) matched += 1;
    out[col.key] = idx;
  });
  return { map: out, matched };
}

function parseTaskImportPayload(rawText) {
  const matrix = parseTabularText(rawText);
  if (!matrix.length) return { ok: false, message: "Вставьте данные из Excel (Ctrl+V)." };

  const firstRow = matrix[0];
  const mapped = mapTaskImportColumnsFromHeader(firstRow, { allowIndexFallback: false });
  const hasHeader = mapped.matched >= Math.ceil(TASK_IMPORT_COLUMNS.length * 0.6);
  const dataRows = hasHeader ? matrix.slice(1) : matrix;
  const colMap = hasHeader ? mapped.map : mapTaskImportColumnsFromHeader([], { allowIndexFallback: true }).map;
  const parsedRows = dataRows
    .map((row) => {
      const item = {};
      TASK_IMPORT_COLUMNS.forEach((col, index) => {
        const sourceIndex = hasHeader ? colMap[col.key] : index;
        item[col.key] = Number.isInteger(sourceIndex) && sourceIndex >= 0 && sourceIndex < row.length ? row[sourceIndex] : "";
      });
      return item;
    })
    .filter((item) => Object.values(item).some((v) => String(v || "").trim() !== ""));

  if (!parsedRows.length) return { ok: false, message: "Не удалось найти строки задач в вставленных данных." };
  return { ok: true, rows: parsedRows, hasHeader };
}

function createTaskRowFromImport(values, taskId) {
  const row = new Array(TASK_COLUMNS.lastSentAt + 1).fill("");
  row[TASK_COLUMNS.number] = normalizeTaskIdValue(taskId);
  row[TASK_COLUMNS.object] = normalizeTaskImportCellValue(values.object);
  row[TASK_COLUMNS.status] = "Новый";
  row[TASK_COLUMNS.priority] = normalizeTaskImportPriority(values.priority);
  row[TASK_COLUMNS.addedDate] = normalizeTaskImportDate(values.addedDate) || getTodayRuDate();
  row[TASK_COLUMNS.phase] = normalizeTaskImportCellValue(values.phase);
  row[TASK_COLUMNS.phaseSection] = normalizeTaskImportCellValue(values.phaseSection);
  row[TASK_COLUMNS.phaseSubsection] = normalizeTaskImportCellValue(values.phaseSubsection);
  row[TASK_COLUMNS.task] = normalizeTaskImportCellValue(values.task);
  row[TASK_COLUMNS.responsible] = normalizeTaskImportCellValue(values.responsible);
  row[TASK_COLUMNS.assignedResponsible] = normalizeTaskImportCellValue(values.assignedResponsible);
  row[TASK_COLUMNS.note] = normalizeTaskImportCellValue(values.note);
  row[TASK_COLUMNS.plan] = "";
  row[TASK_COLUMNS.fact] = "";
  row[TASK_COLUMNS.dueDate] = normalizeTaskImportDate(values.dueDate);
  row[TASK_COLUMNS.closedDate] = "";
  row[TASK_COLUMNS.mediaBefore] = "";
  row[TASK_COLUMNS.mediaAfter] = "";
  row[TASK_COLUMNS.readState] = composeTaskReadState(false, "—");
  row[TASK_COLUMNS.lastSentAt] = "—";
  return row;
}

function getCatalogValueSet(sectionId) {
  const rows = getSectionById(sectionId)?.rows || [];
  return new Set(rows.map((row) => String(row?.[1] || "").trim()).filter(Boolean));
}

function inspectImportedHierarchyValue(values, catalogs = null) {
  const catalogSets = catalogs || {
    phases: getCatalogValueSet("phases"),
    sections: getCatalogValueSet("phaseSections"),
    subsections: getCatalogValueSet("phaseSubsections")
  };
  const phase = normalizeTaskImportCellValue(values.phase);
  const phaseSection = normalizeTaskImportCellValue(values.phaseSection);
  const phaseSubsection = normalizeTaskImportCellValue(values.phaseSubsection);
  const missing = [];
  if (phase && !catalogSets.phases.has(phase)) missing.push(`Фаза: ${phase}`);
  if (phaseSection && !catalogSets.sections.has(phaseSection)) missing.push(`Раздел: ${phaseSection}`);
  if (phaseSubsection && !catalogSets.subsections.has(phaseSubsection)) missing.push(`Подраздел: ${phaseSubsection}`);
  return {
    phase,
    phaseSection,
    phaseSubsection,
    missing,
    hasMissing: missing.length > 0
  };
}

function ensureHierarchyValuesInCatalogs(values) {
  const phase = normalizeTaskImportCellValue(values.phase);
  const phaseSection = normalizeTaskImportCellValue(values.phaseSection);
  const phaseSubsection = normalizeTaskImportCellValue(values.phaseSubsection);
  if (phase) upsertCatalogValue("phases", 1, phase);
  if (phaseSection) upsertCatalogValue("phaseSections", 1, phaseSection);
  if (phaseSubsection) upsertCatalogValue("phaseSubsections", 1, phaseSubsection);
}

function ensureResponsibleHierarchyLink(phase, phaseSection, phaseSubsection) {
  const p = String(phase || "").trim();
  const s = String(phaseSection || "").trim();
  const ss = String(phaseSubsection || "").trim();
  if (!p || !s || !ss) return;
  const dataSection = getSectionById("data");
  if (!dataSection) return;
  const exists = dataSection.rows.some((row) =>
    String(row?.[1] || "").trim() === p
    && String(row?.[2] || "").trim() === s
    && String(row?.[3] || "").trim() === ss
  );
  if (exists) return;
  const numericIds = dataSection.rows.map((row) => Number(row?.[0])).filter((n) => Number.isFinite(n));
  const nextId = numericIds.length ? Math.max(...numericIds) + 1 : dataSection.rows.length + 1;
  dataSection.rows.push([String(nextId), p, s, ss, "", "Добавлено при импорте"]);
}

function getDataRowByHierarchy(phase, phaseSection, phaseSubsection) {
  const dataSection = getSectionById("data");
  if (!dataSection) return null;
  return dataSection.rows.find((row) =>
    String(row[1] || "").trim() === String(phase || "").trim()
    && String(row[2] || "").trim() === String(phaseSection || "").trim()
    && String(row[3] || "").trim() === String(phaseSubsection || "").trim());
}

function getResponsibleByHierarchy(phase, phaseSection, phaseSubsection) {
  const row = getDataRowByHierarchy(phase, phaseSection, phaseSubsection);
  if (!row) return [];
  return parseEmployeesCell(row[4]);
}

function getPhaseSections(phase) {
  const dataSection = getSectionById("data");
  if (!dataSection) return [];
  const filtered = dataSection.rows.filter((row) =>
    !phase || String(row[1] || "").trim() === String(phase || "").trim());
  return getUniqueValues(filtered, 2);
}

function getPhaseSubsections(phase, phaseSection) {
  const dataSection = getSectionById("data");
  if (!dataSection) return [];
  const filtered = dataSection.rows.filter((row) => {
    const phaseMatch = !phase || String(row[1] || "").trim() === String(phase || "").trim();
    const sectionMatch = !phaseSection || String(row[2] || "").trim() === String(phaseSection || "").trim();
    return phaseMatch && sectionMatch;
  });
  return getUniqueValues(filtered, 3);
}

function upsertCatalogValue(sectionId, valueIndex, value) {
  const section = getSectionById(sectionId);
  const normalized = String(value || "").trim();
  if (!section || !normalized) return;
  const exists = section.rows.some((row) => String(row[valueIndex] || "").trim() === normalized);
  if (exists) return;
  const row = new Array(section.columns.length).fill("");
  const numericIds = section.rows.map((item) => Number(item[0])).filter((num) => Number.isFinite(num));
  row[0] = String(numericIds.length ? Math.max(...numericIds) + 1 : section.rows.length + 1);
  row[valueIndex] = normalized;
  section.rows.push(row);
}

function renderSelectFilter(id, label, options, selectedValue) {
  return `
    <label class="filter-field" for="${id}">
      <span>${label}</span>
      <select id="${id}">
        <option value="">Все</option>
        ${options.map((value) => `<option value="${value}" ${value === selectedValue ? "selected" : ""}>${value}</option>`).join("")}
      </select>
    </label>
  `;
}

function renderFilters(section, sectionFilters, isOpen) {
  const commonSearch = sectionFilters.search || "";
  if (!isOpen) {
    return "";
  }

  if (section.id !== "tasks") {
    return `
      <div class="filter-panel">
        <label class="filter-field filter-grow" for="filterSearch">
          <span>Поиск</span>
          <input id="filterSearch" type="text" placeholder="Поиск по таблице..." value="${commonSearch}" />
        </label>
        <button id="filterResetBtn" type="button" class="secondary">Сбросить</button>
      </div>
    `;
  }

  const statusValues = getUniqueValues(section.rows, 2);
  const responsibleValues = getUniqueValues(section.rows, TASK_COLUMNS.assignedResponsible);
  const objectValues = getUniqueValues(section.rows, 1);
  const phaseValues = getUniqueValues(section.rows, TASK_COLUMNS.phase);
  const sectionValues = getUniqueValues(section.rows, TASK_COLUMNS.phaseSection);
  const subsectionValues = getUniqueValues(section.rows, TASK_COLUMNS.phaseSubsection);
  const readValues = ["Прочитано", "Не прочитано"];

  return `
    <div class="filter-panel">
      <label class="filter-field filter-grow" for="filterSearch">
        <span>Поиск</span>
        <input id="filterSearch" type="text" placeholder="Поиск по задачам..." value="${commonSearch}" />
      </label>
      ${renderSelectFilter("filterStatus", "Статус", statusValues, sectionFilters.status || "")}
      ${renderSelectFilter("filterResponsible", "Ответственный", responsibleValues, sectionFilters.responsible || "")}
      ${renderSelectFilter("filterObject", "Объект", objectValues, sectionFilters.object || "")}
      ${renderSelectFilter("filterPhase", "Фаза", phaseValues, sectionFilters.phase || "")}
      ${renderSelectFilter("filterSection", "Раздел", sectionValues, sectionFilters.section || "")}
      ${renderSelectFilter("filterSubsection", "Подраздел", subsectionValues, sectionFilters.subsection || "")}
      ${renderSelectFilter("filterReadState", "Ознакомление", readValues, sectionFilters.readState || "")}
      <button id="filterResetBtn" type="button" class="secondary">Сбросить</button>
    </div>
  `;
}

function getFilteredRows(section, sectionFilters) {
  const normalizedSearch = String(sectionFilters.search || "").trim().toLowerCase();
  const activeStatusTab = statusTabBySection[section.id] || "all";
  if (activeStatusTab === "trash") {
    const trash = getTrashRows(section.id);
    return trash
      .map((item, index) => ({ row: item.row, rowIndex: index, deletedAt: item.deletedAt, expiresAt: item.expiresAt }))
      .filter(({ row }) => !normalizedSearch || row.some((cell) => String(cell).toLowerCase().includes(normalizedSearch)));
  }

  const pendingImportedIds = section.id === "tasks" ? cleanupPendingImportedTaskIds({ save: false }) : null;
  return section.rows
    .map((row, rowIndex) => ({ row, rowIndex }))
    .filter(({ row }) => {
    const searchMatch = !normalizedSearch
      || row.some((cell) => String(cell).toLowerCase().includes(normalizedSearch));

    if (!searchMatch) return false;

    if (section.id !== "tasks") return true;

    const taskId = normalizeTaskIdValue(row[TASK_COLUMNS.number]);
    const isPendingImported = pendingImportedIds?.has(taskId) === true;
    const statusTabMatch = activeStatusTab === "all"
      || (activeStatusTab === TASKS_UNSENT_TAB_ID ? isPendingImported : row[TASK_COLUMNS.status] === activeStatusTab);
    if (!statusTabMatch) return false;

    const statusMatch = !sectionFilters.status || row[TASK_COLUMNS.status] === sectionFilters.status;
    const responsibleMatch = !sectionFilters.responsible || row[TASK_COLUMNS.assignedResponsible] === sectionFilters.responsible;
    const objectMatch = !sectionFilters.object || row[TASK_COLUMNS.object] === sectionFilters.object;
    const phaseMatch = !sectionFilters.phase || row[TASK_COLUMNS.phase] === sectionFilters.phase;
    const sectionMatch = !sectionFilters.section || row[TASK_COLUMNS.phaseSection] === sectionFilters.section;
    const subsectionMatch = !sectionFilters.subsection || row[TASK_COLUMNS.phaseSubsection] === sectionFilters.subsection;
    const readStateLabel = getTaskReadStateParts(row[TASK_COLUMNS.readState]).statusText;
    const readStateMatch = !sectionFilters.readState || readStateLabel === sectionFilters.readState;

    return statusMatch && responsibleMatch && objectMatch && phaseMatch && sectionMatch && subsectionMatch && readStateMatch;
  });
}

function getWideColumnClass(colIndex) {
  if (colIndex === TASK_COLUMNS.task || colIndex === TASK_COLUMNS.note || colIndex === TASK_COLUMNS.plan) {
    return "wide-text-col";
  }
  return "";
}

function isMediaColumn(colIndex) {
  return colIndex === TASK_COLUMNS.mediaBefore || colIndex === TASK_COLUMNS.mediaAfter;
}

function isReadonlyColumn(section, colIndex) {
  // Системный ID/№ всегда только для чтения.
  if (colIndex === 0 || isMediaColumn(colIndex)) {
    return true;
  }
  if (section.id === "employees" && (colIndex === EMPLOYEE_COLUMNS.chatId || colIndex === EMPLOYEE_COLUMNS.activity)) {
    return true;
  }
  if (section.id === "roles" && colIndex === 2) {
    return true;
  }
  if (section.id === "departments" && colIndex === 3) {
    return true;
  }
  return false;
}

function getVisibleColumnIndexes(section) {
  ensureColumnDisplayState(section);
  const visibility = visibleColumnsBySection[section.id];
  const order = columnOrderBySection[section.id];
  return order.filter((index) => visibility[index] !== false);
}

function ensureColumnDisplayState(section) {
  const FIXED_COLUMN_INDEX = 0;
  const size = section.columns.length;
  if (!visibleColumnsBySection[section.id]) {
    visibleColumnsBySection[section.id] = section.columns.map(() => true);
  }
  const visibility = visibleColumnsBySection[section.id];
  if (visibility.length !== size) {
    while (visibility.length < size) visibility.push(true);
    visibility.length = size;
  }
  // ID/№ всегда видимый и фиксированный.
  if (size > 0) visibility[FIXED_COLUMN_INDEX] = true;

  if (!Array.isArray(columnOrderBySection[section.id])) {
    columnOrderBySection[section.id] = section.columns.map((_, index) => index);
  }
  const rawOrder = columnOrderBySection[section.id];
  const seen = new Set();
  const normalized = [];
  for (const index of rawOrder) {
    const n = Number(index);
    if (!Number.isInteger(n) || n < 0 || n >= size || seen.has(n)) continue;
    seen.add(n);
    normalized.push(n);
  }
  for (let i = 0; i < size; i += 1) {
    if (!seen.has(i)) normalized.push(i);
  }
  // ID/№ всегда в начале порядка.
  const fixedPos = normalized.indexOf(FIXED_COLUMN_INDEX);
  if (fixedPos > 0) {
    normalized.splice(fixedPos, 1);
    normalized.unshift(FIXED_COLUMN_INDEX);
  }
  columnOrderBySection[section.id] = normalized;
}

function renderStatusTabs(section) {
  if (section.id !== "tasks") return "";
  const pendingImportedIds = cleanupPendingImportedTaskIds({ save: false });
  const activeStatusTab = statusTabBySection[section.id] || "all";

  const tabsHtml = STATUS_TABS.map((tab) => {
    const count = tab.id === "trash"
      ? getTrashRows(section.id).length
      : tab.id === TASKS_UNSENT_TAB_ID
        ? section.rows.filter((row) => pendingImportedIds.has(normalizeTaskIdValue(row[TASK_COLUMNS.number]))).length
        : section.rows.filter((row) => tab.id === "all" || row[TASK_COLUMNS.status] === tab.id).length;
    const activeClass = activeStatusTab === tab.id ? "active" : "";
    return `<button type="button" class="status-tab-btn ${activeClass}" data-status-tab="${tab.id}">${tab.label} <span>${count}</span></button>`;
  }).join("");

  return `<div class="status-tabs-row">${tabsHtml}</div>`;
}

function renderTasksScreenModeSwitch(section) {
  if (section.id !== "tasks") return "";
  return `
    <div class="tasks-screen-switch-row">
      <span class="tasks-screen-switch-label">Сводная / Объект</span>
      <div class="tasks-segment-group" role="radiogroup" aria-label="Переключение вида задач">
        <label class="tasks-segment">
          <input type="radio" name="tasksQuickBrowseMode" value="flat" ${displaySettings.tasksListBrowseMode !== "byObject" ? "checked" : ""} />
          <span>Сводная</span>
        </label>
        <label class="tasks-segment">
          <input type="radio" name="tasksQuickBrowseMode" value="byObject" ${displaySettings.tasksListBrowseMode === "byObject" ? "checked" : ""} />
          <span>Объект</span>
        </label>
      </div>
    </div>
  `;
}

function resolveObjectPhotoSourceUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (raw === "-" || raw === "—" || raw.toLowerCase() === "null" || raw.toLowerCase() === "undefined") return "";
  if (/^(data:|blob:)/i.test(raw)) return raw;
  if (/^https?:\/\//i.test(raw)) return raw;
  if (raw.startsWith("/media/")) return raw;
  if (raw.startsWith("media/")) return `/${raw}`;
  if (raw.startsWith("/")) return raw;
  return `/media/${encodeURIComponent(raw)}`;
}

function renderObjectPhotoCell(row, value, rowIndex) {
  const oid = String(row[OBJECT_COLUMNS.id] ?? (rowIndex >= 0 ? rowIndex : "")).trim();
  const pkey = oid ? `obj-ph-${oid}` : "";
  const preview = pkey ? objectPhotoPreviewStore[pkey] : null;
  const fileName = String(value || "").trim();
  const safeFileName = resolveObjectPhotoSourceUrl(fileName) ? fileName : "";
  const displayName = getMediaDisplayName(safeFileName);
  const fallbackSrc = resolveObjectPhotoSourceUrl(fileName);
  const imgSrcRaw = preview?.url || fallbackSrc;
  const imgSrc = imgSrcRaw ? escapeHtmlAttr(imgSrcRaw) : "";
  if (imgSrcRaw) {
    return `<div class="object-photo-slot object-photo-slot--has-img" data-object-id="${escapeHtmlAttr(oid)}">
      <div class="object-photo-slot-visual">
        <img src="${imgSrc}" alt="" class="object-photo-thumb" />
        <button type="button" class="object-photo-remove" data-object-photo-remove="1" title="Удалить фото">×</button>
      </div>
    </div>`;
  }
  if (safeFileName) {
    return `<div class="object-photo-slot object-photo-slot--nostore" data-object-id="${escapeHtmlAttr(oid)}">
      <span class="object-photo-filename-hint" title="${escapeHtmlAttr(safeFileName)}">Файл: ${escapeHtmlText(displayName)}</span>
      <div class="object-photo-slot-actions">
        <button type="button" class="object-photo-add" data-object-photo-add="1" title="Заменить фото">Заменить</button>
        <button type="button" class="object-photo-remove" data-object-photo-remove="1" title="Удалить">×</button>
      </div>
    </div>`;
  }
  return `<div class="object-photo-slot object-photo-slot--empty" data-object-id="${escapeHtmlAttr(oid)}">
    <button type="button" class="object-photo-add" data-object-photo-add="1" title="Выбрать одно фото объекта">+</button>
  </div>`;
}

function getTaskAssigneeProgressSummary(taskRow) {
  const names = parseTaskAssigneeNames(taskRow?.[TASK_COLUMNS.assignedResponsible]);
  const total = names.length;
  if (total <= 1) return null;
  const taskId = getTaskIdForMultiState(taskRow);
  const map = getTaskMultiAssigneeMap(taskId) || {};
  let closed = 0;
  names.forEach((name) => {
    const st = String(map?.[name]?.status || taskRow?.[TASK_COLUMNS.status] || "").trim();
    if (st === "Закрыт") closed += 1;
  });
  return { total, closed };
}

function renderTaskTitleCell(taskRow, rowIndex) {
  const text = escapeHtmlText(String(taskRow?.[TASK_COLUMNS.task] || ""));
  return text;
}

function renderTaskAccordionReadonlyCell(taskRow, colIndex, assigneeName, assigneeState, subId) {
  const parentValue = taskRow?.[colIndex];
  const taskId = getTaskIdForMultiState(taskRow);
  if (colIndex === TASK_COLUMNS.number) {
    return escapeHtmlText(subId);
  }
  if (colIndex === TASK_COLUMNS.assignedResponsible) {
    const shown = escapeHtmlText(assigneeName || "—");
    if (!taskId || !assigneeName) return shown;
    return `<span class="task-sub-assignee-edit" data-task-id="${escapeHtmlAttr(taskId)}" data-assignee-old="${escapeHtmlAttr(assigneeName)}" title="Сменить ответственного в подзадаче">${shown}</span>`;
  }
  if (colIndex === TASK_COLUMNS.status) {
    const status = String(assigneeState?.status || taskRow?.[TASK_COLUMNS.status] || "Новый").trim() || "Новый";
    return `<span class="status-badge status-${slugify(status)}">${escapeHtmlText(status)}</span>`;
  }
  if (colIndex === TASK_COLUMNS.plan) {
    const text = String(assigneeState?.comment || taskRow?.[TASK_COLUMNS.plan] || "").trim();
    return escapeHtmlText(text || "—");
  }
  if (colIndex === TASK_COLUMNS.readState) {
    const readAt = String(assigneeState?.readAt || "").trim();
    if (readAt) {
      return `<span class="task-read-state is-read">Прочитано</span><br><span class="task-read-time">${escapeHtmlText(readAt)}</span>`;
    }
    return `<span class="task-read-state is-unread">Не прочитано</span><br><span class="task-read-time">—</span>`;
  }
  if (isMediaColumn(colIndex)) {
    const items = getMediaItems(parentValue);
    return escapeHtmlText(items.length ? `Фото: ${items.length}` : "—");
  }
  if (colIndex === TASK_COLUMNS.priority) {
    const p = String(parentValue || "").trim();
    return `<span class="priority-text priority-${slugify(p)}">${escapeHtmlText(p || "—")}</span>`;
  }
  if (colIndex === TASK_COLUMNS.closedDate && assigneeState?.closedAt) {
    return escapeHtmlText(String(assigneeState.closedAt));
  }
  if (colIndex === TASK_COLUMNS.addedDate || colIndex === TASK_COLUMNS.dueDate || colIndex === TASK_COLUMNS.closedDate || colIndex === TASK_COLUMNS.lastSentAt) {
    return escapeHtmlText(String(parentValue || "").trim() || "—");
  }
  return escapeHtmlText(String(parentValue ?? "").trim() || "—");
}

function renderTaskAssigneesAccordionRows(taskRow, visibleColumnIndexes, isTrashView = false) {
  const names = parseTaskAssigneeNames(taskRow?.[TASK_COLUMNS.assignedResponsible]);
  if (isTrashView || names.length <= 1) return "";
  const taskId = getTaskIdForMultiState(taskRow);
  const expanded = expandedTaskAssigneeRows.has(taskId);
  const map = getTaskMultiAssigneeMap(taskId) || {};
  const rowsHtml = names
    .map((name, idx) => {
      const state = map?.[name] || {};
      const subId = `${String(taskRow[TASK_COLUMNS.number] || "—")}.${idx + 1}`;
      const tds = visibleColumnIndexes
        .map((colIndex, viewOrder) => {
          const stickyClass = colIndex === 0 && viewOrder === 0 ? "number-col" : "";
          const statusClass = colIndex === TASK_COLUMNS.status ? "status-col" : "";
          const objectClass = colIndex === TASK_COLUMNS.object ? "object-col" : "";
          const mediaClass = isMediaColumn(colIndex) ? "media-col" : "";
          const wideClass = getWideColumnClass(colIndex);
          return `<td class="task-accordion-readonly-cell ${stickyClass} ${statusClass} ${objectClass} ${mediaClass} ${wideClass}">${renderTaskAccordionReadonlyCell(taskRow, colIndex, name, state, subId)}</td>`;
        })
        .join("");
      const trashMetaCells = isTrashView ? `<td class="trash-meta-col">—</td><td class="trash-meta-col">—</td>` : "";
      const actionButtons = `
        <button type="button" class="icon-action-btn task-sub-view-btn" title="Просмотр задачи" data-task-id="${escapeHtmlAttr(taskId)}" data-assignee="${escapeHtmlAttr(name)}">
          <i data-lucide="eye" class="lucide-icon" aria-hidden="true"></i>
        </button>
        <button type="button" class="icon-action-btn task-sub-send-btn" title="Отправить подзадачу" data-task-id="${escapeHtmlAttr(taskId)}" data-assignee="${escapeHtmlAttr(name)}">
          <i data-lucide="send" class="lucide-icon" aria-hidden="true"></i>
        </button>
        <button type="button" class="icon-action-btn danger-btn task-sub-remove-btn" title="Удалить подзадачу" data-task-id="${escapeHtmlAttr(taskId)}" data-assignee="${escapeHtmlAttr(name)}">
          <i data-lucide="trash-2" class="lucide-icon" aria-hidden="true"></i>
        </button>
      `;
      return `
        <tr class="task-assignee-subrow ${expanded ? "" : "hidden"}" data-task-assignees-row="${escapeHtmlAttr(taskId)}">
          <td class="checkbox-col"><input type="checkbox" disabled aria-label="Подзадача ${escapeHtmlAttr(subId)}" /></td>
          ${tds}
          ${trashMetaCells}
          <td class="actions-col task-accordion-actions-col">${actionButtons}</td>
        </tr>
      `;
    })
    .join("");
  return rowsHtml;
}

function renderCellContent(section, row, colIndex, value, rowIndexForPhoto = -1) {
  if (section.id === "objects" && colIndex === OBJECT_COLUMNS.photo) {
    return renderObjectPhotoCell(row, value, rowIndexForPhoto);
  }
  if (section.id === "data" && colIndex === 5) {
    return `<em>${String(value || "")}</em>`;
  }
  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.telegram) {
    const normalized = String(value || "").trim();
    const stateClass = normalized === "Подключен" ? "telegram-state-connected" : "telegram-state-disconnected";
    return `<span class="telegram-state ${stateClass}">${normalized || "Не подключен"}</span>`;
  }
  if (section.id !== "tasks") {
    return value;
  }
  if (colIndex === TASK_COLUMNS.readState) {
    const rs = getTaskReadStateParts(value);
    return `<span class="task-read-state ${rs.isRead ? "is-read" : "is-unread"}">${escapeHtmlText(rs.statusText)}</span><br><span class="task-read-time">${escapeHtmlText(rs.whenText)}</span>`;
  }
  if (colIndex === TASK_COLUMNS.lastSentAt) {
    return String(value || "").trim() || "—";
  }

  if (colIndex === TASK_COLUMNS.number) {
    const summary = getTaskAssigneeProgressSummary(row);
    if (!summary) return escapeHtmlText(String(value || ""));
    const taskId = getTaskIdForMultiState(row);
    return `<span class="task-parent-id-toggle" data-task-assignees-toggle="${escapeHtmlAttr(taskId)}" title="Нажмите на ID, чтобы раскрыть подзадачи">${escapeHtmlText(String(value || ""))}</span>`;
  }

  if (colIndex === TASK_COLUMNS.status) {
    const className = `status-badge status-${slugify(value)}`;
    const summary = getTaskAssigneeProgressSummary(row);
    const summaryHtml = summary ? `<div class="task-accordion-meta">${summary.closed}/${summary.total} закрыто</div>` : "";
    return `<span class="${className}">${value}</span>${summaryHtml}`;
  }

  if (colIndex === TASK_COLUMNS.priority) {
    return `<span class="priority-text priority-${slugify(value)}">${value}</span>`;
  }

  if (colIndex === TASK_COLUMNS.addedDate || colIndex === TASK_COLUMNS.closedDate) {
    if (value == null || value === "") return "";
    return formatStoredDateForDisplay(String(value));
  }

  if (colIndex === TASK_COLUMNS.dueDate) {
    return renderDueDateCell(value);
  }

  if (colIndex === TASK_COLUMNS.task) {
    return renderTaskTitleCell(row, rowIndexForPhoto);
  }

  if (isMediaColumn(colIndex)) {
    return renderMediaSlots(value);
  }

  return value;
}

function renderMediaSlots(value) {
  const items = getMediaItems(value);
  const slots = Array.from({ length: 5 }, (_, index) => items[index] || "+");
  return `<div class="media-slots">${slots
    .map((slot, index) => {
      const isFilled = slot !== "+";
      const slotClass = isFilled ? "media-slot filled" : "media-slot";
      const slotLabel = isFilled ? "Удалить файл" : "Выбрать файл";
      const text = isFilled ? "x" : "+";
      return `<button type="button" class="${slotClass}" data-media-slot="${index}" title="${slotLabel}">${text}</button>`;
    })
    .join("")}</div>`;
}

function getMediaItems(value) {
  return String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .slice(0, 5);
}

function setMediaItems(section, rowIndex, colIndex, items) {
  section.rows[rowIndex][colIndex] = items.filter(Boolean).slice(0, 5).join(", ");
  saveSectionsData();
}

function getMediaSlotKey(taskId, colIndex, slotIndex) {
  return `${taskId}-${colIndex}-${slotIndex}`;
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : "");
    reader.onerror = () => reject(new Error("read_failed"));
    reader.readAsDataURL(file);
  });
}

async function uploadTaskMediaToServer(file) {
  if (!isHostedRuntime() || !getAuthToken()) return null;
  const dataUrl = await readFileAsDataUrl(file);
  if (!dataUrl.startsWith("data:")) return null;
  const r = await fetch("/api/media/upload", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${getAuthToken()}`
    },
    body: JSON.stringify({
      dataUrl,
      fileName: String(file.name || "media")
    })
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j?.url) {
    throw new Error(String(j?.error || `Ошибка ${r.status}`));
  }
  return String(j.url);
}

async function resolveStoredMediaFromFile(file) {
  const uploadedUrl = await uploadTaskMediaToServer(file);
  if (isHostedRuntime() && getAuthToken() && !uploadedUrl) {
    throw new Error("Сервер не вернул URL медиа. Проверьте MEDIA_STORAGE_PATH и Volume в Railway.");
  }
  if (uploadedUrl) {
    return {
      stored: uploadedUrl,
      preview: { name: file.name, type: file.type, url: uploadedUrl }
    };
  }
  return {
    stored: file.name || `media-${Date.now()}.png`,
    preview: { name: file.name, type: file.type, url: URL.createObjectURL(file) }
  };
}

function attachMediaSlotHandlers(section) {
  const mediaCells = Array.from(document.querySelectorAll(".media-col.editable-cell"));
  if (!mediaCells.length) return;

  mediaCells.forEach((cell) => {
    const rowIndex = Number(cell.dataset.rowIndex);
    const colIndex = Number(cell.dataset.colIndex);
    const slots = Array.from(cell.querySelectorAll(".media-slot"));

    slots.forEach((slotButton) => {
      slotButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();

        const slotIndex = Number(slotButton.dataset.mediaSlot);
        const items = getMediaItems(section.rows[rowIndex][colIndex]);
        const hasFile = Boolean(items[slotIndex]);
        const row = section.rows[rowIndex];
        const taskId = row ? String(row[TASK_COLUMNS.number]) : "";
        const slotKey = getMediaSlotKey(taskId, colIndex, slotIndex);

        if (hasFile) {
          items.splice(slotIndex, 1);
          setMediaItems(section, rowIndex, colIndex, items);
          delete mediaPreviewStore[slotKey];
          renderTablePreserveScroll();
          return;
        }

        pickFile(async (file) => {
          if (!file) return;
          const nextItems = getMediaItems(section.rows[rowIndex][colIndex]);
          const resolved = await resolveStoredMediaFromFile(file).catch((e) => {
            window.alert(`Не удалось загрузить медиа: ${String(e?.message || e)}`);
            return null;
          });
          if (!resolved) return;
          nextItems[slotIndex] = resolved.stored;
          setMediaItems(section, rowIndex, colIndex, nextItems);
          mediaPreviewStore[slotKey] = {
            ...resolved.preview
          };
          renderTablePreserveScroll();
        });
      });
    });
  });
}

function attachObjectPhotoHandlers(section) {
  if (section.id !== "objects") return;

  document.querySelectorAll(".object-photo-col.editable-cell .object-photo-add").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const cell = btn.closest(".editable-cell");
      const rowIndex = Number(cell?.dataset.rowIndex);
      const colIndex = Number(cell?.dataset.colIndex);
      if (!Number.isFinite(rowIndex) || colIndex !== OBJECT_COLUMNS.photo) return;
      pickFile((file) => {
        if (!file) return;
        if (!file.type.startsWith("image/")) {
          window.alert("Выберите файл изображения (JPEG, PNG, WebP и т.д.).");
          return;
        }
        const row = section.rows[rowIndex];
        const oid = String(row[OBJECT_COLUMNS.id] ?? rowIndex);
        const pkey = `obj-ph-${oid}`;
        const prev = objectPhotoPreviewStore[pkey];
        if (prev?.url && String(prev.url).startsWith("blob:")) URL.revokeObjectURL(prev.url);
        const reader = new FileReader();
        reader.onload = () => {
          const dataUrl = typeof reader.result === "string" ? reader.result : "";
          if (!dataUrl.startsWith("data:")) return;
          row[colIndex] = file.name;
          objectPhotoPreviewStore[pkey] = {
            name: file.name,
            type: file.type,
            url: dataUrl
          };
          saveObjectPhotoThumbsToStorage();
          saveSectionsData();
          renderTablePreserveScroll();
        };
        reader.onerror = () => window.alert("Не удалось прочитать файл.");
        reader.readAsDataURL(file);
      }, "image/*");
    });
  });

  document.querySelectorAll(".object-photo-col.editable-cell .object-photo-remove").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const cell = btn.closest(".editable-cell");
      const rowIndex = Number(cell?.dataset.rowIndex);
      const colIndex = Number(cell?.dataset.colIndex);
      if (!Number.isFinite(rowIndex) || colIndex !== OBJECT_COLUMNS.photo) return;
      const row = section.rows[rowIndex];
      const oid = String(row[OBJECT_COLUMNS.id] ?? rowIndex);
      const pkey = `obj-ph-${oid}`;
      row[colIndex] = "";
      const prev = objectPhotoPreviewStore[pkey];
      if (prev?.url && String(prev.url).startsWith("blob:")) URL.revokeObjectURL(prev.url);
      delete objectPhotoPreviewStore[pkey];
      saveObjectPhotoThumbsToStorage();
      saveSectionsData();
      renderTablePreserveScroll();
    });
  });
}

function pickFile(onPick, accept) {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = accept || "image/*,video/*,.pdf,.doc,.docx,.xls,.xlsx";
  input.style.display = "none";
  document.body.appendChild(input);

  input.addEventListener("change", () => {
    const file = input.files && input.files[0];
    onPick(file || null);
    input.remove();
  }, { once: true });

  input.click();
}

function getRowHighlightClass(section, row) {
  if (section.id !== "tasks") return "";
  const status = row[TASK_COLUMNS.status];

  if (displaySettings.highlightClosed && status === "Закрыт") {
    return "row-highlight-closed";
  }

  if (displaySettings.highlightNeedDecision && status === STATUS_DECISION) {
    return "row-highlight-need-decision";
  }

  return "";
}

function renderDueDateCell(dateValue) {
  if (!dateValue) {
    return '<div class="due-cell"><span>-</span></div>';
  }

  const dueParts = parseRuDateStringToParts(dateValue);
  if (!dueParts) {
    return `<div class="due-cell"><span>${dateValue}</span></div>`;
  }

  const todayParts = getCalendarDatePartsInTimeZone(new Date(), getServerTimezone());
  if (!todayParts) {
    const shown = formatStoredDateForDisplay(String(dateValue));
    return `<div class="due-cell"><span>${shown}</span></div>`;
  }

  const diffDays = calendarDiffDays(todayParts, dueParts);
  const shown = formatStoredDateForDisplay(String(dateValue));

  if (diffDays >= 0) {
    return `<div class="due-cell"><span>${shown}</span><small class="due-ok">Осталось ${diffDays} дн.</small></div>`;
  }

  return `<div class="due-cell"><span>${shown}</span><small class="due-late">Просрочено ${Math.abs(diffDays)} дн.</small></div>`;
}

function slugify(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "-")
    .replace(/[^a-zа-я0-9-]/g, "");
}

function getSelectedRowsSet(sectionId) {
  if (!selectedRowsBySection[sectionId]) {
    selectedRowsBySection[sectionId] = new Set();
  }
  return selectedRowsBySection[sectionId];
}

function normalizeSelectedRowsAfterDelete(sectionId, deletedIndex) {
  const selectedRows = getSelectedRowsSet(sectionId);
  const nextSelected = new Set();
  selectedRows.forEach((index) => {
    if (index < deletedIndex) {
      nextSelected.add(index);
    } else if (index > deletedIndex) {
      nextSelected.add(index - 1);
    }
  });
  selectedRowsBySection[sectionId] = nextSelected;
}

function attachTableActionHandlers(section, filteredEntries) {
  const selectAllCheckbox = document.getElementById("selectAllRows");
  const rowCheckboxes = Array.from(document.querySelectorAll(".row-checkbox"));
  const viewButtons = Array.from(document.querySelectorAll(".view-row-btn"));
  const restoreButtons = Array.from(document.querySelectorAll(".restore-row-btn"));
  const sendButtons = Array.from(document.querySelectorAll(".send-row-btn"));
  const copyEmployeeMsgButtons = Array.from(document.querySelectorAll(".copy-employee-msg-btn"));
  const deleteButtons = Array.from(document.querySelectorAll(".delete-row-btn"));
  const bulkSendBtn = document.getElementById("bulkSendBtn");
  const bulkDeleteBtn = document.getElementById("bulkDeleteBtn");
  const bulkRestoreBtn = document.getElementById("bulkRestoreBtn");
  const selectedRows = getSelectedRowsSet(getSelectionKey(section.id));
  const trashView = isTrashTab(section.id);

  if (selectAllCheckbox) {
    selectAllCheckbox.addEventListener("change", () => {
      const checked = selectAllCheckbox.checked;
      filteredEntries.forEach((entry) => {
        if (checked) {
          selectedRows.add(entry.rowIndex);
        } else {
          selectedRows.delete(entry.rowIndex);
        }
      });
      renderTable();
    });
  }

  rowCheckboxes.forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const rowIndex = Number(checkbox.dataset.rowIndex);
      if (checkbox.checked) {
        selectedRows.add(rowIndex);
      } else {
        selectedRows.delete(rowIndex);
      }
      renderTable();
    });
  });

  viewButtons.forEach((button) => {
    button.addEventListener("click", () => {
      if (section.id !== "tasks") return;
      const rowIndex = Number(button.dataset.rowIndex);
      const row = section.rows[rowIndex];
      if (!row) return;
      openTaskDetailsModal(section, row, rowIndex);
    });
  });

  sendButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const rowIndex = Number(button.dataset.rowIndex);
      const row = section.rows[rowIndex];
      const rowTitle = row ? row[1] : "";
      if (section.id === "tasks") {
        confirmAction({
          message: `Отправить уведомление в Telegram по задаче «${rowTitle}»?`,
          confirmLabel: "Отправить",
          onConfirm: () => {
            void (async () => {
              const result = await sendTaskRowTelegramNotification(row);
              if (result.ok) {
                const taskId = String(row[TASK_COLUMNS.number] ?? "");
                appendTaskHistoryEntry(
                  taskId,
                  `Telegram: отправлено ${result.okCount}/${result.total}${result.missingNames?.length ? ` (не доставлено: ${result.missingNames.join(", ")})` : ""}; в сообщении — кнопки: смена статуса, комментарий, фото (обрабатываются на сервере при настроенном webhook)`
                );
                button.classList.add("is-sent");
                button.title = `Отправлено: ${rowTitle}`;
                renderTablePreserveScroll();
              }
            })();
          }
        });
        return;
      }
      confirmAction({
        message: `Отправить задачу "${rowTitle}"?`,
        confirmLabel: "Отправить",
        onConfirm: () => {
          button.classList.add("is-sent");
          button.title = `Отправлено: ${rowTitle}`;
        }
      });
    });
  });

  copyEmployeeMsgButtons.forEach((button) => {
    button.addEventListener("click", async () => {
      if (section.id !== "employees") return;
      const rowIndex = Number(button.dataset.rowIndex);
      const row = section.rows[rowIndex];
      if (!row) return;
      const text = buildEmployeeOnboardingMessage(row);
      try {
        await navigator.clipboard.writeText(text);
      } catch (_) {
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        ta.remove();
      }
      button.title = "Скопировано";
      setTimeout(() => {
        button.title = "Скопировать сообщение сотруднику";
      }, 1200);
    });
  });

  deleteButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const rowIndex = Number(button.dataset.rowIndex);
      const guardResult = getDeleteGuardMessage(section.id, rowIndex);
      if (guardResult) {
        window.alert(guardResult);
        return;
      }
      if (section.id === "roles") {
        const roleRow = section.rows[rowIndex];
        const roleName = String(roleRow?.[1] || "").trim();
        if (SYSTEM_ROLES.includes(roleName)) {
          window.alert("Системную должность удалить нельзя.");
          return;
        }
      }
      if (section.id === "departments") {
        const departmentRow = section.rows[rowIndex];
        const departmentName = String(departmentRow?.[1] || "").trim();
        if (SYSTEM_DEPARTMENTS.includes(departmentName)) {
          window.alert("Системный отдел удалить нельзя.");
          return;
        }
      }
      const itemLabel = section.id === "tasks" ? "задачу" : "запись";
      confirmAction({
        message: `Перенести ${itemLabel} в корзину?`,
        confirmLabel: "Да",
        onConfirm: () => {
          moveTaskToTrash(section.id, rowIndex);
          normalizeSelectedRowsAfterDelete(getSelectionKey(section.id), rowIndex);
          saveSectionsData();
          if (section.id === "tasks") saveDisplaySettings();
          saveTrashData();
          renderTable();
        }
      });
    });
  });

  if (bulkDeleteBtn) {
    bulkDeleteBtn.addEventListener("click", () => {
      const guarded = Array.from(selectedRows).map((idx) => ({
        idx,
        message: getDeleteGuardMessage(section.id, idx)
      })).find((item) => item.message);
      if (guarded?.message) {
        window.alert(guarded.message);
        return;
      }
      if (section.id === "roles") {
        const selectedProtected = Array.from(selectedRows).some((idx) => {
          const roleName = String(section.rows[idx]?.[1] || "").trim();
          return SYSTEM_ROLES.includes(roleName);
        });
        if (selectedProtected) {
          window.alert("Системные должности нельзя удалять. Снимите их выделение.");
          return;
        }
      }
      if (section.id === "departments") {
        const selectedProtected = Array.from(selectedRows).some((idx) => {
          const departmentName = String(section.rows[idx]?.[1] || "").trim();
          return SYSTEM_DEPARTMENTS.includes(departmentName);
        });
        if (selectedProtected) {
          window.alert("Системные отделы нельзя удалять. Снимите их выделение.");
          return;
        }
      }
      const itemLabel = section.id === "tasks" ? "задачи" : "записи";
      confirmAction({
        message: `Перенести выбранные ${itemLabel} в корзину?`,
        confirmLabel: "Да",
        onConfirm: () => {
          const toDelete = Array.from(selectedRows).sort((a, b) => b - a);
          toDelete.forEach((rowIndex) => {
            moveTaskToTrash(section.id, rowIndex);
          });
          selectedRows.clear();
          saveSectionsData();
          if (section.id === "tasks") saveDisplaySettings();
          saveTrashData();
          renderTable();
        }
      });
    });
  }

  if (bulkSendBtn) {
    bulkSendBtn.addEventListener("click", () => {
      const payload = Array.from(selectedRows)
        .sort((a, b) => a - b)
        .map((rowIndex) => {
          const row = section.rows[rowIndex];
          if (!row) return null;
          return {
            id: row[0],
            object: row[TASK_COLUMNS.object] || "",
            status: row[TASK_COLUMNS.status] || "",
            priority: row[TASK_COLUMNS.priority] || "",
            task: row[TASK_COLUMNS.task] || "",
            responsible: row[TASK_COLUMNS.responsible] || "",
            dueDate: row[TASK_COLUMNS.dueDate] || ""
          };
        })
        .filter(Boolean);

      confirmAction({
        message: `Отправить выбранные задачи (${payload.length}) в Telegram?`,
        confirmLabel: "Отправить",
        onConfirm: () => {
          void (async () => {
            let okTasks = 0;
            const errors = [];
            const rowIndices = Array.from(selectedRows).sort((a, b) => a - b);
            for (const rowIndex of rowIndices) {
              const row = section.rows[rowIndex];
              if (!row) continue;
              const r = await sendTaskRowTelegramNotification(row, { suppressAlerts: true });
              if (r.ok) {
                okTasks += 1;
                const taskId = String(row[TASK_COLUMNS.number] ?? "");
                appendTaskHistoryEntry(
                  taskId,
                  `Telegram (массово): отправлено ${r.okCount}/${r.total}${r.missingNames?.length ? ` (не доставлено: ${r.missingNames.join(", ")})` : ""}; кнопки в сообщении (сервер/webhook)`
                );
              } else {
                const label = String(row[TASK_COLUMNS.task] || row[TASK_COLUMNS.number] || rowIndex);
                errors.push(`«${label}»: ${r.message || r.reason || "ошибка"}`);
              }
            }
            const summary =
              errors.length === 0
                ? `Готово: успешно обработано задач: ${okTasks} из ${rowIndices.length}.`
                : `Обработано: ${okTasks} из ${rowIndices.length}.\n\nОшибки:\n${errors.slice(0, 12).join("\n")}${errors.length > 12 ? `\n… и ещё ${errors.length - 12}` : ""}`;
            showStatusDialog({
              title: "Отправка в Telegram",
              message: summary,
              type: errors.length === 0 ? "success" : "error"
            });
            renderTablePreserveScroll();
          })();
        }
      });
    });
  }

  restoreButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const rowIndex = Number(button.dataset.rowIndex);
      restoreTaskFromTrash(section.id, rowIndex);
      normalizeSelectedRowsAfterDelete(getSelectionKey(section.id), rowIndex);
      saveSectionsData();
      saveTrashData();
      renderTable();
    });
  });

  if (bulkRestoreBtn) {
    bulkRestoreBtn.addEventListener("click", () => {
      const indexes = Array.from(selectedRows).sort((a, b) => b - a);
      indexes.forEach((index) => {
        restoreTaskFromTrash(section.id, index);
      });
      selectedRows.clear();
      saveSectionsData();
      saveTrashData();
      renderTable();
    });
  }
}

function getDeleteGuardMessage(sectionId, rowIndex) {
  const section = getSectionById(sectionId);
  const row = section?.rows?.[rowIndex];
  if (!row) return "";

  const tasks = getSectionById("tasks")?.rows || [];
  const dataRows = getSectionById("data")?.rows || [];
  if (sectionId === "phases") {
    const phase = String(row[1] || "").trim();
    const inResponsible = dataRows.some((item) => String(item[1] || "").trim() === phase);
    if (inResponsible) return `Фаза "${phase}" используется в таблице "Ответственные".`;
    const inTasks = tasks.some((item) => String(item[TASK_COLUMNS.phase] || "").trim() === phase);
    if (inTasks) return `Фаза "${phase}" используется в таблице "Задачи".`;
  }

  if (sectionId === "phaseSections") {
    const phaseSection = String(row[1] || "").trim();
    const inResponsible = dataRows.some((item) => String(item[2] || "").trim() === phaseSection);
    if (inResponsible) return `Раздел "${phaseSection}" используется в таблице "Ответственные".`;
    const inTasks = tasks.some((item) => String(item[TASK_COLUMNS.phaseSection] || "").trim() === phaseSection);
    if (inTasks) return `Раздел "${phaseSection}" используется в таблице "Задачи".`;
  }

  if (sectionId === "phaseSubsections") {
    const phaseSubsection = String(row[1] || "").trim();
    const inResponsible = dataRows.some((item) => String(item[3] || "").trim() === phaseSubsection);
    if (inResponsible) return `Подраздел "${phaseSubsection}" используется в таблице "Ответственные".`;
    const inTasks = tasks.some((item) => String(item[TASK_COLUMNS.phaseSubsection] || "").trim() === phaseSubsection);
    if (inTasks) return `Подраздел "${phaseSubsection}" используется в таблице "Задачи".`;
  }

  if (sectionId === "departments") {
    const department = String(row[1] || "").trim();
    const employees = getSectionById("employees")?.rows || [];
    const inEmployees = employees.some((item) => String(item[EMPLOYEE_COLUMNS.department] || "").trim() === department);
    if (inEmployees) return `Отдел "${department}" используется в таблице "Сотрудники".`;
  }

  return "";
}

function attachEditableCellHandlers(section) {
  const editableCells = Array.from(document.querySelectorAll(".editable-cell"));

  editableCells.forEach((cell) => {
    cell.addEventListener("click", () => {
      if (cell.querySelector(".cell-editor")) return;
      if (cell.isContentEditable) return;
      const rowIndex = Number(cell.dataset.rowIndex);
      const colIndex = Number(cell.dataset.colIndex);
      if (isReadonlyColumn(section, colIndex)) return;
      openCellEditor(section, cell, rowIndex, colIndex);
    });
  });
}

function openCellEditor(section, cell, rowIndex, colIndex) {
  const row = section.rows[rowIndex] || [];
  if (section.id === "data" && colIndex === 4) {
    openResponsibleMultiSelectModal(section, rowIndex);
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.phase) {
    openTaskHierarchyQuickSelectModal(row, (next) => {
      const prev = {
        phase: String(row[TASK_COLUMNS.phase] || "").trim(),
        section: String(row[TASK_COLUMNS.phaseSection] || "").trim(),
        subsection: String(row[TASK_COLUMNS.phaseSubsection] || "").trim(),
        responsible: String(row[TASK_COLUMNS.assignedResponsible] || "").trim()
      };
      row[TASK_COLUMNS.phase] = String(next.phase || "").trim();
      row[TASK_COLUMNS.phaseSection] = String(next.section || "").trim();
      row[TASK_COLUMNS.phaseSubsection] = String(next.subsection || "").trim();
      row[TASK_COLUMNS.assignedResponsible] = String(next.responsible || "").trim();
      cleanupTaskMultiStateForRow(row);
      if (
        prev.phase !== row[TASK_COLUMNS.phase]
        || prev.section !== row[TASK_COLUMNS.phaseSection]
        || prev.subsection !== row[TASK_COLUMNS.phaseSubsection]
        || prev.responsible !== row[TASK_COLUMNS.assignedResponsible]
      ) {
        appendTaskHistoryEntry(
          String(row[TASK_COLUMNS.number] || ""),
          `Быстрый выбор: Фаза «${shortenHistorySnippet(prev.phase)}» → «${shortenHistorySnippet(row[TASK_COLUMNS.phase])}», Раздел «${shortenHistorySnippet(prev.section)}» → «${shortenHistorySnippet(row[TASK_COLUMNS.phaseSection])}», Подраздел «${shortenHistorySnippet(prev.subsection)}» → «${shortenHistorySnippet(row[TASK_COLUMNS.phaseSubsection])}», Ответственный «${shortenHistorySnippet(prev.responsible)}» → «${shortenHistorySnippet(row[TASK_COLUMNS.assignedResponsible])}»`
        );
      }
      saveSectionsData();
    });
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.object) {
    openSingleLookupModal(
      "Выбор объекта",
      getUniqueValues(getSectionById("objects")?.rows || [], OBJECT_COLUMNS.name),
      row[TASK_COLUMNS.object],
      (value) => {
        row[TASK_COLUMNS.object] = value;
      },
      taskHistoryCtx(section, rowIndex, colIndex),
      { allowClear: true }
    );
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.phaseSection) {
    openSingleLookupModal(
      "Выбор раздела",
      getPhaseSections(row[TASK_COLUMNS.phase]),
      row[TASK_COLUMNS.phaseSection],
      (value) => {
        row[TASK_COLUMNS.phaseSection] = value;
        normalizeRowAfterEdit(section, rowIndex, colIndex);
      },
      taskHistoryCtx(section, rowIndex, colIndex),
      { allowClear: true }
    );
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.phaseSubsection) {
    openSingleLookupModal(
      "Выбор подраздела",
      getPhaseSubsections(row[TASK_COLUMNS.phase], row[TASK_COLUMNS.phaseSection]),
      row[TASK_COLUMNS.phaseSubsection],
      (value) => {
        row[TASK_COLUMNS.phaseSubsection] = value;
        normalizeRowAfterEdit(section, rowIndex, colIndex);
      },
      taskHistoryCtx(section, rowIndex, colIndex),
      { allowClear: true }
    );
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.assignedResponsible) {
    openTaskAssignedMultiSelectModal(row, (value) => {
      const prev = String(row[TASK_COLUMNS.assignedResponsible] || "").trim();
      row[TASK_COLUMNS.assignedResponsible] = value;
      cleanupTaskMultiStateForRow(row);
      if (prev !== String(value || "").trim()) {
        appendTaskHistoryEntry(
          String(row[TASK_COLUMNS.number] || ""),
          `Ответственный: «${shortenHistorySnippet(prev)}» → «${shortenHistorySnippet(value)}»`
        );
      }
      saveSectionsData();
    });
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.responsible) {
    openSingleLookupModal(
      "Выбор контролирующего ответственного",
      getEmployeesList(),
      row[TASK_COLUMNS.responsible],
      (value) => {
        row[TASK_COLUMNS.responsible] = value;
      },
      taskHistoryCtx(section, rowIndex, colIndex)
    );
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.status) {
    openSingleLookupModal(
      "Выбор статуса",
      STATUS_OPTIONS,
      row[TASK_COLUMNS.status],
      (value) => {
        const nextStatus = String(value || "").trim();
        if (nextStatus === "Закрыт") {
          const summary = getTaskAssigneeProgressSummary(row);
          if (summary && summary.closed < summary.total) {
            showStatusDialog({
              title: "Закрытие недоступно",
              message: `Нельзя закрыть родительскую задачу: закрыто ${summary.closed} из ${summary.total} подзадач.`,
              type: "error"
            });
            return;
          }
        }
        row[TASK_COLUMNS.status] = value;
        cleanupTaskMultiStateForRow(row);
      },
      taskHistoryCtx(section, rowIndex, colIndex)
    );
    return;
  }
  if (section.id === "tasks" && colIndex === TASK_COLUMNS.priority) {
    openSingleLookupModal(
      "Выбор приоритета",
      PRIORITY_OPTIONS,
      row[TASK_COLUMNS.priority],
      (value) => {
        row[TASK_COLUMNS.priority] = value;
      },
      taskHistoryCtx(section, rowIndex, colIndex)
    );
    return;
  }
  if (section.id === "tasks" && (colIndex === TASK_COLUMNS.addedDate || colIndex === TASK_COLUMNS.dueDate || colIndex === TASK_COLUMNS.closedDate)) {
    openDatePickerModal(
      "Выбор даты",
      row[colIndex],
      (value) => {
        row[colIndex] = value;
      },
      taskHistoryCtx(section, rowIndex, colIndex)
    );
    return;
  }
  if (section.id === "data" && colIndex === 1) {
    openSingleLookupModal("Выбор фазы", getUniqueValues(getSectionById("phases")?.rows || [], 1), row[1], (value) => {
      row[1] = value;
      normalizeRowAfterEdit(section, rowIndex, colIndex);
    });
    return;
  }
  if (section.id === "data" && colIndex === 2) {
    const phase = String(row[1] || "").trim();
    const fromHierarchy = getPhaseSections(phase);
    const fromCatalog = getUniqueValues(getSectionById("phaseSections")?.rows || [], 1);
    const sectionOptions = Array.from(new Set([...fromHierarchy, ...fromCatalog].filter(Boolean)));
    openSingleLookupModal("Выбор раздела", sectionOptions, row[2], (value) => {
      row[2] = value;
      normalizeRowAfterEdit(section, rowIndex, colIndex);
    });
    return;
  }
  if (section.id === "data" && colIndex === 3) {
    const phase = String(row[1] || "").trim();
    const sec = String(row[2] || "").trim();
    const fromHierarchy = getPhaseSubsections(phase, sec);
    const fromCatalog = getUniqueValues(getSectionById("phaseSubsections")?.rows || [], 1);
    const subsectionOptions = Array.from(new Set([...fromHierarchy, ...fromCatalog].filter(Boolean)));
    openSingleLookupModal("Выбор подраздела", subsectionOptions, row[3], (value) => {
      row[3] = value;
      normalizeRowAfterEdit(section, rowIndex, colIndex);
    });
    return;
  }
  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.position) {
    openSingleLookupModal("Выбор должности", getUniqueValues(getSectionById("roles")?.rows || [], 1), row[EMPLOYEE_COLUMNS.position], (value) => {
      row[EMPLOYEE_COLUMNS.position] = value;
      normalizeRowAfterEdit(section, rowIndex, colIndex);
    });
    return;
  }
  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.department) {
    openSingleLookupModal("Выбор отдела", getUniqueValues(getSectionById("departments")?.rows || [], 1), row[EMPLOYEE_COLUMNS.department], (value) => {
      row[EMPLOYEE_COLUMNS.department] = value;
    });
    return;
  }
  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.phone) {
    openEmployeePhoneEditorModal(section, rowIndex, colIndex);
    return;
  }
  if (section.id === "departments" && colIndex === 2) {
    openEmployeeLookupModal("Выбор руководителя отдела", row[2], (value) => {
      row[2] = value;
    });
    return;
  }
  if (section.id === "objects" && (colIndex === OBJECT_COLUMNS.rp || colIndex === OBJECT_COLUMNS.zrp)) {
    openEmployeeSingleSelectModal(section, rowIndex, colIndex);
    return;
  }

  activeRowBySection[section.id] = rowIndex;
  markFocusedRow(cell);
  let nextCellToFocus = null;

  const queueFocusByTab = (backward = false) => {
    nextCellToFocus = getNextEditableCell(section, rowIndex, colIndex, backward);
  };

  const currentValue = section.rows[rowIndex][colIndex] || "";
  const editor = buildEditor(section, rowIndex, colIndex, currentValue);

  if (!editor) {
    cell.contentEditable = "true";
    cell.classList.add("editing");
    const previousTextValue = cell.textContent || "";
    let canceled = false;
    cell.focus();
    document.getSelection()?.selectAllChildren(cell);

    const onCellKeydown = (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        cell.blur();
        return;
      }
      if (event.key === "Escape") {
        event.preventDefault();
        canceled = true;
        cell.blur();
        return;
      }
      if (event.key === "Tab") {
        event.preventDefault();
        queueFocusByTab(event.shiftKey);
        cell.blur();
      }
    };
    cell.addEventListener("keydown", onCellKeydown);

    cell.addEventListener("blur", () => {
      if (canceled) {
        cell.textContent = previousTextValue;
      } else {
        const newVal = cell.textContent?.trim() || "";
        if (section.id === "tasks" && previousTextValue !== newVal) {
          appendTaskHistoryEntry(
            String(section.rows[rowIndex][TASK_COLUMNS.number]),
            `${section.columns[colIndex]}: «${shortenHistorySnippet(previousTextValue)}» → «${shortenHistorySnippet(newVal)}»`
          );
        }
        section.rows[rowIndex][colIndex] = newVal;
        if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.phone) {
          section.rows[rowIndex].__prevEmployeePhoneForNormalize = previousTextValue;
        }
        normalizeRowAfterEdit(section, rowIndex, colIndex);
      }
      cell.contentEditable = "false";
      cell.classList.remove("editing");
      clearActiveRow(section.id);
      saveSectionsData();
      renderTablePreserveScroll();
      focusQueuedCell(nextCellToFocus);
      cell.removeEventListener("keydown", onCellKeydown);
    }, { once: true });
    return;
  }

  cell.classList.add("editing");
  const previousHtml = cell.innerHTML;
  cell.innerHTML = "";
  cell.appendChild(editor);
  editor.addEventListener("click", (event) => event.stopPropagation());
  editor.addEventListener("mousedown", (event) => event.stopPropagation());
  editor.focus();

  let isCommitted = false;
  const commit = (nextValue) => {
    if (isCommitted) return;
    isCommitted = true;
    const prev = section.rows[rowIndex][colIndex];
    section.rows[rowIndex][colIndex] = nextValue;
    if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.phone) {
      section.rows[rowIndex].__prevEmployeePhoneForNormalize = prev;
    }
    if (section.id === "tasks" && String(prev ?? "") !== String(nextValue ?? "")) {
      appendTaskHistoryEntry(
        String(section.rows[rowIndex][TASK_COLUMNS.number]),
        `${section.columns[colIndex]}: «${shortenHistorySnippet(prev)}» → «${shortenHistorySnippet(nextValue)}»`
      );
    }
    normalizeRowAfterEdit(section, rowIndex, colIndex);
    cleanupEditorResources(editor);
    cell.classList.remove("editing");
    clearActiveRow(section.id);
    saveSectionsData();
    renderTablePreserveScroll();
    focusQueuedCell(nextCellToFocus);
  };

  const rollback = () => {
    cleanupEditorResources(editor);
    cell.classList.remove("editing");
    cell.innerHTML = previousHtml;
    clearActiveRow(section.id);
    renderTablePreserveScroll();
  };

  editor.addEventListener("change", () => {
    if (editor.classList.contains("multi-select-editor")) return;
    commit(getEditorValue(editor));
  });

  const isDropdownOrDate = editor.tagName === "SELECT" || (editor instanceof HTMLInputElement && editor.type === "date");
  const isCustomMultiSelect = editor.classList.contains("multi-select-editor");
  if (!isDropdownOrDate && !isCustomMultiSelect) {
    editor.addEventListener("blur", () => {
      commit(getEditorValue(editor));
    }, { once: true });
  }

  editor.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      event.preventDefault();
      rollback();
      return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      queueFocusByTab(event.shiftKey);
      commit(getEditorValue(editor));
      return;
    }
    if (event.key === "Enter" && editor.tagName !== "SELECT" && !isCustomMultiSelect) {
      event.preventDefault();
      commit(getEditorValue(editor));
    }
  });

  if (isCustomMultiSelect) {
    const applyBtn = editor.querySelector(".multi-select-apply-btn");
    applyBtn?.addEventListener("click", () => {
      commit(getEditorValue(editor));
    });
  }

  if (isDropdownOrDate) {
    editor.addEventListener("blur", () => {
      // Для select/date не коммитим на blur, но снимаем режим фокуса строки.
      if (!isCommitted) {
        rollback();
      }
    }, { once: true });
  }
}

function getNextEditableCell(section, rowIndex, colIndex, backward) {
  const visible = getVisibleColumnIndexes(section).filter((idx) => !isReadonlyColumn(section, idx));
  if (!visible.length) return null;
  const currentPos = visible.indexOf(colIndex);
  const normalizedPos = currentPos >= 0 ? currentPos : 0;
  const step = backward ? -1 : 1;
  const rowDelta = backward ? -1 : 1;

  let nextPos = normalizedPos + step;
  let nextRow = rowIndex;

  if (nextPos >= 0 && nextPos < visible.length) {
    return { rowIndex: nextRow, colIndex: visible[nextPos] };
  }

  nextRow += rowDelta;
  if (nextRow < 0 || nextRow >= section.rows.length) {
    return null;
  }
  nextPos = backward ? visible.length - 1 : 0;
  return { rowIndex: nextRow, colIndex: visible[nextPos] };
}

function focusQueuedCell(targetCell) {
  if (!targetCell) return;
  requestAnimationFrame(() => {
    const selector = `.editable-cell[data-row-index="${targetCell.rowIndex}"][data-col-index="${targetCell.colIndex}"]`;
    const next = document.querySelector(selector);
    if (next instanceof HTMLElement) {
      next.click();
    }
  });
}

function normalizeRowAfterEdit(section, rowIndex, colIndex) {
  const row = section.rows[rowIndex];
  if (!row) return;

  if (section.id === "tasks") {
    if (colIndex === TASK_COLUMNS.phase) {
      row[TASK_COLUMNS.phaseSection] = "";
      row[TASK_COLUMNS.phaseSubsection] = "";
      row[TASK_COLUMNS.assignedResponsible] = "";
      cleanupTaskMultiStateForRow(row);
      return;
    }

    if (colIndex === TASK_COLUMNS.phaseSection) {
      row[TASK_COLUMNS.phaseSubsection] = "";
      row[TASK_COLUMNS.assignedResponsible] = "";
      cleanupTaskMultiStateForRow(row);
      return;
    }

    if (colIndex === TASK_COLUMNS.phaseSubsection) {
      row[TASK_COLUMNS.assignedResponsible] = "";
      cleanupTaskMultiStateForRow(row);
      return;
    }
    if (colIndex === TASK_COLUMNS.assignedResponsible || colIndex === TASK_COLUMNS.status || colIndex === TASK_COLUMNS.number) {
      cleanupTaskMultiStateForRow(row);
    }
    return;
  }

  if (section.id === "employees") {
    if (colIndex === EMPLOYEE_COLUMNS.phone) {
      const prevPhoneRaw = row.__prevEmployeePhoneForNormalize || "";
      row[EMPLOYEE_COLUMNS.phone] = formatUzPhoneDisplay(normalizeUzPhone(row[EMPLOYEE_COLUMNS.phone]));
      resetEmployeeTelegramBindingOnPhoneChange(row, prevPhoneRaw);
      delete row.__prevEmployeePhoneForNormalize;
    }
    if (colIndex === EMPLOYEE_COLUMNS.telegram || colIndex === EMPLOYEE_COLUMNS.phone) {
      applyEmployeeTelegramDerivedFields(row);
    }
    if (colIndex === EMPLOYEE_COLUMNS.position) {
      row[EMPLOYEE_COLUMNS.position] = String(row[EMPLOYEE_COLUMNS.position] || "").trim();
    }
    enforceEmployeeUniquenessAfterEdit(section, rowIndex);
    return;
  }

  if (section.id === "roles") {
    row[2] = getRoleTypeLabel(row[1]);
    return;
  }
  if (section.id === "departments") {
    row[3] = getDepartmentTypeLabel(row[1]);
    return;
  }

  if (section.id === "data") {
    upsertCatalogValue("phases", 1, row[1]);
    upsertCatalogValue("phaseSections", 1, row[2]);
    upsertCatalogValue("phaseSubsections", 1, row[3]);
    return;
  }

  if (section.id === "phaseSections") {
    upsertCatalogValue("phaseSections", 1, row[1]);
    return;
  }

  if (section.id === "phaseSubsections") {
    upsertCatalogValue("phaseSubsections", 1, row[1]);
  }
}

function buildEditor(section, rowIndex, colIndex, currentValue) {
  if (section.id === "tasks") {
    const row = section.rows[rowIndex] || [];
    if (colIndex === TASK_COLUMNS.object) {
      return createSearchableSelectEditor(getUniqueValues(getSectionById("objects").rows, OBJECT_COLUMNS.name), currentValue);
    }
    if (colIndex === TASK_COLUMNS.assignedResponsible) {
      const options = getResponsibleByHierarchy(
        row[TASK_COLUMNS.phase],
        row[TASK_COLUMNS.phaseSection],
        row[TASK_COLUMNS.phaseSubsection]
      );
      return createSearchableSelectEditor(options.length ? options : getEmployeesList(), currentValue);
    }
    if (colIndex === TASK_COLUMNS.responsible) {
      return createSearchableSelectEditor(getEmployeesList(), currentValue);
    }
    if (colIndex === TASK_COLUMNS.phase) {
      return createSearchableSelectEditor(getUniqueValues(getSectionById("phases")?.rows || [], 1), currentValue);
    }
    if (colIndex === TASK_COLUMNS.phaseSection) {
      return createSearchableSelectEditor(getPhaseSections(row[TASK_COLUMNS.phase]), currentValue);
    }
    if (colIndex === TASK_COLUMNS.phaseSubsection) {
      return createSearchableSelectEditor(
        getPhaseSubsections(row[TASK_COLUMNS.phase], row[TASK_COLUMNS.phaseSection]),
        currentValue
      );
    }
    if (colIndex === TASK_COLUMNS.status) {
      return createSelectEditor(STATUS_OPTIONS, currentValue);
    }
    if (colIndex === TASK_COLUMNS.priority) {
      return createSelectEditor(PRIORITY_OPTIONS, currentValue);
    }
    if (colIndex === TASK_COLUMNS.addedDate || colIndex === TASK_COLUMNS.dueDate || colIndex === TASK_COLUMNS.closedDate) {
      return createDateEditor(currentValue);
    }
  }

  if (section.id === "objects" && (colIndex === OBJECT_COLUMNS.rp || colIndex === OBJECT_COLUMNS.zrp)) {
    return createSearchableSelectEditor(getEmployeesList(), currentValue);
  }

  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.position) {
    const roles = getSectionById("roles");
    return createSelectEditor(getUniqueValues(roles?.rows || [], 1), currentValue);
  }

  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.telegram) {
    return createSelectEditor(EMPLOYEE_TELEGRAM_OPTIONS, currentValue);
  }

  if (section.id === "employees" && colIndex === EMPLOYEE_COLUMNS.phone) {
    return createUzPhoneCellEditor(currentValue);
  }

  if (section.id === "data" && colIndex === 1) {
    return createSearchableSelectEditor(getUniqueValues(getSectionById("phases")?.rows || [], 1), currentValue);
  }
  if (section.id === "data" && colIndex === 2) {
    return createSearchableSelectEditor(getUniqueValues(getSectionById("phaseSections")?.rows || [], 1), currentValue);
  }
  if (section.id === "data" && colIndex === 3) {
    return createSearchableSelectEditor(getUniqueValues(getSectionById("phaseSubsections")?.rows || [], 1), currentValue);
  }

  if (section.id === "phases" && colIndex === 1) {
    return createSearchableSelectEditor(getUniqueValues(section.rows, 1), currentValue);
  }

  if (section.id === "phaseSections" && colIndex === 1) {
    return createSearchableSelectEditor(getUniqueValues(section.rows, 1), currentValue);
  }

  if (section.id === "phaseSubsections" && colIndex === 1) {
    return createSearchableSelectEditor(getUniqueValues(section.rows, 1), currentValue);
  }

  return null;
}

function createSelectEditor(options, currentValue) {
  const select = document.createElement("select");
  select.className = "cell-editor";
  const list = Array.from(new Set(options)).filter(Boolean);
  select.innerHTML = list.map((option) => `<option value="${option}" ${option === currentValue ? "selected" : ""}>${option}</option>`).join("");
  return select;
}

function createDateEditor(currentValue) {
  const input = document.createElement("input");
  input.type = "date";
  input.className = "cell-editor";
  input.value = toInputDate(currentValue);
  return input;
}

function createUzPhoneCellEditor(currentValue) {
  const input = document.createElement("input");
  input.type = "tel";
  input.inputMode = "tel";
  input.className = "cell-editor";
  input.maxLength = PHONE_MAX_LENGTH;
  input.pattern = "^\\+[0-9]{8,15}$";
  input.autocomplete = "off";
  input.value = formatUzPhoneDisplay(normalizeUzPhone(currentValue || DEFAULT_PHONE_PREFIX));
  attachStrictPhoneInputBehavior(input);
  return input;
}

function createSearchableSelectEditor(options, currentValue) {
  const input = document.createElement("input");
  input.type = "text";
  input.className = "cell-editor";
  input.value = currentValue || "";
  input.autocomplete = "off";

  const list = Array.from(new Set(options)).filter(Boolean);
  const listId = `cell-list-${Math.random().toString(36).slice(2, 10)}`;
  const datalist = document.createElement("datalist");
  datalist.id = listId;
  datalist.innerHTML = list.map((option) => `<option value="${option}"></option>`).join("");
  document.body.appendChild(datalist);
  input.setAttribute("list", listId);
  input.dataset.datalistId = listId;

  return input;
}

function makeChatIdFromPhone(phoneValue) {
  const normalized = normalizeUzPhone(phoneValue);
  const digits = normalized.replace(/\D/g, "");
  return digits || "";
}

/** Раньше Chat ID подставлялся из 9 цифр номера — для Telegram API это неверно; такие значения сбрасываем. */
function isPhoneDerivedEmployeeChatId(phoneValue, chatIdRaw) {
  const pseudo = makeChatIdFromPhone(phoneValue);
  const cid = String(chatIdRaw ?? "").trim();
  if (!cid || !pseudo) return false;
  return cid === pseudo;
}

function resetEmployeeTelegramBindingOnPhoneChange(row, prevPhoneRaw) {
  const prev = normalizeUzPhone(prevPhoneRaw || "");
  const next = normalizeUzPhone(row[EMPLOYEE_COLUMNS.phone] || "");
  if (!prev || !next || prev === next) return;
  row[EMPLOYEE_COLUMNS.chatId] = "";
  row[EMPLOYEE_COLUMNS.telegram] = "Не подключен";
  row[EMPLOYEE_COLUMNS.activity] = "Не активен";
}

/** При «Подключен» не заполняем Chat ID из телефона — только реальный user id (/start или вручную). */
function applyEmployeeTelegramDerivedFields(row) {
  const isConnected = String(row[EMPLOYEE_COLUMNS.telegram] || "").trim() === "Подключен";
  if (!isConnected) {
    row[EMPLOYEE_COLUMNS.chatId] = "";
    row[EMPLOYEE_COLUMNS.activity] = "Не активен";
    return;
  }
  if (isPhoneDerivedEmployeeChatId(row[EMPLOYEE_COLUMNS.phone], row[EMPLOYEE_COLUMNS.chatId])) {
    row[EMPLOYEE_COLUMNS.chatId] = "";
  }
  row[EMPLOYEE_COLUMNS.activity] = "Активен";
}

function getEmployeesForSelection() {
  const employeesSection = getSectionById("employees");
  if (!employeesSection) return [];
  return employeesSection.rows.map((row) => ({
    fullName: String(row[EMPLOYEE_COLUMNS.fullName] || "").trim(),
    role: String(row[EMPLOYEE_COLUMNS.position] || "").trim()
  })).filter((item) => item.fullName);
}

function getRoleTypeLabel(roleName) {
  return SYSTEM_ROLES.includes(String(roleName || "").trim()) ? "Системный" : "Не системный";
}

function getDepartmentTypeLabel(departmentName) {
  return SYSTEM_DEPARTMENTS.includes(String(departmentName || "").trim()) ? "Системный" : "Не системный";
}

function ensureSystemRoles() {
  const rolesSection = getSectionById("roles");
  if (!rolesSection) return;
  const existing = new Set(rolesSection.rows.map((row) => String(row[1] || "").trim()).filter(Boolean));
  rolesSection.rows.forEach((row) => {
    row[2] = getRoleTypeLabel(row[1]);
  });
  SYSTEM_ROLES.forEach((role) => {
    if (!existing.has(role)) {
      const nextId = String(rolesSection.rows.length + 1);
      rolesSection.rows.push([nextId, role, "Системный"]);
    }
  });
}

function ensureSystemDepartments() {
  const departmentsSection = getSectionById("departments");
  if (!departmentsSection) return;
  const existing = new Set(departmentsSection.rows.map((row) => String(row[1] || "").trim()).filter(Boolean));
  departmentsSection.rows.forEach((row) => {
    row[3] = getDepartmentTypeLabel(row[1]);
  });
  SYSTEM_DEPARTMENTS.forEach((department) => {
    if (!existing.has(department)) {
      const nextId = String(departmentsSection.rows.length + 1);
      departmentsSection.rows.push([nextId, department, "", "Системный"]);
    }
  });
}

function attachTaskAccordionHandlers(section) {
  if (section.id !== "tasks") return;
  document.querySelectorAll(".task-parent-id-toggle").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const taskId = String(button.getAttribute("data-task-assignees-toggle") || "").trim();
      if (!taskId) return;
      if (expandedTaskAssigneeRows.has(taskId)) expandedTaskAssigneeRows.delete(taskId);
      else expandedTaskAssigneeRows.add(taskId);
      renderTablePreserveScroll();
    });
  });
  document.querySelectorAll(".task-sub-assignee-edit").forEach((el) => {
    el.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const taskId = String(el.getAttribute("data-task-id") || "").trim();
      const oldName = normalizePersonName(el.getAttribute("data-assignee-old") || "");
      if (!taskId || !oldName) return;
      const tasks = getSectionById("tasks");
      const row = (tasks?.rows || []).find((r) => String(r[TASK_COLUMNS.number] || "").trim() === taskId);
      if (!row) return;
      const optsByHierarchy = getResponsibleByHierarchy(
        row[TASK_COLUMNS.phase],
        row[TASK_COLUMNS.phaseSection],
        row[TASK_COLUMNS.phaseSubsection]
      );
      const options = (optsByHierarchy.length ? optsByHierarchy : getEmployeesList()).map((x) => normalizePersonName(x)).filter(Boolean);
      openSingleLookupModal(
        "Смена ответственного (подзадача)",
        options,
        oldName,
        (newName) => {
          const next = normalizePersonName(newName);
          if (!next || next === oldName) return;
          replaceTaskSubAssignee(taskId, oldName, next);
        }
      );
    });
  });
  document.querySelectorAll(".task-sub-view-btn").forEach((el) => {
    el.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const taskId = String(el.getAttribute("data-task-id") || "").trim();
      const tasks = getSectionById("tasks");
      const rowIndex = (tasks?.rows || []).findIndex((r) => String(r[TASK_COLUMNS.number] || "").trim() === taskId);
      if (!tasks || rowIndex < 0) return;
      const row = tasks.rows[rowIndex];
      openTaskDetailsModal(tasks, row, rowIndex);
    });
  });
  document.querySelectorAll(".task-sub-send-btn").forEach((el) => {
    el.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const taskId = String(el.getAttribute("data-task-id") || "").trim();
      const assignee = normalizePersonName(el.getAttribute("data-assignee") || "");
      if (!taskId || !assignee) return;
      const tasks = getSectionById("tasks");
      const row = (tasks?.rows || []).find((r) => String(r[TASK_COLUMNS.number] || "").trim() === taskId);
      if (!row) return;
      confirmAction({
        message: `Отправить в Telegram подзадачу ${taskId} для «${assignee}»?`,
        confirmLabel: "Отправить",
        onConfirm: () => {
          void (async () => {
            await sendTaskRowTelegramNotification(row, { targetAssigneeNames: [assignee] });
          })();
        }
      });
    });
  });
  document.querySelectorAll(".task-sub-remove-btn").forEach((el) => {
    el.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const taskId = String(el.getAttribute("data-task-id") || "").trim();
      const assignee = normalizePersonName(el.getAttribute("data-assignee") || "");
      if (!taskId || !assignee) return;
      confirmAction({
        message: `Удалить подзадачу исполнителя «${assignee}»?`,
        confirmLabel: "Удалить",
        onConfirm: () => {
          const tasks = getSectionById("tasks");
          const row = (tasks?.rows || []).find((r) => String(r[TASK_COLUMNS.number] || "").trim() === taskId);
          if (!row) return;
          const list = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
          if (list.length <= 1) {
            showStatusDialog({
              title: "Удаление подзадачи",
              message: "Нельзя удалить последнего ответственного.",
              type: "error"
            });
            return;
          }
          const next = list.filter((x) => String(x).toLowerCase() !== String(assignee).toLowerCase());
          row[TASK_COLUMNS.assignedResponsible] = next.join(", ");
          const state = getTaskMultiAssigneeMap(taskId);
          if (state && state[assignee]) delete state[assignee];
          cleanupTaskMultiStateForRow(row);
          appendTaskHistoryEntry(String(taskId), `Подзадача удалена: ${assignee}`);
          saveSectionsData();
          renderTablePreserveScroll();
        }
      });
    });
  });
}

function replaceTaskSubAssignee(taskId, oldName, newName) {
  const tasks = getSectionById("tasks");
  if (!tasks) return false;
  const row = tasks.rows.find((r) => String(r[TASK_COLUMNS.number] || "").trim() === String(taskId || "").trim());
  if (!row) return false;
  const list = parseTaskAssigneeNames(row[TASK_COLUMNS.assignedResponsible]);
  const oldKey = String(oldName || "").toLowerCase();
  const idx = list.findIndex((n) => String(n).toLowerCase() === oldKey);
  if (idx < 0) return false;
  const duplicateIdx = list.findIndex((n, i) => i !== idx && String(n).toLowerCase() === String(newName || "").toLowerCase());
  if (duplicateIdx >= 0) {
    showStatusDialog({
      title: "Смена ответственного",
      message: "Такой сотрудник уже есть в подзадачах этой задачи.",
      type: "error"
    });
    return false;
  }
  list[idx] = newName;
  row[TASK_COLUMNS.assignedResponsible] = list.join(", ");
  const state = getTaskMultiAssigneeMap(taskId, { create: true });
  if (state && state[oldName]) {
    if (!state[newName]) {
      state[newName] = state[oldName];
    } else {
      state[newName] = {
        ...state[oldName],
        ...state[newName]
      };
    }
    delete state[oldName];
  }
  cleanupTaskMultiStateForRow(row);
  appendTaskHistoryEntry(
    String(taskId),
    `Ответственный (подзадача): «${shortenHistorySnippet(oldName)}» → «${shortenHistorySnippet(newName)}»`
  );
  saveSectionsData();
  return true;
}

function syncEmployeesDerivedFields() {
  const employeesSection = getSectionById("employees");
  if (!employeesSection) return;
  const defaultDepartment = String(getSectionById("departments")?.rows?.[0]?.[1] || "");
  employeesSection.rows.forEach((row) => {
    row[EMPLOYEE_COLUMNS.phone] = formatUzPhoneDisplay(normalizeUzPhone(row[EMPLOYEE_COLUMNS.phone] || DEFAULT_PHONE_PREFIX));
    if (!String(row[EMPLOYEE_COLUMNS.department] || "").trim()) {
      row[EMPLOYEE_COLUMNS.department] = defaultDepartment;
    }
    applyEmployeeTelegramDerivedFields(row);
  });
  dedupeEmployeesInPlace(employeesSection);
}

/** Только цифры — не считаем человекочитаемым названием фазы/раздела/подраздела. */
function isCatalogNumericPlaceholder(s) {
  return /^\d+$/.test(String(s ?? "").trim());
}

/**
 * Достаёт подпись строки справочника фаз (устойчиво к перепутанным колонкам и «номерам вместо текста» из импорта).
 * @param {"phase"|"section"|"subsection"} kind
 */
function pickPhaseCatalogLabel(row, kind) {
  const r = Array.isArray(row) ? row : [];
  const t = (i) => String(r[i] ?? "").trim();
  const pickNonNumeric = (...candidates) => {
    for (const c of candidates) {
      if (c && !isCatalogNumericPlaceholder(c)) return c;
    }
    return "";
  };
  if (kind === "subsection") {
    return pickNonNumeric(t(1), t(3), t(2), t(0)) || t(1) || t(3) || "";
  }
  return pickNonNumeric(t(1), t(2), t(0)) || t(1) || t(2) || "";
}

function normalizePhaseAndSectionCatalogs() {
  const phases = getSectionById("phases");
  if (phases) {
    phases.rows = phases.rows
      .map((row, index) => {
        const value = pickPhaseCatalogLabel(row, "phase");
        return [String(index + 1), value];
      })
      .filter((row) => row[1]);
  }

  const phaseSections = getSectionById("phaseSections");
  if (phaseSections) {
    phaseSections.rows = phaseSections.rows
      .map((row, index) => {
        const value = pickPhaseCatalogLabel(row, "section");
        return [String(index + 1), value];
      })
      .filter((row) => row[1]);
  }

  const phaseSubsections = getSectionById("phaseSubsections");
  if (phaseSubsections) {
    phaseSubsections.rows = phaseSubsections.rows
      .map((row, index) => {
        const value = pickPhaseCatalogLabel(row, "subsection");
        return [String(index + 1), value];
      })
      .filter((row) => row[1]);
  }
}

/** После нормализации: если справочник опустел — подставляем встроенный набор. @returns {boolean} были ли правки */
function ensureNonEmptyPhaseCatalogs() {
  let changed = false;
  for (const id of ["phases", "phaseSections", "phaseSubsections"]) {
    const sec = getSectionById(id);
    if (!sec || !Array.isArray(sec.rows)) continue;
    const names = getUniqueValues(sec.rows, 1);
    const fallback = PHASE_CATALOG_DEFAULT_ROWS[id];
    const onlyNumericLabels = names.length > 0 && names.every((n) => isCatalogNumericPlaceholder(n));
    if (fallback?.length && (names.length === 0 || onlyNumericLabels)) {
      sec.rows = JSON.parse(JSON.stringify(fallback));
      changed = true;
    }
  }
  return changed;
}

function openResponsibleMultiSelectModal(section, rowIndex) {
  const row = section.rows[rowIndex];
  if (!row) return;

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  const selected = new Set(parseEmployeesCell(row[4]));
  const options = getEmployeesForSelection();

  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>Выбор ответственных</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск сотрудника..." />
      <label class="responsible-select-all">
        <input type="checkbox" class="responsible-select-all-checkbox" />
        <span>Выбрать все</span>
      </label>
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Готово</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const selectAllCheckbox = overlay.querySelector(".responsible-select-all-checkbox");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const refreshSelectAllState = (filtered) => {
    if (!selectAllCheckbox) return;
    if (!filtered.length) {
      selectAllCheckbox.checked = false;
      selectAllCheckbox.indeterminate = false;
      return;
    }
    const selectedCount = filtered.filter((item) => selected.has(item.fullName)).length;
    selectAllCheckbox.checked = selectedCount === filtered.length;
    selectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < filtered.length;
  };

  let lastFiltered = options;

  const renderOptions = (query = "") => {
    const normalized = String(query).trim().toLowerCase();
    const filtered = options.filter((item) => (
      item.fullName.toLowerCase().includes(normalized)
      || item.role.toLowerCase().includes(normalized)
    ));
    lastFiltered = filtered;
    list.innerHTML = filtered.length
      ? filtered.map((item) => `
        <label class="responsible-option-item">
          <input type="checkbox" value="${item.fullName}" ${selected.has(item.fullName) ? "checked" : ""} />
          <span class="responsible-option-name">${item.fullName}</span>
          <span class="responsible-option-role">${item.role || "-"}</span>
        </label>
      `).join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
    refreshSelectAllState(filtered);
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") return;
    if (target.checked) {
      selected.add(target.value);
    } else {
      selected.delete(target.value);
    }
  });

  search?.addEventListener("input", () => renderOptions(search.value));
  selectAllCheckbox?.addEventListener("change", () => {
    const shouldSelect = Boolean(selectAllCheckbox.checked);
    lastFiltered.forEach((item) => {
      if (shouldSelect) {
        selected.add(item.fullName);
      } else {
        selected.delete(item.fullName);
      }
    });
    renderOptions(search?.value || "");
  });
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    row[4] = Array.from(selected).join(", ");
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });

  renderOptions();
  search?.focus();
}

function openEmployeeSingleSelectModal(section, rowIndex, colIndex) {
  const row = section.rows[rowIndex];
  if (!row) return;
  const currentValue = String(row[colIndex] || "").trim();
  const requiredRole = colIndex === OBJECT_COLUMNS.rp ? "РП" : "ЗРП";
  const options = getEmployeesForSelection().filter((item) => item.role === requiredRole);

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>Выбор сотрудника (${requiredRole})</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск сотрудника..." />
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");
  let selectedName = currentValue;

  const renderOptions = (query = "") => {
    const normalized = String(query).trim().toLowerCase();
    const filtered = options.filter((item) => (
      item.fullName.toLowerCase().includes(normalized)
      || item.role.toLowerCase().includes(normalized)
    ));
    list.innerHTML = filtered.length
      ? filtered.map((item) => `
        <label class="responsible-option-item">
          <input type="radio" name="singleEmployeeSelect" value="${item.fullName}" ${item.fullName === selectedName ? "checked" : ""} />
          <span class="responsible-option-name">${item.fullName}</span>
          <span class="responsible-option-role">${item.role || "-"}</span>
        </label>
      `).join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "radio") return;
    selectedName = target.value;
  });

  search?.addEventListener("input", () => renderOptions(search.value));
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    row[colIndex] = selectedName || "";
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });

  renderOptions();
  search?.focus();
}

function openSingleLookupModal(title, options, currentValue, onApply, historyCtx, modalOptions = {}) {
  const normalizedOptions = Array.from(new Set((options || []).map((item) => String(item || "").trim()).filter(Boolean)));
  let selectedValue = String(currentValue || "").trim();
  const allowClear = Boolean(modalOptions?.allowClear);

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>${title}</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск..." />
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const renderOptions = (query = "") => {
    const normalized = String(query).trim().toLowerCase();
    const filtered = normalizedOptions.filter((value) => value.toLowerCase().includes(normalized));
    list.innerHTML = filtered.length
      ? filtered.map((value) => `
        <label class="responsible-option-item">
          <input type="radio" name="lookupSingleValue" value="${value}" ${value === selectedValue ? "checked" : ""} />
          <span class="responsible-option-name">${value}</span>
        </label>
      `).join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "radio") return;
    selectedValue = target.value;
  });

  search?.addEventListener("input", () => renderOptions(search.value));
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    const newVal = selectedValue || "";
    if (historyCtx?.taskId && typeof historyCtx.getOld === "function") {
      const oldVal = historyCtx.getOld();
      if (String(oldVal ?? "") !== String(newVal ?? "")) {
        appendTaskHistoryEntry(
          historyCtx.taskId,
          `${historyCtx.columnLabel}: «${shortenHistorySnippet(oldVal)}» → «${shortenHistorySnippet(newVal)}»`
        );
      }
    }
    onApply?.(newVal);
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });

  renderOptions();
  search?.focus();
}

function openTaskHierarchyQuickSelectModal(taskRow, onApply) {
  if (!taskRow) return;
  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  const phases = getUniqueValues(getSectionById("phases")?.rows || [], 1);
  const employees = getEmployeesList();
  const state = {
    phase: String(taskRow[TASK_COLUMNS.phase] || "").trim(),
    section: String(taskRow[TASK_COLUMNS.phaseSection] || "").trim(),
    subsection: String(taskRow[TASK_COLUMNS.phaseSubsection] || "").trim(),
    responsibleList: parseTaskAssigneeNames(taskRow[TASK_COLUMNS.assignedResponsible])
  };

  overlay.innerHTML = `
    <div class="responsible-modal task-hierarchy-modal">
      <h4>Быстрый выбор: фаза → раздел → подраздел → ответственный</h4>
      <div class="task-hierarchy-fixed-grid">
        <div class="task-hierarchy-pane">
          <div class="task-hierarchy-pane-title">Фаза</div>
          <div class="task-hierarchy-list" data-task-hierarchy-list="phase"></div>
        </div>
        <div class="task-hierarchy-pane">
          <div class="task-hierarchy-pane-title">Раздел</div>
          <div class="task-hierarchy-list" data-task-hierarchy-list="section"></div>
        </div>
        <div class="task-hierarchy-pane">
          <div class="task-hierarchy-pane-title">Подраздел</div>
          <div class="task-hierarchy-list" data-task-hierarchy-list="subsection"></div>
        </div>
        <div class="task-hierarchy-pane">
          <div class="task-hierarchy-pane-title">Ответственный (мультивыбор)</div>
          <div class="task-hierarchy-list" data-task-hierarchy-list="responsible"></div>
        </div>
      </div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary task-hierarchy-reset-btn">Сбросить</button>
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const phaseList = overlay.querySelector('[data-task-hierarchy-list="phase"]');
  const sectionList = overlay.querySelector('[data-task-hierarchy-list="section"]');
  const subsectionList = overlay.querySelector('[data-task-hierarchy-list="subsection"]');
  const respList = overlay.querySelector('[data-task-hierarchy-list="responsible"]');
  const resetBtn = overlay.querySelector(".task-hierarchy-reset-btn");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const fillList = (el, options, selectedValue, levelKey) => {
    if (!(el instanceof HTMLElement)) return;
    const list = Array.from(new Set((options || []).map((x) => String(x || "").trim()).filter(Boolean)));
    if (!list.length) {
      el.innerHTML = '<div class="task-hierarchy-empty">Нет данных</div>';
      return;
    }
    el.innerHTML = list
      .map((v) => `
        <button type="button" class="task-hierarchy-item ${
          Array.isArray(selectedValue)
            ? selectedValue.includes(v) ? "is-selected" : ""
            : v === selectedValue ? "is-selected" : ""
        }" data-task-hierarchy-item="${escapeHtmlAttr(levelKey)}" data-value="${escapeHtmlAttr(v)}">${escapeHtmlText(v)}</button>
      `)
      .join("");
  };

  const refreshCascade = () => {
    fillList(phaseList, phases, state.phase, "phase");
    const sectionOptions = state.phase ? getPhaseSections(state.phase) : [];
    if (!sectionOptions.includes(state.section)) state.section = "";
    fillList(sectionList, sectionOptions, state.section, "section");
    const subsectionOptions = state.phase && state.section ? getPhaseSubsections(state.phase, state.section) : [];
    if (!subsectionOptions.includes(state.subsection)) state.subsection = "";
    fillList(subsectionList, subsectionOptions, state.subsection, "subsection");
    const responsibleOptions = getResponsibleByHierarchy(state.phase, state.section, state.subsection);
    const responsibleList = responsibleOptions.length ? responsibleOptions : employees;
    state.responsibleList = state.responsibleList.filter((name) => responsibleList.includes(name));
    fillList(respList, responsibleList, state.responsibleList, "responsible");
  };

  overlay.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
      return;
    }
    if (!target.classList.contains("task-hierarchy-item")) return;
    const level = String(target.getAttribute("data-task-hierarchy-item") || "").trim();
    const value = String(target.getAttribute("data-value") || "").trim();
    if (!level) return;
    if (level === "phase") {
      state.phase = value;
      state.section = "";
      state.subsection = "";
      state.responsibleList = [];
    } else if (level === "section") {
      state.section = value;
      state.subsection = "";
      state.responsibleList = [];
    } else if (level === "subsection") {
      state.subsection = value;
      state.responsibleList = [];
    } else if (level === "responsible") {
      if (state.responsibleList.includes(value)) {
        state.responsibleList = state.responsibleList.filter((x) => x !== value);
      } else {
        state.responsibleList = [...state.responsibleList, value];
      }
    }
    refreshCascade();
  });

  resetBtn?.addEventListener("click", () => {
    state.phase = "";
    state.section = "";
    state.subsection = "";
    state.responsibleList = [];
    refreshCascade();
  });
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    onApply?.({
      phase: state.phase,
      section: state.section,
      subsection: state.subsection,
      responsible: state.responsibleList.join(", ")
    });
    overlay.remove();
    renderTablePreserveScroll();
  });

  refreshCascade();
}

function openTaskAssignedMultiSelectModal(taskRow, onApply) {
  if (!taskRow) return;
  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  const baseOptions = getResponsibleByHierarchy(
    taskRow[TASK_COLUMNS.phase],
    taskRow[TASK_COLUMNS.phaseSection],
    taskRow[TASK_COLUMNS.phaseSubsection]
  );
  const allEmployees = getEmployeesForSelection();
  const fallback = allEmployees.map((x) => x.fullName);
  const names = (baseOptions.length ? baseOptions : fallback).map((x) => normalizePersonName(x)).filter(Boolean);
  const selected = new Set(parseTaskAssigneeNames(taskRow[TASK_COLUMNS.assignedResponsible]));
  const options = names.map((fullName) => {
    const found = allEmployees.find((e) => normalizePersonName(e.fullName) === fullName);
    return { fullName, role: String(found?.role || "").trim() || "—" };
  });

  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>Выбор ответственных (мультивыбор)</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск сотрудника..." />
      <label class="responsible-select-all">
        <input type="checkbox" class="responsible-select-all-checkbox" />
        <span>Выбрать всех</span>
      </label>
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Готово</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const selectAllCheckbox = overlay.querySelector(".responsible-select-all-checkbox");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");
  let lastFiltered = options;

  const refreshSelectAllState = (filtered) => {
    if (!selectAllCheckbox) return;
    if (!filtered.length) {
      selectAllCheckbox.checked = false;
      selectAllCheckbox.indeterminate = false;
      return;
    }
    const selectedCount = filtered.filter((item) => selected.has(item.fullName)).length;
    selectAllCheckbox.checked = selectedCount === filtered.length;
    selectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < filtered.length;
  };

  const renderOptions = (query = "") => {
    const q = String(query || "").trim().toLowerCase();
    const filtered = options.filter((item) => (
      item.fullName.toLowerCase().includes(q) || item.role.toLowerCase().includes(q)
    ));
    lastFiltered = filtered;
    list.innerHTML = filtered.length
      ? filtered.map((item) => `
        <label class="responsible-option-item">
          <input type="checkbox" value="${escapeHtmlAttr(item.fullName)}" ${selected.has(item.fullName) ? "checked" : ""} />
          <span class="responsible-option-name">${escapeHtmlText(item.fullName)}</span>
          <span class="responsible-option-role">${escapeHtmlText(item.role)}</span>
        </label>
      `).join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
    refreshSelectAllState(filtered);
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") return;
    const value = normalizePersonName(target.value);
    if (!value) return;
    if (target.checked) selected.add(value);
    else selected.delete(value);
    refreshSelectAllState(lastFiltered);
  });
  search?.addEventListener("input", () => renderOptions(search.value));
  selectAllCheckbox?.addEventListener("change", () => {
    const shouldSelect = Boolean(selectAllCheckbox.checked);
    lastFiltered.forEach((item) => {
      if (shouldSelect) selected.add(item.fullName);
      else selected.delete(item.fullName);
    });
    renderOptions(search?.value || "");
  });

  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    const values = options.map((o) => o.fullName).filter((name) => selected.has(name));
    onApply?.(values.join(", "));
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });

  renderOptions();
  search?.focus();
}

function openEmployeeLookupModal(title, currentValue, onApply, historyCtx) {
  const employees = getEmployeesForSelection();
  let selectedValue = String(currentValue || "").trim();

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>${title}</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск сотрудника..." />
      <select class="cell-editor responsible-modal-position-filter">
        <option value="">Все должности</option>
      </select>
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        ${allowClear ? '<button type="button" class="secondary responsible-clear-btn">Сбросить</button>' : ""}
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const positionFilter = overlay.querySelector(".responsible-modal-position-filter");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const positions = Array.from(new Set(employees.map((x) => String(x.role || "").trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b, "ru"));
  if (positionFilter) {
    positionFilter.innerHTML = `<option value="">Все должности</option>${positions
      .map((p) => `<option value="${escapeHtmlAttr(p)}">${escapeHtmlText(p)}</option>`)
      .join("")}`;
  }

  const renderOptions = () => {
    const q = String(search?.value || "").trim().toLowerCase();
    const role = String(positionFilter?.value || "").trim();
    const filtered = employees.filter((item) => {
      const fullName = String(item.fullName || "").trim();
      const itemRole = String(item.role || "").trim();
      if (role && itemRole !== role) return false;
      if (!q) return true;
      return fullName.toLowerCase().includes(q) || itemRole.toLowerCase().includes(q);
    });
    list.innerHTML = filtered.length
      ? filtered.map((item) => `
        <label class="responsible-option-item">
          <input type="radio" name="employeeLookupSingleValue" value="${escapeHtmlAttr(item.fullName)}" ${item.fullName === selectedValue ? "checked" : ""} />
          <span class="responsible-option-name">${escapeHtmlText(item.fullName)}</span>
          <span class="responsible-option-role">${escapeHtmlText(item.role || "-")}</span>
        </label>
      `).join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "radio") return;
    selectedValue = target.value;
  });

  search?.addEventListener("input", renderOptions);
  positionFilter?.addEventListener("change", renderOptions);
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    const newVal = selectedValue || "";
    if (historyCtx?.taskId && typeof historyCtx.getOld === "function") {
      const oldVal = historyCtx.getOld();
      if (String(oldVal ?? "") !== String(newVal ?? "")) {
        appendTaskHistoryEntry(
          historyCtx.taskId,
          `${historyCtx.columnLabel}: «${shortenHistorySnippet(oldVal)}» → «${shortenHistorySnippet(newVal)}»`
        );
      }
    }
    onApply?.(newVal);
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });

  renderOptions();
  search?.focus();
}

function openEmployeePhoneEditorModal(section, rowIndex, colIndex) {
  const row = section.rows[rowIndex];
  if (!row) return;
  const prevValue = String(row[colIndex] || "");

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>Номер телефона</h4>
      <div class="phone-input-wrap">
        <button type="button" class="phone-input-flag-btn" id="employeePhoneCountryBtn" title="Выбрать страну">
          <img class="phone-input-flag" id="employeePhoneFlag" alt="" src="" />
        </button>
        <input type="tel" inputmode="tel" class="cell-editor" id="employeePhoneInput" value="${escapeHtmlAttr(
          formatUzPhoneDisplay(normalizeUzPhone(prevValue || DEFAULT_PHONE_PREFIX))
        )}" maxlength="${PHONE_MAX_LENGTH}" autocomplete="off" />
      </div>
      <div class="responsible-option-empty" style="margin-top:8px">Формат: +код_страны номер (длина зависит от страны).</div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Сохранить</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const input = overlay.querySelector("#employeePhoneInput");
  const flag = overlay.querySelector("#employeePhoneFlag");
  const countryBtn = overlay.querySelector("#employeePhoneCountryBtn");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const updateLocalFlag = () => {
    if (!flag || !input) return;
    const iso = detectCountryIsoByPhone(input.value);
    flag.src = flagSvgUrlByIso(iso) || globeSvgDataUrl();
    flag.alt = iso ? String(COUNTRY_NAME_BY_ISO[iso] || iso) : "Страна";
    flag.title = iso ? `Страна: ${String(COUNTRY_NAME_BY_ISO[iso] || iso)}` : "Страна не определена";
  };

  const openCountryPickerForEmployeePhone = () => {
    if (!input) return;
    const options = buildCountryPhoneOptions();
    let selectedDial = detectDialCodeByPhone(input.value) || "998";

    const second = document.createElement("div");
    second.className = "responsible-modal-overlay";
    second.innerHTML = `
      <div class="responsible-modal">
        <h4>Выбор страны</h4>
        <input type="text" class="responsible-modal-search" placeholder="Поиск страны или кода..." />
        <div class="responsible-modal-list"></div>
        <div class="responsible-modal-actions">
          <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
          <button type="button" class="responsible-apply-btn">Выбрать</button>
        </div>
      </div>
    `;
    document.body.appendChild(second);

    const list = second.querySelector(".responsible-modal-list");
    const search = second.querySelector(".responsible-modal-search");
    const secondCancel = second.querySelector(".responsible-cancel-btn");
    const secondApply = second.querySelector(".responsible-apply-btn");

    const render = () => {
      const q = String(search?.value || "").trim().toLowerCase();
      const filtered = options.filter((opt) => {
        if (!q) return true;
        return opt.name.toLowerCase().includes(q) || opt.iso.toLowerCase().includes(q) || `+${opt.dial}`.includes(q);
      });
      list.innerHTML = filtered.length
        ? filtered.map((opt) => `
          <label class="responsible-option-item">
            <input type="radio" name="employeeCountryDial" value="${opt.dial}" ${opt.dial === selectedDial ? "checked" : ""} />
            <span class="responsible-option-name"><img class="country-flag-svg" src="${escapeHtmlAttr(opt.flagUrl || globeSvgDataUrl())}" alt="" />${escapeHtmlText(opt.name)}</span>
            <span class="responsible-option-role">+${opt.dial}</span>
          </label>
        `).join("")
        : '<div class="responsible-option-empty">Ничего не найдено</div>';
    };

    list.addEventListener("change", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLInputElement) || target.type !== "radio") return;
      selectedDial = String(target.value || "").trim();
    });
    search?.addEventListener("input", render);
    secondCancel?.addEventListener("click", () => second.remove());
    secondApply?.addEventListener("click", () => {
      input.value = applyCountryDialToPhone(input.value, selectedDial);
      input.value = formatUzPhoneDisplay(input.value);
      updateLocalFlag();
      second.remove();
      input.focus();
      setCaretAfterDialCode(input);
    });
    second.addEventListener("click", (event) => {
      if (event.target === second) second.remove();
    });
    render();
    search?.focus();
  };

  if (input instanceof HTMLInputElement) {
    attachStrictPhoneInputBehavior(input, updateLocalFlag);
  }
  countryBtn?.addEventListener("click", openCountryPickerForEmployeePhone);
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    const normalized = normalizeUzPhone(input?.value || "");
    if (!isPhoneLengthValid(normalized)) {
      window.alert(`Введите корректный номер (${getPhoneLengthHint(normalized)} цифр после + для выбранной страны).`);
      input?.focus();
      return;
    }
    row.__prevEmployeePhoneForNormalize = prevValue;
    row[colIndex] = normalized;
    normalizeRowAfterEdit(section, rowIndex, colIndex);
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });

  updateLocalFlag();
  input?.focus();
  setCaretAfterDialCode(input);
}

function openReportFilterPickerModal(title, allLabel, items, currentValue, onApply) {
  const normalizedItems = Array.from(
    new Set((items || []).map((item) => String(item || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, "ru"));
  let selectedValue = String(currentValue || "").trim();

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>${escapeHtmlText(title)}</h4>
      <input type="text" class="responsible-modal-search" placeholder="Поиск..." />
      <div class="responsible-modal-list"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const list = overlay.querySelector(".responsible-modal-list");
  const search = overlay.querySelector(".responsible-modal-search");
  const clearBtn = overlay.querySelector(".responsible-clear-btn");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  const renderOptions = (query = "") => {
    const q = String(query).trim().toLowerCase();
    const allRow = `
      <label class="responsible-option-item">
        <input type="radio" name="reportFilterPick" value="" ${selectedValue === "" ? "checked" : ""} />
        <span class="responsible-option-name">${escapeHtmlText(allLabel)}</span>
      </label>
    `;
    const filtered = normalizedItems.filter((value) => value.toLowerCase().includes(q));
    const rest = filtered.length
      ? filtered
          .map(
            (value) => `
        <label class="responsible-option-item">
          <input type="radio" name="reportFilterPick" value="${escapeHtmlAttr(value)}" ${value === selectedValue ? "checked" : ""} />
          <span class="responsible-option-name">${escapeHtmlText(value)}</span>
        </label>
      `
          )
          .join("")
      : '<div class="responsible-option-empty">Ничего не найдено</div>';
    list.innerHTML = allRow + rest;
  };

  list.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "radio") return;
    selectedValue = String(target.value || "");
  });

  search?.addEventListener("input", () => renderOptions(search.value));
  clearBtn?.addEventListener("click", () => {
    const newVal = "";
    if (historyCtx?.taskId && typeof historyCtx.getOld === "function") {
      const oldVal = historyCtx.getOld();
      if (String(oldVal ?? "") !== "") {
        appendTaskHistoryEntry(
          historyCtx.taskId,
          `${historyCtx.columnLabel}: «${shortenHistorySnippet(oldVal)}» → «—»`
        );
      }
    }
    onApply?.(newVal);
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
  });
  applyBtn?.addEventListener("click", () => {
    onApply?.(selectedValue);
    overlay.remove();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
    }
  });

  renderOptions();
  search?.focus();
}

function openDatePickerModal(title, currentRuDate, onApply, historyCtx) {
  let selected = toInputDate(currentRuDate);
  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h4>${title}</h4>
      <input type="date" class="date-modal-input" value="${selected}" />
      <div class="responsible-modal-actions">
        <button type="button" class="secondary responsible-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn">Выбрать</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  const input = overlay.querySelector(".date-modal-input");
  const cancelBtn = overlay.querySelector(".responsible-cancel-btn");
  const applyBtn = overlay.querySelector(".responsible-apply-btn");

  input?.addEventListener("change", () => {
    selected = String(input.value || "");
  });
  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
    renderTablePreserveScroll();
  });
  applyBtn?.addEventListener("click", () => {
    const newVal = fromInputDate(selected);
    if (historyCtx?.taskId && typeof historyCtx.getOld === "function") {
      const oldVal = historyCtx.getOld();
      if (String(oldVal ?? "") !== String(newVal ?? "")) {
        appendTaskHistoryEntry(
          historyCtx.taskId,
          `${historyCtx.columnLabel}: «${shortenHistorySnippet(oldVal)}» → «${shortenHistorySnippet(newVal)}»`
        );
      }
    }
    onApply?.(newVal);
    saveSectionsData();
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      overlay.remove();
      renderTablePreserveScroll();
    }
  });
  input?.focus();
}

function createMultiSelectEditor(options, currentValue) {
  const wrapper = document.createElement("div");
  wrapper.className = "cell-editor multi-select-editor";
  wrapper.tabIndex = 0;

  const selected = new Set(parseEmployeesCell(currentValue));
  const uniqueOptions = Array.from(new Set(options)).filter(Boolean);

  const searchInput = document.createElement("input");
  searchInput.type = "text";
  searchInput.className = "multi-select-search";
  searchInput.placeholder = "Поиск сотрудника...";

  const listBox = document.createElement("div");
  listBox.className = "multi-select-options";

  const applyBtn = document.createElement("button");
  applyBtn.type = "button";
  applyBtn.className = "multi-select-apply-btn";
  applyBtn.textContent = "Готово";

  const renderOptions = (query = "") => {
    const normalized = query.trim().toLowerCase();
    const filtered = uniqueOptions.filter((name) => String(name).toLowerCase().includes(normalized));
    listBox.innerHTML = filtered.length
      ? filtered.map((name) => `
          <label class="multi-option-item">
            <input type="checkbox" value="${name}" ${selected.has(name) ? "checked" : ""} />
            <span>${name}</span>
          </label>
        `).join("")
      : '<div class="multi-option-empty">Ничего не найдено</div>';
  };

  listBox.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") return;
    if (target.checked) {
      selected.add(target.value);
    } else {
      selected.delete(target.value);
    }
    wrapper.dataset.value = Array.from(selected).join(", ");
  });

  searchInput.addEventListener("input", () => {
    renderOptions(searchInput.value);
  });

  wrapper.appendChild(searchInput);
  wrapper.appendChild(listBox);
  wrapper.appendChild(applyBtn);
  wrapper.dataset.value = Array.from(selected).join(", ");
  renderOptions();

  setTimeout(() => searchInput.focus(), 0);
  return wrapper;
}

function cleanupEditorResources(editor) {
  if (!(editor instanceof HTMLElement)) return;
  const listId = editor.dataset.datalistId;
  if (!listId) return;
  const datalist = document.getElementById(listId);
  if (datalist) {
    datalist.remove();
  }
  delete editor.dataset.datalistId;
}

function getEditorValue(editor) {
  if (editor.classList.contains("multi-select-editor")) {
    return editor.dataset.value || "";
  }
  if (editor instanceof HTMLInputElement && editor.type === "date") {
    return fromInputDate(editor.value);
  }
  return editor.value || "";
}

function toInputDate(ruDate) {
  if (!ruDate) return "";
  const [day, month, year] = ruDate.split(".");
  if (!day || !month || !year) return "";
  return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function fromInputDate(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  if (!year || !month || !day) return "";
  return `${day}.${month}.${year}`;
}

function attachHeaderActionHandlers(section, filteredEntries) {
  const refreshSectionBtn = document.getElementById("refreshSectionBtn");
  const googleSheetsSyncTasksBtn = document.getElementById("googleSheetsSyncTasksBtn");
  const sendOverdueTasksBtn = document.getElementById("sendOverdueTasksBtn");
  const openTaskImportModalBtn = document.getElementById("openTaskImportModalBtn");
  const addRowBtn = document.getElementById("addRowBtn");
  const toggleFiltersBtn = document.getElementById("toggleFiltersBtn");
  const toggleTableSettingsBtn = document.getElementById("toggleTableSettingsBtn");
  const openExportModalBtn = document.getElementById("openExportModalBtn");

  if (refreshSectionBtn) {
    refreshSectionBtn.addEventListener("click", () => {
      refreshCurrentViewData();
    });
  }

  if (googleSheetsSyncTasksBtn) {
    googleSheetsSyncTasksBtn.addEventListener("click", async () => {
      await triggerGoogleSheetsManualSync();
    });
  }
  if (sendOverdueTasksBtn && section.id === "tasks") {
    sendOverdueTasksBtn.addEventListener("click", async () => {
      if (sendOverdueTasksBtn.dataset.busy === "1") return;
      sendOverdueTasksBtn.dataset.busy = "1";
      sendOverdueTasksBtn.disabled = true;
      try {
        await triggerOverdueTaskManualNotifications();
      } finally {
        sendOverdueTasksBtn.disabled = false;
        sendOverdueTasksBtn.dataset.busy = "0";
      }
    });
  }

  if (openTaskImportModalBtn && section.id === "tasks") {
    openTaskImportModalBtn.addEventListener("click", () => {
      openTaskImportModal(section);
    });
  }

  if (addRowBtn) {
    addRowBtn.addEventListener("click", () => {
      addEmptyRow(section);
      saveSectionsData();
      renderTablePreserveScroll();
    });
  }

  if (toggleFiltersBtn) {
    toggleFiltersBtn.classList.toggle("is-active", filterPanelOpenBySection[section.id] === true);
    toggleFiltersBtn.addEventListener("click", () => {
      filterPanelOpenBySection[section.id] = !(filterPanelOpenBySection[section.id] === true);
      renderTablePreserveScroll();
    });
  }

  if (toggleTableSettingsBtn) {
    toggleTableSettingsBtn.addEventListener("click", () => {
      openTableSettingsModal(section);
    });
  }

  if (openExportModalBtn) {
    openExportModalBtn.addEventListener("click", () => {
      openExportFormatModal(section, filteredEntries);
    });
  }

  const statusTabButtons = Array.from(document.querySelectorAll(".status-tab-btn"));
  statusTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const tabValue = button.dataset.statusTab || "all";
      statusTabBySection[section.id] = tabValue;
      if (section.id === "tasks") resetTasksListPagingWindow();
      renderTablePreserveScroll();
    });
  });

  const sectionTabButtons = Array.from(document.querySelectorAll(".section-subtab-btn"));
  sectionTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const sectionId = String(button.dataset.sectionTab || "");
      if (!sectionId) return;
      selectSection(sectionId);
    });
  });

  Array.from(document.querySelectorAll('input[name="tasksQuickBrowseMode"]')).forEach((r) => {
    r.addEventListener("change", () => {
      displaySettings.tasksListBrowseMode = r.value === "byObject" ? "byObject" : "flat";
      if (displaySettings.tasksListBrowseMode === "flat") tasksBrowseObjectKey = null;
      saveDisplaySettings();
      resetTasksListPagingWindow();
      renderTablePreserveScroll();
    });
  });
}

function renderOtherSettingsPanel() {
  const allowedSettingsTabs = new Set(["general", "dateTime", "telegram", "googleSheets", "notifications", "taskFormat", "globalDup"]);
  const activeSettingsTab = allowedSettingsTabs.has(otherSettingsActiveTab) ? otherSettingsActiveTab : "general";
  const getStatusClass = (status) => {
    if (status === "Новый") return "status-legend-new";
    if (status === "В процессе") return "status-legend-progress";
    if (status === STATUS_DECISION) return "status-legend-decision";
    if (status === "Закрыт") return "status-legend-closed";
    return "";
  };

  const placeholderBarHtml = TASK_MESSAGE_PLACEHOLDERS_UI.map(
    (p) =>
      `<button type="button" class="task-placeholder-insert-btn" data-insert-token="${escapeHtmlText(p.token)}" title="${escapeHtmlText(p.label)}">${escapeHtmlText(p.token)}</button>`
  ).join("");

  const reminderPlaceholderBarHtml = TASK_MESSAGE_PLACEHOLDERS_UI.map(
    (p) =>
      `<button type="button" class="reminder-placeholder-insert-btn" data-insert-token="${escapeHtmlText(p.token)}" title="${escapeHtmlText(p.label)}">${escapeHtmlText(p.token)}</button>`
  ).join("");

  const taskFormatByStatus = displaySettings.taskMessageTemplatesByStatus || {};

  const curTz =
    displaySettings.serverTimezone === undefined || displaySettings.serverTimezone === null
      ? ""
      : String(displaySettings.serverTimezone);
  let tzOptionsList = [...SERVER_TIMEZONE_OPTIONS];
  if (curTz && !tzOptionsList.some((o) => o.id === curTz)) {
    tzOptionsList = [{ id: curTz, label: curTz }, ...tzOptionsList];
  }
  const tzOptsHtml = tzOptionsList
    .map((o) => {
      const sel = o.id === curTz ? " selected" : "";
      return `<option value="${escapeHtmlText(o.id)}"${sel}>${escapeHtmlText(o.label)}</option>`;
    })
    .join("");
  const df = normalizeDateDisplayFormatId(displaySettings.dateDisplayFormat);
  const dfOptsHtml = DATE_DISPLAY_FORMAT_OPTIONS.map(
    (o) => `<option value="${o.id}" ${o.id === df ? "selected" : ""}>${escapeHtmlText(o.label)}</option>`
  ).join("");
  const tf = normalizeTimeDisplayFormatId(displaySettings.timeDisplayFormat);
  const tfOptsHtml = TIME_DISPLAY_FORMAT_OPTIONS.map(
    (o) => `<option value="${o.id}" ${o.id === tf ? "selected" : ""}>${escapeHtmlText(o.label)}</option>`
  ).join("");
  const dupPositionOptsHtml = buildDupPositionFilterOptionsHtml();

  return `
    <section class="table-card">
      <div class="table-header">
        <h3>${withIcon("settings", "Прочие настройки")}</h3>
        <div class="table-header-right">
          <button type="button" class="icon-action-btn" id="otherSettingsRefreshBtn" title="Обновить">
            <i data-lucide="refresh-cw" class="lucide-icon" aria-hidden="true"></i>
          </button>
        </div>
      </div>
      <div class="other-settings-panel other-settings-panel--fill">
        <div class="other-settings-tabs" role="tablist" aria-label="Вкладки прочих настроек">
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "general" ? "active" : ""}" data-other-settings-tab="general">Основные</button>
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "dateTime" ? "active" : ""}" data-other-settings-tab="dateTime">Дата и время</button>
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "telegram" ? "active" : ""}" data-other-settings-tab="telegram">Telegram</button>
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "googleSheets" ? "active" : ""}" data-other-settings-tab="googleSheets">Google Sheets</button>
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "notifications" ? "active" : ""}" data-other-settings-tab="notifications">Настройки оповещения</button>
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "taskFormat" ? "active" : ""}" data-other-settings-tab="taskFormat">Шаблон сообщений</button>
          <button type="button" class="other-settings-tab-btn ${activeSettingsTab === "globalDup" ? "active" : ""}" data-other-settings-tab="globalDup">Получатели копий</button>
        </div>
        <div class="other-settings-section ${activeSettingsTab === "general" ? "" : "hidden"}" data-other-settings-pane="general">
          <h4 class="other-settings-section-title">Основные</h4>
          <div class="other-settings-block">
            <h4>Отображение строк</h4>
            <label class="settings-option">
              <input class="other-settings-checkbox" type="checkbox" data-setting="highlightClosed" ${displaySettings.highlightClosed ? "checked" : ""} />
              <span>Закрыт: светло-зеленый фон</span>
            </label>
          </div>
          <div class="other-settings-block other-settings-block--tasks-list">
            <h4>Список задач</h4>
            <div class="tasks-list-settings-compact">
              <div class="tasks-setting-block">
                <div class="tasks-setting-block-title">Экран</div>
                <div class="tasks-segment-group" role="radiogroup" aria-label="Режим экрана задач">
                  <label class="tasks-segment">
                    <input type="radio" name="tasksListBrowseMode" value="flat" ${displaySettings.tasksListBrowseMode !== "byObject" ? "checked" : ""} />
                    <span>Сводная</span>
                  </label>
                  <label class="tasks-segment">
                    <input type="radio" name="tasksListBrowseMode" value="byObject" ${displaySettings.tasksListBrowseMode === "byObject" ? "checked" : ""} />
                    <span>Объект</span>
                  </label>
                </div>
              </div>
              <div class="tasks-setting-block tasks-setting-block--row">
                <div class="tasks-setting-block-head">
                  <div class="tasks-setting-block-title">Строки</div>
                  <div class="tasks-segment-group" role="radiogroup" aria-label="Способ просмотра строк">
                    <label class="tasks-segment">
                      <input type="radio" name="tasksListPagingMode" value="pagination" ${displaySettings.tasksListPagingMode !== "chunks" ? "checked" : ""} />
                      <span>Пагинация</span>
                    </label>
                    <label class="tasks-segment">
                      <input type="radio" name="tasksListPagingMode" value="chunks" ${displaySettings.tasksListPagingMode === "chunks" ? "checked" : ""} />
                      <span>Показать ещё</span>
                    </label>
                  </div>
                </div>
                <div class="tasks-list-page-size-field">
                  <label class="tasks-page-size-label" for="tasksListPageSizeInput">За раз</label>
                  <input id="tasksListPageSizeInput" class="tasks-list-page-size-input" type="number" min="5" max="500" step="1" value="${displaySettings.tasksListPageSize ?? 50}" />
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="other-settings-section ${activeSettingsTab === "dateTime" ? "" : "hidden"}" data-other-settings-pane="dateTime">
          <h4 class="other-settings-section-title">Дата и время</h4>
          <div class="other-settings-block date-time-settings-block">
            <h4>Сервер и отображение</h4>
            <label class="settings-field-label" for="serverTimezoneSelect">Часовой пояс</label>
            <select id="serverTimezoneSelect" class="date-time-settings-select">${tzOptsHtml}</select>
            <label class="settings-field-label" for="dateDisplayFormatSelect">Формат даты в таблице</label>
            <select id="dateDisplayFormatSelect" class="date-time-settings-select">${dfOptsHtml}</select>
            <label class="settings-field-label" for="timeDisplayFormatSelect">Формат времени (дата удаления в корзине)</label>
            <select id="timeDisplayFormatSelect" class="date-time-settings-select">${tfOptsHtml}</select>
            <label class="settings-option settings-option--compact">
              <input type="checkbox" id="timeShowSecondsCheckbox" ${displaySettings.timeShowSeconds === true ? "checked" : ""} />
              <span>Показывать секунды во времени</span>
            </label>
          </div>
        </div>

        <div class="other-settings-section ${activeSettingsTab === "telegram" ? "" : "hidden"}" data-other-settings-pane="telegram">
          <h4 class="other-settings-section-title">Telegram</h4>
          <div class="other-settings-block">
            <h4>Бот</h4>
            <label class="settings-field-label" for="telegramBotTokenInput">Токен бота</label>
            <div class="token-field">
              <input id="telegramBotTokenInput" type="text" value="${escapeHtmlText(String(displaySettings.telegramBotToken || ""))}" placeholder="Введите токен Telegram бота" autocomplete="off" />
              <button type="button" id="copyTelegramTokenBtn" class="token-copy-btn" title="Скопировать токен" aria-label="Скопировать токен">
                <i data-lucide="copy" class="lucide-icon" aria-hidden="true"></i>
              </button>
            </div>
            <label class="settings-field-label" for="telegramBotUsernameReadonly">Ник бота в Telegram</label>
            <input id="telegramBotUsernameReadonly" type="text" class="other-settings-readonly-input" readonly tabindex="-1" aria-readonly="true" value="" />
            <label class="settings-field-label" for="telegramBotDisplayNameReadonly">Название бота</label>
            <input id="telegramBotDisplayNameReadonly" type="text" class="other-settings-readonly-input" readonly tabindex="-1" aria-readonly="true" value="" />
            <label class="settings-field-label" for="telegramAdminChatIdInput">Chat ID администратора</label>
            <input id="telegramAdminChatIdInput" type="text" inputmode="numeric" class="telegram-admin-chat-input" value="${escapeHtmlText(String(displaySettings.telegramAdminChatId || ""))}" placeholder="Ваш числовой Telegram ID" autocomplete="off" />
            <p class="other-settings-hint other-settings-hint--tight">Например из @userinfobot. Кнопка «Проверить бота» отправляет тест <strong>только</strong> в этот чат, а не всем сотрудникам с подключённым Telegram.</p>
            <div class="other-settings-actions">
              <button type="button" id="saveTelegramTokenBtn" class="secondary">Сохранить токен</button>
              <button type="button" id="testTelegramBotBtn">Проверить бота</button>
            </div>
            <p class="other-settings-hint">После «Сохранить токен» данные уходят на сервер и вызывается <strong>регистрация webhook</strong> — бот начинает принимать обновления на вашем домене. Сотрудник открывает бота и нажимает <strong>Старт</strong>: система сохраняет его реальный Telegram ID в колонку «Chat ID» и приветствует в чате.</p>
            <p class="other-settings-hint">Надёжная привязка: персональная ссылка <code>${escapeHtmlText(
              String(displaySettings.telegramBotUsername || "").trim()
                ? `https://t.me/${String(displaySettings.telegramBotUsername).trim()}?start=e_<ID>`
                : "https://t.me/<бот>?start=e_<ID>"
            )}</code>, где <strong>ID</strong> — значение из первой колонки сотрудника (например <code>e_3</code> в ссылке для ID 3). Если открыть бота без параметра, сопоставление идёт по <strong>имени и фамилии</strong> в профиле Telegram и ФИО в таблице (полное совпадение токенов имени).</p>
            <p class="other-settings-hint">Поле «Chat ID» не заполняется из номера телефона — только реальный Telegram user id после команды /start у бота (или вручную).</p>
          </div>
        </div>

        <div class="other-settings-section ${activeSettingsTab === "googleSheets" ? "" : "hidden"}" data-other-settings-pane="googleSheets">
          <h4 class="other-settings-section-title">Google Sheets</h4>
          <div class="other-settings-block">
            <h4>Синхронизация задач</h4>
            <label class="settings-option">
              <input type="checkbox" id="googleSheetsEnabledCheckbox" ${displaySettings.googleSheetsEnabled === true ? "checked" : ""} />
              <span>Включить интеграцию Google Sheets</span>
            </label>
            <label class="settings-option">
              <input type="checkbox" id="googleSheetsAutoSyncEnabledCheckbox" ${displaySettings.googleSheetsAutoSyncEnabled === true ? "checked" : ""} />
              <span>Автосинхронизация</span>
            </label>
            <label class="settings-field-label" for="googleSheetsSpreadsheetIdInput">Spreadsheet ID</label>
            <input id="googleSheetsSpreadsheetIdInput" type="text" value="${escapeHtmlAttr(String(displaySettings.googleSheetsSpreadsheetId || ""))}" placeholder="ID Google таблицы" autocomplete="off" />
            <label class="settings-field-label" for="googleSheetsSummarySheetNameInput">Лист «Сводная»</label>
            <input id="googleSheetsSummarySheetNameInput" type="text" value="${escapeHtmlAttr(String(displaySettings.googleSheetsSummarySheetName || "Сводная"))}" placeholder="Сводная" autocomplete="off" />
            <label class="settings-option">
              <input type="checkbox" id="googleSheetsIncludeObjectSheetsCheckbox" ${displaySettings.googleSheetsIncludeObjectSheets !== false ? "checked" : ""} />
              <span>Создавать отдельные листы по объектам</span>
            </label>
            <label class="settings-field-label" for="googleSheetsSyncIntervalInput">Интервал автосинхронизации (мин.)</label>
            <input id="googleSheetsSyncIntervalInput" type="number" min="1" max="1440" step="1" value="${normalizeGoogleSheetsInterval(displaySettings.googleSheetsSyncIntervalMinutes)}" />
            <div class="other-settings-actions">
              <button type="button" id="googleSheetsSyncNowBtn">Синхронизировать сейчас</button>
            </div>
            <p class="other-settings-hint">Ключи доступа хранятся только в Railway Variables: <code>GOOGLE_SHEETS_CLIENT_EMAIL</code> и <code>GOOGLE_SHEETS_PRIVATE_KEY</code>.</p>
            <label class="settings-field-label" for="googleSheetsSyncStatusText">Статус</label>
            <textarea id="googleSheetsSyncStatusText" class="other-settings-readonly-input" readonly rows="5">${escapeHtmlText(formatGoogleSheetsSyncStatusText())}</textarea>
          </div>
        </div>

        <div class="other-settings-section ${activeSettingsTab === "notifications" ? "" : "hidden"}" data-other-settings-pane="notifications">
          <h4 class="other-settings-section-title">Настройки оповещения</h4>
          <div class="other-settings-block">
            <h4>Просроченные задачи</h4>
            <label class="settings-option">
              <input type="checkbox" id="overdueNotificationsEnabledCheckbox" ${displaySettings.overdueNotificationsEnabled === true ? "checked" : ""} />
              <span>Отправлять уведомления о просроченных задачах</span>
            </label>
            <label class="settings-field-label" for="overdueNotificationsTimeInput">Время отправки</label>
            <div class="overdue-time-row">
              <input id="overdueNotificationsTimeInput" type="time" value="${escapeHtmlAttr(normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime))}" />
              <button type="button" id="overdueNotificationsTimeSaveBtn" class="secondary hidden">Сохранить</button>
            </div>
            <p class="other-settings-hint">Ежедневно в выбранное время система отправит уведомления по всем открытым просроченным задачам.</p>
          </div>
          <div class="settings-two-column settings-two-column--with-emulator">
            <div class="settings-two-column-main">
              <div class="other-settings-block">
                <h4>Напоминания по статусам задач</h4>
                <div class="task-placeholder-bar reminder-placeholder-bar" aria-label="Токены для напоминания">${reminderPlaceholderBarHtml}</div>
                <div class="reminder-status-grid">
                  ${STATUS_OPTIONS.map((status) => {
                    const conf = displaySettings.reminderSettings?.[status] || { days: "none", text: "" };
                    const statusClass = getStatusClass(status);
                    const cardMod = REMINDER_CARD_UI[status] || "new";
                    return `
                  <div class="reminder-item reminder-item--${cardMod}">
                    <div class="reminder-item-head">
                      <span class="reminder-status ${statusClass}">${status}</span>
                    </div>
                    <div class="reminder-item-body" data-reminder-body="${status}">
                      <label class="settings-field-label">Периодичность</label>
                      <select class="reminder-days-select" data-status="${status}">
                        ${REMINDER_DAYS_OPTIONS.map((opt) => `<option value="${opt.value}" ${String(conf.days) === opt.value ? "selected" : ""}>${opt.label}</option>`).join("")}
                      </select>
                      <div class="reminder-text-wrap" data-reminder-text-wrap="${status}">
                        <label class="settings-field-label">Текст напоминания</label>
                        <textarea class="reminder-text-input" data-status="${status}" rows="3" placeholder="Текст для сотрудника">${escapeHtmlText(String(conf.text || ""))}</textarea>
                      </div>
                    </div>
                  </div>
                `;
                  }).join("")}
                </div>
              </div>
            </div>
            <aside class="telegram-emulator" id="reminderTelegramEmulator" aria-label="Предпросмотр Telegram">
              <div class="telegram-emulator-heading">Предпросмотр</div>
              <div class="telegram-emulator-frame">
                <div class="telegram-emulator-notch" aria-hidden="true"></div>
                <div class="telegram-emulator-chat-bg">
                  <div id="reminderTelegramCaption" class="telegram-emulator-subtitle">Напоминание</div>
                  <div class="telegram-emulator-bubble-wrap">
                    <div class="telegram-emulator-bubble telegram-emulator-bubble--incoming">
                      <p id="reminderTelegramBubbleText" class="telegram-emulator-bubble-text">—</p>
                    </div>
                  </div>
                </div>
              </div>
              <p class="telegram-emulator-footnote">Тот же пример задачи №42, что и в «Шаблон сообщений».</p>
            </aside>
          </div>
        </div>

        <div class="other-settings-section global-dup-pane ${activeSettingsTab === "globalDup" ? "" : "hidden"}" data-other-settings-pane="globalDup">
          <h4 class="other-settings-section-title">Получатели копий (Telegram)</h4>
          <p id="globalDupRecipientSummary" class="global-dup-help-summary global-dup-help-summary--under-title"></p>
          <div class="dup-recipient-filters dup-recipient-filters--row">
            <div class="dup-recipient-filter-field dup-recipient-filter-field--search">
              <label class="settings-field-label" for="dupRecipientSearchInput">Поиск</label>
              <input type="search" id="dupRecipientSearchInput" class="dup-recipient-search" placeholder="ФИО или должность" autocomplete="off" />
            </div>
            <div class="dup-recipient-filter-field dup-recipient-filter-field--position">
              <label class="settings-field-label" for="dupRecipientPositionFilter">Должность</label>
              <select id="dupRecipientPositionFilter" class="dup-recipient-position-filter">${dupPositionOptsHtml}</select>
            </div>
          </div>
          <div class="dup-shuttle dup-shuttle--embedded global-dup-shuttle">
            <div class="dup-shuttle-col">
              <div class="dup-shuttle-head">
                <div class="dup-shuttle-head-row">
                  <span class="dup-shuttle-title-line">Сотрудники</span>
                </div>
              </div>
              <div class="dup-shuttle-grid-header">
                <span class="dup-shuttle-hcell dup-shuttle-hcell--master-cb">
                  <label class="dup-shuttle-select-all" title="Выбрать все">
                    <input type="checkbox" id="dupRecipientSelectAllLeft" aria-label="Выбрать все в списке" />
                  </label>
                </span>
                <span class="dup-shuttle-hcell">ФИО сотрудника</span>
                <span class="dup-shuttle-hcell">Должность</span>
              </div>
              <div class="dup-shuttle-scroll" id="dupRecipientLeftScroll"></div>
            </div>
            <div class="dup-shuttle-mid">
              <button type="button" class="secondary" id="dupRecipientAddBtn">Добавить →</button>
              <button type="button" class="secondary" id="dupRecipientRemoveBtn">← Убрать</button>
            </div>
            <div class="dup-shuttle-col">
              <div class="dup-shuttle-head">
                <div class="dup-shuttle-head-row">
                  <span class="dup-shuttle-title-line">Получают копии по всем задачам</span>
                </div>
              </div>
              <div class="dup-shuttle-grid-header dup-shuttle-grid-header--right">
                <span class="dup-shuttle-hcell dup-shuttle-hcell--master-cb">
                  <label class="dup-shuttle-select-all" title="Выбрать все">
                    <input type="checkbox" id="dupRecipientSelectAllRight" aria-label="Выбрать все в списке" />
                  </label>
                </span>
                <span class="dup-shuttle-hcell">ФИО сотрудника</span>
                <span class="dup-shuttle-hcell">Должность</span>
                <span class="dup-shuttle-hcell dup-shuttle-hcell--confirm" title="Подтверждение закрытия задачи в Telegram">Подтв.</span>
              </div>
              <div class="dup-shuttle-scroll" id="dupRecipientRightScroll"></div>
            </div>
          </div>
        </div>

        <div class="other-settings-section ${activeSettingsTab === "taskFormat" ? "" : "hidden"}" data-other-settings-pane="taskFormat">
          <h4 class="other-settings-section-title">Шаблон сообщений</h4>
          <div class="settings-two-column settings-two-column--with-emulator">
            <div class="settings-two-column-main">
              <div class="other-settings-block">
                <h4>Подсказки (поля таблицы задач)</h4>
                <div class="task-placeholder-bar" aria-label="Подсказки для вставки">${placeholderBarHtml}</div>
                <div class="task-format-status-list">
            ${STATUS_OPTIONS.map((status, idx) => {
              const statusClass = getStatusClass(status);
              const cardMod = REMINDER_CARD_UI[status] || "new";
              const tpl = String(taskFormatByStatus[status] ?? "");
              const fieldId = `taskMsgTpl_${idx}`;
              return `
                <div class="other-settings-block task-format-block task-format-block--${cardMod}">
                  <div class="task-format-status-head"><span class="reminder-status ${statusClass}">${status}</span></div>
                  <label class="settings-field-label" for="${fieldId}">Текст для отправки при этом статусе</label>
                  <textarea id="${fieldId}" class="task-message-template-input" data-status="${status}" rows="5" placeholder="Например: Задача [ид_задачи]: [название_задачи], срок [срок_задачи]">${escapeHtmlText(tpl)}</textarea>
                </div>
              `;
            }).join("")}
                </div>
              </div>
              <div class="other-settings-block task-format-close-confirm-block">
                <label class="settings-field-label" for="telegramCloseAcceptedInput">Текст после подтверждения закрытия в Telegram</label>
                <textarea id="telegramCloseAcceptedInput" class="telegram-close-accepted-input" rows="3" placeholder="Задача [ид_задачи]: …">${escapeHtmlText(
                  String(displaySettings.telegramCloseAcceptedTemplate || "")
                )}</textarea>
              </div>
            </div>
            <aside class="telegram-emulator" id="taskMsgTelegramEmulator" aria-label="Предпросмотр Telegram">
              <div class="telegram-emulator-heading">Предпросмотр</div>
              <div class="telegram-emulator-frame">
                <div class="telegram-emulator-notch" aria-hidden="true"></div>
                <div class="telegram-emulator-chat-bg">
                  <div id="taskMsgTelegramCaption" class="telegram-emulator-subtitle">Сообщение по задаче</div>
                  <div class="telegram-emulator-bubble-wrap">
                    <div class="telegram-emulator-bubble telegram-emulator-bubble--incoming">
                      <p id="taskMsgTelegramBubbleText" class="telegram-emulator-bubble-text">—</p>
                    </div>
                  </div>
                </div>
                <div id="taskMsgTelegramKeyboard" class="telegram-emulator-keyboard">
                  <div class="telegram-emulator-keyboard-row">
                    <span class="telegram-emulator-btn">⌛️ Сменить статус</span>
                    <span class="telegram-emulator-btn">🗣 Комментарий</span>
                  </div>
                  <div class="telegram-emulator-keyboard-row">
                    <span class="telegram-emulator-btn telegram-emulator-btn--wide">📸 Отправить фото</span>
                  </div>
                </div>
              </div>
              <p class="telegram-emulator-footnote">Пример подстановок: задача №42, тот же набор полей для всех статусов.</p>
            </aside>
          </div>
        </div>
      </div>
    </section>
  `;
}

function attachOtherSettingsHandlers() {
  document.getElementById("otherSettingsRefreshBtn")?.addEventListener("click", () => {
    refreshCurrentViewData();
  });
  const tabButtons = Array.from(document.querySelectorAll(".other-settings-tab-btn"));
  const tabPanes = Array.from(document.querySelectorAll("[data-other-settings-pane]"));
  tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const tabId = String(button.dataset.otherSettingsTab || "");
      if (!tabId) return;
      otherSettingsActiveTab = tabId;
      tabButtons.forEach((btn) => btn.classList.toggle("active", btn === button));
      tabPanes.forEach((pane) => {
        pane.classList.toggle("hidden", pane.dataset.otherSettingsPane !== tabId);
      });
      if (tabId === "globalDup") {
        renderGlobalDupShuttleIntoDom();
      }
      if (tabId === "taskFormat" && typeof window._mbcRefreshTaskFormatPreview === "function") {
        window._mbcRefreshTaskFormatPreview();
      }
      if (tabId === "notifications" && typeof window._mbcRefreshReminderPreview === "function") {
        window._mbcRefreshReminderPreview();
      }
    });
  });

  let lastTaskTemplateTextarea = null;
  const templateInputs = Array.from(document.querySelectorAll(".task-message-template-input"));
  templateInputs.forEach((ta) => {
    const mark = () => {
      lastTaskTemplateTextarea = ta;
    };
    ta.addEventListener("focus", mark);
    ta.addEventListener("click", mark);
  });
  const commitTaskTemplate = (input) => {
    const status = String(input.dataset.status || "");
    if (!status) return;
    if (!displaySettings.taskMessageTemplatesByStatus) {
      displaySettings.taskMessageTemplatesByStatus = Object.fromEntries(STATUS_OPTIONS.map((s) => [s, ""]));
    }
    displaySettings.taskMessageTemplatesByStatus[status] = String(input.value || "");
    saveDisplaySettings();
  };
  templateInputs.forEach((input) => {
    input.addEventListener("blur", () => commitTaskTemplate(input));
    input.addEventListener("change", () => commitTaskTemplate(input));
  });
  Array.from(document.querySelectorAll(".task-placeholder-insert-btn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const token = String(btn.getAttribute("data-insert-token") || "");
      let ta = lastTaskTemplateTextarea;
      if (!ta) ta = document.querySelector(".task-message-template-input");
      if (!ta || !token) return;
      const start = typeof ta.selectionStart === "number" ? ta.selectionStart : ta.value.length;
      const end = typeof ta.selectionEnd === "number" ? ta.selectionEnd : ta.value.length;
      const v = ta.value;
      ta.value = `${v.slice(0, start)}${token}${v.slice(end)}`;
      const pos = start + token.length;
      ta.focus();
      if (typeof ta.setSelectionRange === "function") {
        ta.setSelectionRange(pos, pos);
      }
      commitTaskTemplate(ta);
      if (typeof window._mbcRefreshTaskFormatPreview === "function") {
        window._mbcRefreshTaskFormatPreview();
      }
    });
  });

  const serverTzEl = document.getElementById("serverTimezoneSelect");
  const dateFmtEl = document.getElementById("dateDisplayFormatSelect");
  const timeFmtEl = document.getElementById("timeDisplayFormatSelect");
  const timeSecEl = document.getElementById("timeShowSecondsCheckbox");
  const commitServerDateTimeSettings = () => {
    if (serverTzEl) displaySettings.serverTimezone = String(serverTzEl.value || "");
    if (dateFmtEl) displaySettings.dateDisplayFormat = normalizeDateDisplayFormatId(dateFmtEl.value);
    if (timeFmtEl) displaySettings.timeDisplayFormat = normalizeTimeDisplayFormatId(timeFmtEl.value);
    if (timeSecEl) displaySettings.timeShowSeconds = Boolean(timeSecEl.checked);
    saveDisplaySettings();
    renderTablePreserveScroll();
  };
  serverTzEl?.addEventListener("change", commitServerDateTimeSettings);
  dateFmtEl?.addEventListener("change", commitServerDateTimeSettings);
  timeFmtEl?.addEventListener("change", commitServerDateTimeSettings);
  timeSecEl?.addEventListener("change", commitServerDateTimeSettings);

  const overdueEnabledEl = document.getElementById("overdueNotificationsEnabledCheckbox");
  const overdueTimeEl = document.getElementById("overdueNotificationsTimeInput");
  const overdueTimeSaveBtn = document.getElementById("overdueNotificationsTimeSaveBtn");
  const updateOverdueSaveBtnState = () => {
    if (!overdueTimeEl || !overdueTimeSaveBtn) return;
    const current = normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime);
    const draft = normalizeOverdueNotifyTimeValue(overdueTimeEl.value);
    overdueTimeSaveBtn.classList.toggle("hidden", current === draft);
  };
  const commitOverdueNotificationsSettings = () => {
    const prevEnabled = Boolean(displaySettings.overdueNotificationsEnabled);
    const prevTime = normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime);
    if (overdueEnabledEl) displaySettings.overdueNotificationsEnabled = Boolean(overdueEnabledEl.checked);
    if (overdueTimeEl) displaySettings.overdueNotificationsTime = normalizeOverdueNotifyTimeValue(overdueTimeEl.value);
    if (overdueTimeEl) overdueTimeEl.value = normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime);
    const nextTime = normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime);
    if (prevEnabled !== displaySettings.overdueNotificationsEnabled || prevTime !== nextTime) {
      saveOverdueNotifyRuntime({ lastRunDate: "" });
    }
    saveDisplaySettings();
    startOverdueTaskNotificationsScheduler();
    runOverdueTaskNotificationTick().catch(() => {});
    updateOverdueSaveBtnState();
  };
  overdueEnabledEl?.addEventListener("change", commitOverdueNotificationsSettings);
  overdueTimeEl?.addEventListener("input", updateOverdueSaveBtnState);
  overdueTimeEl?.addEventListener("change", updateOverdueSaveBtnState);
  overdueTimeSaveBtn?.addEventListener("click", commitOverdueNotificationsSettings);
  updateOverdueSaveBtnState();

  const checkboxes = Array.from(document.querySelectorAll(".other-settings-checkbox"));
  checkboxes.forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const settingName = checkbox.dataset.setting;
      if (!settingName) return;
      displaySettings[settingName] = checkbox.checked;
      saveDisplaySettings();
    });
  });

  Array.from(document.querySelectorAll('input[name="tasksListBrowseMode"]')).forEach((r) => {
    r.addEventListener("change", () => {
      displaySettings.tasksListBrowseMode = r.value === "byObject" ? "byObject" : "flat";
      if (displaySettings.tasksListBrowseMode === "flat") tasksBrowseObjectKey = null;
      saveDisplaySettings();
      resetTasksListPagingWindow();
      renderTablePreserveScroll();
    });
  });
  Array.from(document.querySelectorAll('input[name="tasksListPagingMode"]')).forEach((r) => {
    r.addEventListener("change", () => {
      displaySettings.tasksListPagingMode = r.value === "chunks" ? "chunks" : "pagination";
      saveDisplaySettings();
      resetTasksListPagingWindow();
      renderTablePreserveScroll();
    });
  });
  const tasksPageSizeInput = document.getElementById("tasksListPageSizeInput");
  const commitTasksPageSize = () => {
    if (!tasksPageSizeInput) return;
    let n = Number(tasksPageSizeInput.value);
    if (!Number.isFinite(n)) n = TASKS_DEFAULT_LIST_PAGE_SIZE;
    displaySettings.tasksListPageSize = Math.min(500, Math.max(5, Math.floor(n)));
    tasksPageSizeInput.value = String(displaySettings.tasksListPageSize);
    saveDisplaySettings();
    resetTasksListPagingWindow();
    renderTablePreserveScroll();
  };
  tasksPageSizeInput?.addEventListener("change", commitTasksPageSize);
  tasksPageSizeInput?.addEventListener("blur", commitTasksPageSize);

  const tokenInput = document.getElementById("telegramBotTokenInput");
  const copyTokenBtn = document.getElementById("copyTelegramTokenBtn");
  const saveButton = document.getElementById("saveTelegramTokenBtn");
  const testButton = document.getElementById("testTelegramBotBtn");
  const persistTokenLocal = () => {
    if (!tokenInput) return;
    displaySettings.telegramBotToken = String(tokenInput.value || "").trim();
    saveDisplaySettings();
  };
  saveButton?.addEventListener("click", async () => {
    const reg = await flushTelegramBotTokenToServer({ silent: true });
    if (reg.ok) {
      const dn = String(displaySettings.telegramBotDisplayName || "").trim();
      window.alert(
        `Токен и данные обновлены на сервере.${reg.webhookUrl ? `\n\nWebhook: ${reg.webhookUrl}` : ""}${displaySettings.telegramBotUsername ? `\nБот: @${String(displaySettings.telegramBotUsername).trim()}` : ""}${dn ? `\nНазвание в Telegram: ${dn}` : ""}`
      );
    } else {
      window.alert(
        `Не удалось полностью обновить сервер:\n${reg.error || "ошибка"}\n\nПроверьте вход в систему на Railway и при необходимости PUBLIC_APP_URL. Локальная копия токена в браузере сохранена.`
      );
    }
  });
  tokenInput?.addEventListener("blur", async () => {
    if (!tokenInput) return;
    const next = String(tokenInput.value || "").trim();
    if (next === String(displaySettings.telegramBotToken || "").trim()) {
      return;
    }
    await flushTelegramBotTokenToServer({ silent: true });
  });

  let telegramTokenProfileTimer = null;
  tokenInput?.addEventListener("input", () => {
    clearTimeout(telegramTokenProfileTimer);
    telegramTokenProfileTimer = setTimeout(async () => {
      if (!tokenInput) return;
      const t = String(tokenInput.value || "").trim();
      if (!t || t.length < 40 || !/^\d+:[A-Za-z0-9_-]+$/.test(t)) return;
      await refreshTelegramBotProfileFromToken(t);
      saveDisplaySettings();
    }, 450);
  });

  const adminChatInput = document.getElementById("telegramAdminChatIdInput");
  const commitAdminChatId = () => {
    if (!adminChatInput) return;
    displaySettings.telegramAdminChatId = String(adminChatInput.value || "").trim();
    saveDisplaySettings();
  };
  adminChatInput?.addEventListener("change", commitAdminChatId);
  adminChatInput?.addEventListener("blur", commitAdminChatId);

  const closeAcceptedInput = document.getElementById("telegramCloseAcceptedInput");
  const commitCloseAccepted = () => {
    if (!closeAcceptedInput) return;
    displaySettings.telegramCloseAcceptedTemplate = String(closeAcceptedInput.value || "").trim();
    saveDisplaySettings();
  };
  closeAcceptedInput?.addEventListener("blur", () => {
    commitCloseAccepted();
    if (typeof window._mbcRefreshTaskFormatPreview === "function") {
      window._mbcRefreshTaskFormatPreview();
    }
  });
  closeAcceptedInput?.addEventListener("change", () => {
    commitCloseAccepted();
    if (typeof window._mbcRefreshTaskFormatPreview === "function") {
      window._mbcRefreshTaskFormatPreview();
    }
  });

  const gsEnabledEl = document.getElementById("googleSheetsEnabledCheckbox");
  const gsAutoEl = document.getElementById("googleSheetsAutoSyncEnabledCheckbox");
  const gsSpreadsheetEl = document.getElementById("googleSheetsSpreadsheetIdInput");
  const gsSummaryEl = document.getElementById("googleSheetsSummarySheetNameInput");
  const gsIncludeObjEl = document.getElementById("googleSheetsIncludeObjectSheetsCheckbox");
  const gsIntervalEl = document.getElementById("googleSheetsSyncIntervalInput");
  const gsSyncBtn = document.getElementById("googleSheetsSyncNowBtn");

  const commitGoogleSheetsSettings = () => {
    if (gsEnabledEl) displaySettings.googleSheetsEnabled = Boolean(gsEnabledEl.checked);
    if (gsAutoEl) displaySettings.googleSheetsAutoSyncEnabled = Boolean(gsAutoEl.checked);
    if (gsSpreadsheetEl) displaySettings.googleSheetsSpreadsheetId = String(gsSpreadsheetEl.value || "").trim();
    if (gsSummaryEl) displaySettings.googleSheetsSummarySheetName = String(gsSummaryEl.value || "").trim() || "Сводная";
    if (gsIncludeObjEl) displaySettings.googleSheetsIncludeObjectSheets = Boolean(gsIncludeObjEl.checked);
    if (gsIntervalEl) {
      displaySettings.googleSheetsSyncIntervalMinutes = normalizeGoogleSheetsInterval(gsIntervalEl.value);
      gsIntervalEl.value = String(displaySettings.googleSheetsSyncIntervalMinutes);
    }
    saveDisplaySettings();
  };
  gsEnabledEl?.addEventListener("change", commitGoogleSheetsSettings);
  gsAutoEl?.addEventListener("change", commitGoogleSheetsSettings);
  gsSpreadsheetEl?.addEventListener("blur", commitGoogleSheetsSettings);
  gsSummaryEl?.addEventListener("blur", commitGoogleSheetsSettings);
  gsIncludeObjEl?.addEventListener("change", commitGoogleSheetsSettings);
  gsIntervalEl?.addEventListener("change", commitGoogleSheetsSettings);
  gsIntervalEl?.addEventListener("blur", commitGoogleSheetsSettings);

  gsSyncBtn?.addEventListener("click", async () => {
    commitGoogleSheetsSettings();
    const prevLabel = gsSyncBtn.textContent;
    gsSyncBtn.disabled = true;
    gsSyncBtn.textContent = "Синхронизация...";
    try {
      await triggerGoogleSheetsManualSync();
    } finally {
      gsSyncBtn.disabled = false;
      gsSyncBtn.textContent = prevLabel || "Синхронизировать сейчас";
    }
  });

  attachGlobalDuplicateRecipientsHandlers();
  attachTelegramTemplatePreviews();

  const reminderDaySelectors = Array.from(document.querySelectorAll(".reminder-days-select"));
  reminderDaySelectors.forEach((select) => {
    select.addEventListener("change", () => {
      const status = String(select.dataset.status || "");
      if (!status) return;
      if (!displaySettings.reminderSettings) {
        displaySettings.reminderSettings = {};
      }
      const current = displaySettings.reminderSettings[status] || { days: "none", text: "" };
      displaySettings.reminderSettings[status] = {
        ...current,
        days: String(select.value || "none")
      };
      saveDisplaySettings();
    });
  });

  let lastReminderTextInput = null;
  const reminderTextInputs = Array.from(document.querySelectorAll(".reminder-text-input"));
  reminderTextInputs.forEach((input) => {
    const markLast = () => {
      lastReminderTextInput = input;
    };
    input.addEventListener("focus", markLast);
    input.addEventListener("click", markLast);
    const commit = () => {
      const status = String(input.dataset.status || "");
      if (!status) return;
      if (!displaySettings.reminderSettings) {
        displaySettings.reminderSettings = {};
      }
      const current = displaySettings.reminderSettings[status] || { days: "none", text: "" };
      displaySettings.reminderSettings[status] = {
        ...current,
        text: String(input.value || "").trim()
      };
      saveDisplaySettings();
    };
    input.addEventListener("blur", commit);
    input.addEventListener("change", commit);
  });
  Array.from(document.querySelectorAll(".reminder-placeholder-insert-btn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const token = String(btn.getAttribute("data-insert-token") || "");
      let el = lastReminderTextInput;
      if (!el) el = document.querySelector(".reminder-text-input");
      if (!el || !token) return;
      const start = typeof el.selectionStart === "number" ? el.selectionStart : el.value.length;
      const end = typeof el.selectionEnd === "number" ? el.selectionEnd : el.value.length;
      const v = el.value;
      el.value = `${v.slice(0, start)}${token}${v.slice(end)}`;
      const pos = start + token.length;
      el.focus();
      if (typeof el.setSelectionRange === "function") {
        el.setSelectionRange(pos, pos);
      }
      const status = String(el.dataset.status || "");
      if (!status) return;
      if (!displaySettings.reminderSettings) {
        displaySettings.reminderSettings = {};
      }
      const current = displaySettings.reminderSettings[status] || { days: "none", text: "" };
      displaySettings.reminderSettings[status] = {
        ...current,
        text: String(el.value || "").trim()
      };
      saveDisplaySettings();
      if (typeof window._mbcRefreshReminderPreview === "function") {
        window._mbcRefreshReminderPreview();
      }
    });
  });

  copyTokenBtn?.addEventListener("click", async () => {
    if (!tokenInput) return;
    const value = String(tokenInput.value || "").trim();
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      copyTokenBtn.title = "Скопировано";
      setTimeout(() => {
        copyTokenBtn.title = "Скопировать токен";
      }, 1200);
    } catch (_) {
      tokenInput.select();
      document.execCommand("copy");
    }
  });
  testButton?.addEventListener("click", async () => {
    const flush = await flushTelegramBotTokenToServer({ silent: true });
    const token = String(displaySettings.telegramBotToken || "").trim();
    if (!token) {
      window.alert("Сначала укажите токен Telegram-бота.");
      return;
    }

    const adminChatRaw = String(displaySettings.telegramAdminChatId || "").trim();
    if (!adminChatRaw) {
      window.alert(
        "Укажите Chat ID администратора в поле выше. Проверочное сообщение отправляется только туда, а не сотрудникам из справочника."
      );
      return;
    }
    if (!/^-?\d+$/.test(adminChatRaw)) {
      window.alert("Chat ID должен быть целым числом (например ID из @userinfobot). Для супергрупп допускается отрицательное значение.");
      return;
    }
    const prof = await refreshTelegramBotProfileFromToken(token);
    if (!prof.ok && prof.soft) {
      window.alert("Не удалось связаться с api.telegram.org. Проверьте сеть и блокировки.");
      return;
    }
    if (!prof.ok) {
      const desc = String(prof.description || "").trim() || "Проверьте токен в @BotFather.";
      window.alert(`Токен отклонён Telegram API: ${desc}`);
      return;
    }
    saveDisplaySettings();

    const botUser = displaySettings.telegramBotUsername ? `@${String(displaySettings.telegramBotUsername).trim()}` : "";
    const webhookLine =
      flush.ok && flush.webhookUrl
        ? `\n\nWebhook: ${flush.webhookUrl}`
        : !flush.ok && flush.error
          ? `\n\nСервер/webhook: ${flush.error}`
          : "";

    const message = "Бот работает, проблем нет";
    let sendOk = false;
    let sendDesc = "";
    try {
      const response = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          chat_id: adminChatRaw,
          text: message
        })
      });
      const data = await response.json().catch(() => ({}));
      sendOk = response.ok && data.ok === true;
      sendDesc = String(data.description || "").trim();
    } catch (_) {
      sendOk = false;
    }

    if (sendOk) {
      window.alert(
        `Токен верный${botUser ? `, бот ${botUser}` : ""}.${webhookLine}\n\nТестовое сообщение отправлено в чат администратора (Chat ID ${adminChatRaw}).`
      );
      return;
    }
    window.alert(
      sendDesc
        ? `Токен верный${botUser ? ` (${botUser})` : ""}, но отправка администратору не прошла:\n${sendDesc}${webhookLine}\n\nПроверьте Chat ID и что вы уже писали боту / нажали «Старт» в его чате.`
        : `Токен верный${botUser ? ` (${botUser})` : ""}, но отправка не удалась.${webhookLine}`
    );
  });

  updateTelegramBotProfileReadonlyDom();
  initLucideIcons();
}

function saveDisplaySettings(opts = {}) {
  localStorage.setItem(DISPLAY_SETTINGS_KEY, JSON.stringify(displaySettings));
  if (!opts.skipServerSync) {
    scheduleServerSync();
  }
}

function restoreDisplaySettings() {
  const raw = localStorage.getItem(DISPLAY_SETTINGS_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    const defaultReminderSettings = Object.fromEntries(
      STATUS_OPTIONS.map((status) => [status, { days: "none", text: "" }])
    );
    const defaultTaskTemplates = Object.fromEntries(STATUS_OPTIONS.map((status) => [status, ""]));
    const parsedReminderSettings = parsed && typeof parsed === "object" ? parsed.reminderSettings : null;
    const parsedTaskTemplates =
      parsed && typeof parsed === "object" && parsed.taskMessageTemplatesByStatus && typeof parsed.taskMessageTemplatesByStatus === "object"
        ? parsed.taskMessageTemplatesByStatus
        : null;
    const remapStatusKeyObject = (obj) => {
      if (!obj || typeof obj !== "object") return {};
      const out = {};
      Object.entries(obj).forEach(([k, v]) => {
        const nk = normalizeTaskStatusValue(k);
        out[nk] = v;
      });
      return out;
    };
    displaySettings = {
      ...displaySettings,
      ...parsed,
      reminderSettings: {
        ...defaultReminderSettings,
        ...remapStatusKeyObject(parsedReminderSettings)
      },
      taskMessageTemplatesByStatus: {
        ...defaultTaskTemplates,
        ...remapStatusKeyObject(parsedTaskTemplates)
      }
    };
    if (typeof displaySettings.serverTimezone !== "string") {
      displaySettings.serverTimezone = "";
    }
    displaySettings.dateDisplayFormat = normalizeDateDisplayFormatId(displaySettings.dateDisplayFormat);
    displaySettings.timeDisplayFormat = normalizeTimeDisplayFormatId(displaySettings.timeDisplayFormat);
    displaySettings.timeShowSeconds = Boolean(displaySettings.timeShowSeconds);
    displaySettings.overdueNotificationsEnabled = Boolean(displaySettings.overdueNotificationsEnabled);
    displaySettings.overdueNotificationsTime = normalizeOverdueNotifyTimeValue(displaySettings.overdueNotificationsTime);
    if (!Array.isArray(displaySettings.telegramGlobalDuplicateRecipientIds)) {
      displaySettings.telegramGlobalDuplicateRecipientIds = [];
    }
    if (!Array.isArray(displaySettings.telegramCloseConfirmAllowedIds)) {
      displaySettings.telegramCloseConfirmAllowedIds = [];
    } else {
      const dupSet = new Set(displaySettings.telegramGlobalDuplicateRecipientIds.map((x) => String(x).trim()));
      displaySettings.telegramCloseConfirmAllowedIds = displaySettings.telegramCloseConfirmAllowedIds
        .map((x) => String(x).trim())
        .filter((id) => dupSet.has(id));
    }
    if (typeof displaySettings.telegramCloseAcceptedTemplate !== "string") {
      displaySettings.telegramCloseAcceptedTemplate =
        "Задача [ид_задачи] ([название_задачи]): закрытие подтверждено.";
    }
    if (typeof displaySettings.telegramBotUsername !== "string") {
      displaySettings.telegramBotUsername = "";
    }
    if (typeof displaySettings.telegramBotDisplayName !== "string") {
      displaySettings.telegramBotDisplayName = "";
    }
    if (typeof displaySettings.telegramAdminChatId !== "string") {
      displaySettings.telegramAdminChatId = "";
    }
    if (!Array.isArray(displaySettings.pendingImportedTaskIds)) {
      displaySettings.pendingImportedTaskIds = [];
    } else {
      displaySettings.pendingImportedTaskIds = displaySettings.pendingImportedTaskIds
        .map((x) => String(x || "").trim())
        .filter(Boolean);
    }
    displaySettings.googleSheetsEnabled = Boolean(displaySettings.googleSheetsEnabled);
    displaySettings.googleSheetsAutoSyncEnabled = Boolean(displaySettings.googleSheetsAutoSyncEnabled);
    displaySettings.googleSheetsIncludeObjectSheets = displaySettings.googleSheetsIncludeObjectSheets !== false;
    if (typeof displaySettings.googleSheetsSpreadsheetId !== "string") {
      displaySettings.googleSheetsSpreadsheetId = "";
    }
    if (typeof displaySettings.googleSheetsSummarySheetName !== "string" || !displaySettings.googleSheetsSummarySheetName.trim()) {
      displaySettings.googleSheetsSummarySheetName = "Сводная";
    }
    displaySettings.googleSheetsSyncIntervalMinutes = normalizeGoogleSheetsInterval(displaySettings.googleSheetsSyncIntervalMinutes);
    if (typeof displaySettings.googleSheetsLastSyncStatus !== "string") displaySettings.googleSheetsLastSyncStatus = "";
    if (typeof displaySettings.googleSheetsLastSyncAt !== "string") displaySettings.googleSheetsLastSyncAt = "";
    if (typeof displaySettings.googleSheetsLastSyncMessage !== "string") displaySettings.googleSheetsLastSyncMessage = "";
    if (typeof displaySettings.googleSheetsLastSyncMode !== "string") displaySettings.googleSheetsLastSyncMode = "";
    displaySettings.googleSheetsLastSyncAtMs = Number(displaySettings.googleSheetsLastSyncAtMs) || 0;
    displaySettings.googleSheetsLastSyncRows = Number(displaySettings.googleSheetsLastSyncRows) || 0;
    if (displaySettings.tasksListPagingMode !== "pagination" && displaySettings.tasksListPagingMode !== "chunks") {
      displaySettings.tasksListPagingMode = "pagination";
    }
    if (displaySettings.tasksListBrowseMode !== "flat" && displaySettings.tasksListBrowseMode !== "byObject") {
      displaySettings.tasksListBrowseMode = "flat";
    }
    let tps = Number(displaySettings.tasksListPageSize);
    if (!Number.isFinite(tps)) tps = 50;
    displaySettings.tasksListPageSize = Math.min(500, Math.max(5, Math.floor(tps)));
  } catch (_) {
    // ignore broken storage and keep defaults
  }
}

function renderTablePreserveScroll() {
  startTurboLoader();
  const tableWrap = document.querySelector(".table-wrap");
  const scrollTop = tableWrap ? tableWrap.scrollTop : 0;
  const scrollLeft = tableWrap ? tableWrap.scrollLeft : 0;
  renderTable();
  const nextWrap = document.querySelector(".table-wrap");
  if (nextWrap) {
    nextWrap.scrollTop = scrollTop;
    nextWrap.scrollLeft = scrollLeft;
  }
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      finishTurboLoader();
    });
  });
}

function markFocusedRow(cell) {
  const table = cell.closest("table");
  if (!table) return;
  table.querySelectorAll("tr.focused-row").forEach((row) => row.classList.remove("focused-row"));
  const row = cell.closest("tr");
  row?.classList.add("focused-row");
}

function clearActiveRow(sectionId) {
  delete activeRowBySection[sectionId];
}

function openTaskDetailsModal(section, row, rowIndex) {
  const columns = section.columns;
  const taskId = String(row[TASK_COLUMNS.number] || "");
  const draftState = {
    before: getMediaItems(row[TASK_COLUMNS.mediaBefore]),
    after: getMediaItems(row[TASK_COLUMNS.mediaAfter]),
    preview: {}
  };
  let isDirty = false;

  draftState.before.forEach((name, idx) => {
    const key = `before-${idx}`;
    const source = mediaPreviewStore[getMediaSlotKey(taskId, TASK_COLUMNS.mediaBefore, idx)];
    draftState.preview[key] = source ? { ...source } : { name, type: "", url: "" };
  });
  draftState.after.forEach((name, idx) => {
    const key = `after-${idx}`;
    const source = mediaPreviewStore[getMediaSlotKey(taskId, TASK_COLUMNS.mediaAfter, idx)];
    draftState.preview[key] = source ? { ...source } : { name, type: "", url: "" };
  });
  const details = columns
    .map((column, index) => {
      if (index === TASK_COLUMNS.mediaBefore || index === TASK_COLUMNS.mediaAfter) return "";
      const value = row[index] || "-";
      return `<div class="detail-item"><span class="detail-label">${column}</span><span class="detail-value">${value}</span></div>`;
    })
    .join("");

  const statusStepper = renderStatusStepper(row[TASK_COLUMNS.status]);
  const closeApprovers = collectTaskCloseApproverNames(row);
  const closeApproversHtml = closeApprovers.length
    ? closeApprovers
      .map((x) => `<span class="close-approver-chip" title="${escapeHtmlAttr(x.reason)}">${escapeHtmlText(x.name)}</span>`)
      .join("")
    : '<span class="close-approver-empty">Не определены (проверьте руководителя отдела и Telegram-подключение).</span>';
  const beforeGallery = buildDraftGallery(draftState.before, draftState.preview, "before");
  const afterGallery = buildDraftGallery(draftState.after, draftState.preview, "after");

  const modal = document.createElement("div");
  modal.className = "details-modal-overlay";
  modal.tabIndex = -1;
  const historyPanelHtml = renderTaskHistoryTableHtml(taskId);
  modal.innerHTML = `
    <div class="details-modal">
      <div class="details-modal-header">
        <h3>Карточка задачи #${taskId}</h3>
        <div class="details-modal-actions">
          <button type="button" class="icon-action-btn save-details-btn" title="Сохранить">
            <i data-lucide="save" class="lucide-icon" aria-hidden="true"></i>
          </button>
          <button type="button" class="icon-action-btn download-pdf-btn" title="Скачать PDF">
            <i data-lucide="download" class="lucide-icon" aria-hidden="true"></i>
          </button>
          <button type="button" class="icon-action-btn close-details-btn" title="Закрыть">
            <i data-lucide="x" class="lucide-icon" aria-hidden="true"></i>
          </button>
        </div>
      </div>
      <div class="task-card-tabs" role="tablist" aria-label="Разделы карточки">
        <button type="button" class="task-card-tab is-active" role="tab" aria-selected="true" data-task-tab="main">Карточка</button>
        <button type="button" class="task-card-tab" role="tab" aria-selected="false" data-task-tab="history">История</button>
      </div>
      <div class="task-card-panel" data-task-panel="main">
        ${statusStepper}
        <div class="close-approvers-box">
          <div class="close-approvers-title">Кто согласует закрытие</div>
          <div class="close-approvers-list">${closeApproversHtml}</div>
        </div>
        <div class="details-grid">${details}</div>
        <div class="gallery-block">
          <h4>Медиа до</h4>
          <div class="gallery-list" data-gallery-kind="before">${beforeGallery}</div>
        </div>
        <div class="gallery-block">
          <h4>Медиа после</h4>
          <div class="gallery-list" data-gallery-kind="after">${afterGallery}</div>
        </div>
      </div>
      <div class="task-card-panel task-card-panel--hidden" data-task-panel="history" role="tabpanel">
        ${historyPanelHtml}
      </div>
    </div>
  `;

  document.body.appendChild(modal);
  initLucideIcons();
  modal.querySelectorAll(".task-card-tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      const id = tab.getAttribute("data-task-tab");
      modal.querySelectorAll(".task-card-tab").forEach((t) => {
        const on = t.getAttribute("data-task-tab") === id;
        t.classList.toggle("is-active", on);
        t.setAttribute("aria-selected", on ? "true" : "false");
      });
      modal.querySelectorAll("[data-task-panel]").forEach((p) => {
        p.classList.toggle("task-card-panel--hidden", p.getAttribute("data-task-panel") !== id);
      });
    });
  });
  requestAnimationFrame(() => {
    modal.querySelector(".status-stepper")?.classList.add("animated");
    modal.focus();
  });

  let activeMediaTarget = null;
  const persistDraftMediaNow = () => {
    applyDraftToTask(section, rowIndex, taskId, draftState, { appendHistory: false });
    isDirty = false;
  };

  const refreshModalGalleries = () => {
    const beforeList = modal.querySelector('[data-gallery-kind="before"]');
    const afterList = modal.querySelector('[data-gallery-kind="after"]');
    if (beforeList) {
      beforeList.innerHTML = buildDraftGallery(draftState.before, draftState.preview, "before");
    }
    if (afterList) {
      afterList.innerHTML = buildDraftGallery(draftState.after, draftState.preview, "after");
    }
    bindGalleryControls();
  };

  const setActiveMediaSlot = (targetButton) => {
    modal.querySelectorAll(".gallery-slot").forEach((slot) => slot.classList.remove("is-active"));
    const slot = targetButton.closest(".gallery-slot");
    if (!slot) return;
    slot.classList.add("is-active");
    activeMediaTarget = {
      kind: String(slot.dataset.kind || "before"),
      slotIndex: Number(slot.dataset.slotIndex)
    };
  };

  const bindGalleryControls = () => {
    const chooseButtons = Array.from(modal.querySelectorAll(".slot-choose-btn"));
    const deleteButtons = Array.from(modal.querySelectorAll(".slot-delete-btn"));
    const imageItems = Array.from(modal.querySelectorAll(".gallery-slot.has-image .slot-preview"));

    chooseButtons.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        setActiveMediaSlot(button);
        pickFile(async (file) => {
          if (!file) return;
          const ok = await upsertDraftMediaSlot(draftState, activeMediaTarget.kind, activeMediaTarget.slotIndex, file);
          if (!ok) return;
          isDirty = true;
          persistDraftMediaNow();
          refreshModalGalleries();
        });
      });
    });

    deleteButtons.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        setActiveMediaSlot(button);
        removeDraftMediaSlot(draftState, activeMediaTarget.kind, activeMediaTarget.slotIndex);
        isDirty = true;
        persistDraftMediaNow();
        refreshModalGalleries();
      });
    });

    imageItems.forEach((item, index) => {
      item.addEventListener("click", () => {
        const clickable = Array.from(modal.querySelectorAll(".gallery-slot.has-image .slot-preview"));
        openGalleryLightbox(clickable, index);
      });
    });

    const allSlots = Array.from(modal.querySelectorAll(".gallery-slot"));
    allSlots.forEach((slot) => {
      slot.addEventListener("click", () => {
        const choose = slot.querySelector(".slot-choose-btn");
        if (choose) {
          setActiveMediaSlot(choose);
        }
      });
    });
  };

  bindGalleryControls();

  const onPaste = async (event) => {
    if (!document.body.contains(modal)) return;
    if (!activeMediaTarget) return;
    const clipboardItems = event.clipboardData?.items || [];
    const imageItem = Array.from(clipboardItems).find((item) => item.type.startsWith("image/"));
    if (!imageItem) return;
    const file = imageItem.getAsFile();
    if (!file) return;
    event.preventDefault();
    const ok = await upsertDraftMediaSlot(draftState, activeMediaTarget.kind, activeMediaTarget.slotIndex, file);
    if (!ok) return;
    isDirty = true;
    persistDraftMediaNow();
    refreshModalGalleries();
  };
  document.addEventListener("paste", onPaste);

  const close = () => modal.remove();
  const askUnsavedClose = () => {
    if (!isDirty) {
      closeAndCleanup();
      return;
    }
    showCustomCloseConfirm(modal, () => {
      closeAndCleanup();
    });
  };
  const closeAndCleanup = () => {
    document.removeEventListener("paste", onPaste);
    close();
  };
  modal.querySelector(".save-details-btn")?.addEventListener("click", () => {
    try {
      applyDraftToTask(section, rowIndex, taskId, draftState);
      isDirty = false;
      renderTablePreserveScroll();
      closeAndCleanup();
    } catch (e) {
      window.alert(`Не удалось сохранить карточку: ${String(e?.message || e)}`);
    }
  });
  modal.querySelector(".close-details-btn")?.addEventListener("click", askUnsavedClose);
  modal.querySelector(".download-pdf-btn")?.addEventListener("click", () => {
    downloadTaskCardAsPdf(modal, taskId);
  });
  modal.addEventListener("click", (event) => {
    if (event.target === modal) askUnsavedClose();
  });
}

function buildMediaGallery(value, taskId, colIndex) {
  const items = getMediaItems(value);
  const slots = Array.from({ length: 5 }, (_, slotIndex) => items[slotIndex] || "");
  return slots
    .map((name, slotIndex) => {
      const displayName = getMediaDisplayName(name);
      if (!name) {
        return `
          <div class="gallery-slot empty" data-col-index="${colIndex}" data-slot-index="${slotIndex}">
            <span>Пусто</span>
            <div class="slot-actions">
              <button type="button" class="slot-action-btn slot-choose-btn" title="Выбрать файл">+</button>
            </div>
          </div>
        `;
      }

      const slotKey = getMediaSlotKey(taskId, colIndex, slotIndex);
      const preview = resolveMediaPreviewForSlot(name, mediaPreviewStore[slotKey]);
      if (preview && String(preview.type || "").startsWith("image/")) {
        return `
          <figure class="gallery-slot has-image" data-col-index="${colIndex}" data-slot-index="${slotIndex}">
            <img class="slot-preview" src="${preview.url}" data-gallery-url="${preview.url}" data-gallery-name="${displayName}" alt="${displayName}" />
            <figcaption title="${escapeHtmlAttr(name)}">${displayName}</figcaption>
            <div class="slot-actions">
              <button type="button" class="slot-action-btn slot-choose-btn" title="Заменить">+</button>
              <button type="button" class="slot-action-btn slot-delete-btn" title="Удалить">x</button>
            </div>
          </figure>
        `;
      }
      return `
        <div class="gallery-slot file-item" data-col-index="${colIndex}" data-slot-index="${slotIndex}">
          <span title="${escapeHtmlAttr(name)}">${displayName}</span>
          <div class="slot-actions">
            <button type="button" class="slot-action-btn slot-choose-btn" title="Заменить">+</button>
            <button type="button" class="slot-action-btn slot-delete-btn" title="Удалить">x</button>
          </div>
        </div>
      `;
    })
    .join("");
}

function buildDraftGallery(items, previewMap, kind) {
  const slots = Array.from({ length: 5 }, (_, i) => items[i] || "");
  return slots.map((name, slotIndex) => {
    const displayName = getMediaDisplayName(name);
    const preview = resolveMediaPreviewForSlot(name, previewMap[`${kind}-${slotIndex}`]);
    if (!name) {
      return `
        <div class="gallery-slot empty" data-kind="${kind}" data-slot-index="${slotIndex}">
          <span>Пусто</span>
          <div class="slot-actions">
            <button type="button" class="slot-action-btn slot-choose-btn" title="Выбрать файл">+</button>
          </div>
        </div>
      `;
    }
    if (preview && String(preview.type || "").startsWith("image/") && preview.url) {
      return `
        <figure class="gallery-slot has-image" data-kind="${kind}" data-slot-index="${slotIndex}">
          <img class="slot-preview" src="${preview.url}" data-gallery-url="${preview.url}" data-gallery-name="${displayName}" alt="${displayName}" />
          <figcaption title="${escapeHtmlAttr(name)}">${displayName}</figcaption>
          <div class="slot-actions">
            <button type="button" class="slot-action-btn slot-choose-btn" title="Заменить">+</button>
            <button type="button" class="slot-action-btn slot-delete-btn" title="Удалить">x</button>
          </div>
        </figure>
      `;
    }
    return `
      <div class="gallery-slot file-item" data-kind="${kind}" data-slot-index="${slotIndex}">
        <span title="${escapeHtmlAttr(name)}">${displayName}</span>
        <div class="slot-actions">
          <button type="button" class="slot-action-btn slot-choose-btn" title="Заменить">+</button>
          <button type="button" class="slot-action-btn slot-delete-btn" title="Удалить">x</button>
        </div>
      </div>
    `;
  }).join("");
}

async function upsertDraftMediaSlot(draftState, kind, slotIndex, file) {
  const arr = kind === "after" ? draftState.after : draftState.before;
  const resolved = await resolveStoredMediaFromFile(file).catch((e) => {
    window.alert(`Не удалось загрузить медиа: ${String(e?.message || e)}`);
    return null;
  });
  if (!resolved) return false;
  arr[slotIndex] = resolved.stored;
  draftState.preview[`${kind}-${slotIndex}`] = {
    ...resolved.preview
  };
  return true;
}

function removeDraftMediaSlot(draftState, kind, slotIndex) {
  const arr = kind === "after" ? draftState.after : draftState.before;
  arr[slotIndex] = "";
  delete draftState.preview[`${kind}-${slotIndex}`];
}

function applyDraftToTask(section, rowIndex, taskId, draftState, options = {}) {
  const appendHistory = options.appendHistory !== false;
  const beforeItems = draftState.before.filter(Boolean).slice(0, 5);
  const afterItems = draftState.after.filter(Boolean).slice(0, 5);
  section.rows[rowIndex][TASK_COLUMNS.mediaBefore] = beforeItems.join(", ");
  section.rows[rowIndex][TASK_COLUMNS.mediaAfter] = afterItems.join(", ");

  // Синхронизируем превью по слотам
  for (let i = 0; i < 5; i += 1) {
    const beforeKey = getMediaSlotKey(taskId, TASK_COLUMNS.mediaBefore, i);
    const afterKey = getMediaSlotKey(taskId, TASK_COLUMNS.mediaAfter, i);
    delete mediaPreviewStore[beforeKey];
    delete mediaPreviewStore[afterKey];
  }
  draftState.before.forEach((name, i) => {
    if (!name) return;
    const preview = draftState.preview[`before-${i}`];
    if (preview) {
      mediaPreviewStore[getMediaSlotKey(taskId, TASK_COLUMNS.mediaBefore, i)] = { ...preview };
    }
  });
  draftState.after.forEach((name, i) => {
    if (!name) return;
    const preview = draftState.preview[`after-${i}`];
    if (preview) {
      mediaPreviewStore[getMediaSlotKey(taskId, TASK_COLUMNS.mediaAfter, i)] = { ...preview };
    }
  });
  if (appendHistory) {
    appendTaskHistoryEntry(taskId, "Сохранение карточки: обновлены «Медиа до» и «Медиа после»");
  }
  saveSectionsData();
}

function showCustomCloseConfirm(modal, onConfirm) {
  const old = modal.querySelector(".unsaved-confirm");
  old?.remove();
  const prompt = document.createElement("div");
  prompt.className = "unsaved-confirm";
  prompt.innerHTML = `
    <div class="unsaved-confirm-box">
      <p>Материалы не сохранены. Закрыть без сохранения?</p>
      <div class="unsaved-confirm-actions">
        <button type="button" class="confirm-btn confirm-cancel-btn">Отмена</button>
        <button type="button" class="confirm-btn confirm-close-btn">Закрыть</button>
      </div>
    </div>
  `;
  modal.appendChild(prompt);
  prompt.querySelector(".confirm-cancel-btn")?.addEventListener("click", () => prompt.remove());
  prompt.querySelector(".confirm-close-btn")?.addEventListener("click", () => {
    prompt.remove();
    onConfirm();
  });
  prompt.addEventListener("click", (event) => {
    if (event.target === prompt) {
      prompt.remove();
    }
  });
}

function upsertMediaSlot(section, rowIndex, taskId, colIndex, slotIndex, file) {
  const items = getMediaItems(section.rows[rowIndex][colIndex]);
  items[slotIndex] = file.name || `media-${Date.now()}.png`;
  setMediaItems(section, rowIndex, colIndex, items);
  const slotKey = getMediaSlotKey(taskId, colIndex, slotIndex);
  mediaPreviewStore[slotKey] = {
    name: file.name,
    type: file.type,
    url: URL.createObjectURL(file)
  };
}

function removeMediaSlot(section, rowIndex, taskId, colIndex, slotIndex) {
  const items = getMediaItems(section.rows[rowIndex][colIndex]);
  items[slotIndex] = "";
  setMediaItems(section, rowIndex, colIndex, items);
  const slotKey = getMediaSlotKey(taskId, colIndex, slotIndex);
  delete mediaPreviewStore[slotKey];
}

function mediaNameLooksLikeImage(name) {
  const n = String(name || "").trim().toLowerCase();
  if (!n) return false;
  return /\.(jpg|jpeg|png|gif|webp|bmp|svg)$/i.test(n);
}

function getMediaDisplayName(storedName) {
  const raw = String(storedName || "").trim();
  if (!raw) return "";
  let candidate = raw;
  if (/^https?:\/\//i.test(raw)) {
    try {
      const u = new URL(raw);
      candidate = u.pathname || raw;
    } catch {
      candidate = raw;
    }
  }
  const cleaned = candidate.split("#")[0].split("?")[0].replace(/\/+$/, "");
  const lastPart = cleaned.includes("/") ? cleaned.slice(cleaned.lastIndexOf("/") + 1) : cleaned;
  if (!lastPart) return raw;
  try {
    return decodeURIComponent(lastPart);
  } catch {
    return lastPart;
  }
}

function buildTelegramMediaPreviewUrl(storedName) {
  const token = String(displaySettings.telegramBotToken || "").trim();
  const p = String(storedName || "").trim().replace(/^\/+/, "");
  if (!token || !p || !p.includes("/")) return "";
  return `https://api.telegram.org/file/bot${token}/${p}`;
}

function toAbsoluteMediaUrl(storedName) {
  const raw = String(storedName || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  if (raw.startsWith("/media/")) return `${location.origin}${raw}`;
  if (raw.startsWith("media/")) return `${location.origin}/${raw}`;
  if (/^[\w.-]+\.(jpg|jpeg|png|gif|webp|bmp|svg)$/i.test(raw)) {
    return `${location.origin}/media/${encodeURIComponent(raw)}`;
  }
  return "";
}

function resolveMediaPreviewForSlot(storedName, preview) {
  if (preview && String(preview.type || "").startsWith("image/") && String(preview.url || "").trim()) {
    return preview;
  }
  const name = String(storedName || "").trim();
  if (!mediaNameLooksLikeImage(name)) return preview || null;
  const directUrl = toAbsoluteMediaUrl(name);
  if (directUrl) {
    return { name, type: "image/url", url: directUrl };
  }
  const tgUrl = buildTelegramMediaPreviewUrl(name);
  if (!tgUrl) return preview || null;
  return { name, type: "image/telegram", url: tgUrl };
}

function renderStatusStepper(currentStatus) {
  const steps = ["Новый", "В процессе", "Закрыт"];
  const currentIndex = getStatusStepIndex(currentStatus);
  const safeIndex = currentIndex >= 0 ? currentIndex : 0;
  const progress = steps.length > 1 ? (safeIndex / (steps.length - 1)) * 100 : 0;
  const items = steps
    .map((step, index) => {
      const stateClass = index < safeIndex ? "done" : index === safeIndex ? "active" : "";
      return `<div class="step ${stateClass}"><span>${step}</span></div>`;
    })
    .join("");
  return `<div class="status-stepper" style="--step-progress:${progress}%">${items}</div>`;
}

function getStatusStepIndex(status) {
  if (status === "Закрыт") return 2;
  if (status === "В процессе") return 1;
  return 0;
}

function openGalleryLightbox(imageItems, initialIndex) {
  if (!imageItems.length) return;
  let currentIndex = initialIndex;

  const lightbox = document.createElement("div");
  lightbox.className = "gallery-lightbox";
  lightbox.innerHTML = `
    <button type="button" class="lightbox-nav prev" title="Назад"><i data-lucide="chevron-left" class="lucide-icon"></i></button>
    <figure class="lightbox-content">
      <img class="lightbox-image" alt="" />
      <figcaption class="lightbox-caption"></figcaption>
    </figure>
    <button type="button" class="lightbox-nav next" title="Вперед"><i data-lucide="chevron-right" class="lucide-icon"></i></button>
    <button type="button" class="icon-action-btn lightbox-close" title="Закрыть"><i data-lucide="x" class="lucide-icon"></i></button>
  `;
  document.body.appendChild(lightbox);
  initLucideIcons();

  const imageEl = lightbox.querySelector(".lightbox-image");
  const captionEl = lightbox.querySelector(".lightbox-caption");

  const renderCurrent = () => {
    const item = imageItems[currentIndex];
    const url = item.dataset.galleryUrl || "";
    const name = item.dataset.galleryName || "";
    imageEl.src = url;
    imageEl.alt = name;
    captionEl.textContent = `${currentIndex + 1} / ${imageItems.length} - ${name}`;
  };

  const close = () => lightbox.remove();
  const prev = () => {
    currentIndex = (currentIndex - 1 + imageItems.length) % imageItems.length;
    renderCurrent();
  };
  const next = () => {
    currentIndex = (currentIndex + 1) % imageItems.length;
    renderCurrent();
  };

  lightbox.querySelector(".prev")?.addEventListener("click", prev);
  lightbox.querySelector(".next")?.addEventListener("click", next);
  lightbox.querySelector(".lightbox-close")?.addEventListener("click", close);
  lightbox.addEventListener("click", (event) => {
    if (event.target === lightbox) close();
  });
  document.addEventListener("keydown", function onKey(event) {
    if (!document.body.contains(lightbox)) {
      document.removeEventListener("keydown", onKey);
      return;
    }
    if (event.key === "Escape") close();
    if (event.key === "ArrowLeft") prev();
    if (event.key === "ArrowRight") next();
  });

  renderCurrent();
}

function downloadTaskCardAsPdf(modal, taskId) {
  const printable = modal.querySelector(".details-modal");
  if (!printable) return;
  const win = window.open("", "_blank");
  if (!win) return;
  win.document.write(`
    <html>
      <head>
        <title>Карточка задачи #${taskId}</title>
        <style>
          body{font-family:Arial,sans-serif;padding:20px;color:#1f2a37}
          .details-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;margin-bottom:14px}
          .detail-item{border:1px solid #dfe6ef;border-radius:8px;padding:8px}
          .detail-label{font-size:12px;color:#6a7685;display:block;margin-bottom:4px}
          .detail-value{font-size:13px}
          .gallery-list{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:8px}
          .gallery-slot{border:1px solid #dce4ee;border-radius:8px;padding:6px;min-height:120px}
          .gallery-slot img{width:100%;height:86px;object-fit:cover;border-radius:6px}
          .gallery-slot figcaption,.gallery-slot span{font-size:11px;word-break:break-word}
          .status-stepper{display:flex;gap:8px;margin:10px 0 14px}
          .status-stepper .step{border:1px solid #dce4ee;border-radius:999px;padding:4px 10px;font-size:12px}
          .status-stepper .step.active,.status-stepper .step.done{background:#ebeaf5;border-color:#c4c5df;color:#3e4095}
        </style>
      </head>
      <body>${printable.innerHTML}</body>
    </html>
  `);
  win.document.close();
  win.focus();
  win.print();
}

function exportRowsToCsv(section, filteredEntries, downloadName) {
  const header = section.columns.join(";");
  const rows = filteredEntries.map((entry) => entry.row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(";"));
  const csvContent = [header, ...rows].join("\n");
  const blob = new Blob([`\uFEFF${csvContent}`], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = downloadName || `${section.title}.xls`;
  link.click();
  URL.revokeObjectURL(url);
}

function printSectionTablePdf(section, filteredEntries) {
  ensureColumnDisplayState(section);
  const visibleColumnIndexes = getVisibleColumnIndexes(section);
  const showHeaderNumbers = headerNumberingBySection[section.id] !== false;
  const colCount = Math.max(1, visibleColumnIndexes.length);
  const fontSize = colCount >= 18 ? 7 : colCount >= 14 ? 8 : 9;
  const cellPadY = colCount >= 16 ? 3 : 5;
  const cellPadX = colCount >= 16 ? 4 : 6;

  const now = new Date();
  const p2 = (v) => String(v).padStart(2, "0");
  const generatedAt = `${p2(now.getDate())}.${p2(now.getMonth() + 1)}.${now.getFullYear()} ${p2(now.getHours())}:${p2(now.getMinutes())}`;
  const logoUrl = `${window.location.origin}/horizontal-v1.svg`;

  const headerMain = visibleColumnIndexes
    .map((colIndex) => `<th>${escapeHtmlText(String(section.columns[colIndex] || ""))}</th>`)
    .join("");
  const headerNumbers = showHeaderNumbers
    ? `<tr class="pdf-head-order">${visibleColumnIndexes.map((_, i) => `<th>${i + 1}</th>`).join("")}</tr>`
    : "";
  const bodyRows = filteredEntries.length
    ? filteredEntries
        .map((entry) => {
          const tds = visibleColumnIndexes.map((colIndex) => {
            const raw = String(entry.row?.[colIndex] ?? "");
            const html = escapeHtmlText(raw).replace(/\r?\n/g, "<br>");
            return `<td>${html || "—"}</td>`;
          }).join("");
          return `<tr>${tds}</tr>`;
        })
        .join("")
    : `<tr><td colspan="${visibleColumnIndexes.length}">Нет данных</td></tr>`;

  const html = `
  <html>
    <head>
      <meta charset="utf-8" />
      <title>${escapeHtmlText(section.title)} — PDF</title>
      <style>
        @page { size: A4 landscape; margin: 8mm; }
        * { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; }
        body { font-family: Inter, Segoe UI, Arial, sans-serif; color: #1f2a37; }
        .pdf-root { width: 100%; }
        .pdf-head {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
          border-bottom: 1px solid #d7e0ea;
          padding-bottom: 8px;
          margin-bottom: 8px;
        }
        .pdf-logo { height: 28px; width: auto; object-fit: contain; }
        .pdf-meta { text-align: right; font-size: 10px; color: #526174; }
        .pdf-meta b { color: #1f2a37; }
        .pdf-title { margin: 0 0 6px; font-size: 12px; font-weight: 700; color: #223247; }
        table {
          width: 100%;
          border-collapse: collapse;
          table-layout: fixed;
          font-size: ${fontSize}pt;
        }
        thead { display: table-header-group; }
        tfoot { display: table-footer-group; }
        tr { page-break-inside: avoid; }
        th, td {
          border: 1px solid #cfd8e3;
          padding: ${cellPadY}px ${cellPadX}px;
          vertical-align: top;
          word-break: break-word;
          overflow-wrap: anywhere;
        }
        thead th {
          background: #edf3f9;
          font-weight: 700;
          text-align: left;
        }
        thead tr.pdf-head-order th {
          text-align: center;
          font-weight: 500;
          color: #5b6b7f;
          border-top: none;
          background: #f5f8fc;
        }
        tbody tr:nth-child(even) td { background: #fbfdff; }
      </style>
    </head>
    <body>
      <div class="pdf-root">
        <div class="pdf-head">
          <img src="${logoUrl}" alt="Логотип" class="pdf-logo" />
          <div class="pdf-meta">
            <div><b>${escapeHtmlText(section.title)}</b></div>
            <div>Дата формирования отчёта: ${generatedAt}</div>
          </div>
        </div>
        <h1 class="pdf-title">Табличная часть</h1>
        <table>
          <thead>
            <tr>${headerMain}</tr>
            ${headerNumbers}
          </thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>
    </body>
  </html>`;

  const win = window.open("", "_blank");
  if (!win) {
    showStatusDialog({ title: "Печать", message: "Разрешите всплывающие окна для формирования PDF.", type: "error" });
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 180);
}

function openExportFormatModal(section, filteredEntries) {
  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal">
      <h3>${withIcon("file-text", "Скачать")}</h3>
      <p class="hint">Выберите формат выгрузки таблицы.</p>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary export-format-pdf-btn">PDF</button>
        <button type="button" class="secondary export-format-xls-btn">Excel</button>
        <button type="button" class="primary export-format-cancel-btn">Отмена</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  initLucideIcons();

  const close = () => overlay.remove();
  overlay.querySelector(".export-format-pdf-btn")?.addEventListener("click", () => {
    close();
    printSectionTablePdf(section, filteredEntries);
  });
  overlay.querySelector(".export-format-xls-btn")?.addEventListener("click", () => {
    close();
    exportRowsToCsv(section, filteredEntries);
  });
  overlay.querySelector(".export-format-cancel-btn")?.addEventListener("click", close);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
}

function openTaskImportModal(section) {
  if (section.id !== "tasks") return;
  const templateHeadHtml = TASK_IMPORT_COLUMNS.map((col) => `<th>${escapeHtmlText(col.label)}</th>`).join("");
  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal table-import-modal">
      <h3>${withIcon("file-up", "Импорт задач")}</h3>
      <p class="hint">Скопируйте строки из Excel по шаблону ниже и вставьте сюда через Ctrl+V.</p>
      <div class="task-import-template-wrap">
        <div class="task-import-template-head">
          <span>Шаблон колонок для импорта</span>
          <button type="button" class="secondary task-import-download-template-btn" id="taskImportDownloadTemplateBtn">Скачать шаблон</button>
        </div>
        <div class="task-import-template-table-wrap">
          <table class="task-import-template-table">
            <thead><tr>${templateHeadHtml}</tr></thead>
          </table>
        </div>
      </div>
      <textarea id="taskImportPasteInput" class="task-import-paste-input" rows="7" placeholder="Вставьте таблицу из Excel..."></textarea>
      <div class="task-import-preview-head">
        <strong>Превью</strong>
        <span id="taskImportPreviewMeta" class="hint">Строк: 0</span>
      </div>
      <div id="taskImportPreviewWrap" class="task-import-preview-wrap">
        <p class="hint">Данные ещё не вставлены.</p>
      </div>
      <label class="settings-option task-import-sync-option">
        <input type="checkbox" id="taskImportSyncCatalogsCheckbox" />
        <span>Добавлять не найденные фаза/раздел/подраздел в справочники и связку «Ответственные»</span>
      </label>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary task-import-cancel-btn">Отмена</button>
        <button type="button" class="secondary task-import-save-btn" disabled>Сохранить</button>
        <button type="button" class="responsible-apply-btn task-import-save-send-btn" disabled>Сохранить и отправить</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  initLucideIcons();

  const input = overlay.querySelector("#taskImportPasteInput");
  const previewWrap = overlay.querySelector("#taskImportPreviewWrap");
  const previewMeta = overlay.querySelector("#taskImportPreviewMeta");
  const saveBtn = overlay.querySelector(".task-import-save-btn");
  const saveSendBtn = overlay.querySelector(".task-import-save-send-btn");
  const cancelBtn = overlay.querySelector(".task-import-cancel-btn");
  const downloadTemplateBtn = overlay.querySelector("#taskImportDownloadTemplateBtn");
  const syncCatalogsCheckbox = overlay.querySelector("#taskImportSyncCatalogsCheckbox");
  const close = () => overlay.remove();

  let parsedRows = [];
  let inspectedRows = [];

  const renderPreview = () => {
    if (!previewWrap || !previewMeta || !saveBtn || !saveSendBtn) return;
    const allowCatalogSync = Boolean(syncCatalogsCheckbox?.checked);
    const canImportRows = inspectedRows.filter((x) => !x.hasMissing || allowCatalogSync);
    const blockedRows = inspectedRows.length - canImportRows.length;
    previewMeta.textContent = `Строк: ${inspectedRows.length} · к импорту: ${canImportRows.length}`;
    if (!parsedRows.length) {
      previewWrap.innerHTML = '<p class="hint">Данные ещё не вставлены.</p>';
      saveBtn.disabled = true;
      saveSendBtn.disabled = true;
      return;
    }
    const head = TASK_IMPORT_COLUMNS.map((col) => `<th>${escapeHtmlText(col.label)}</th>`).join("");
    const body = inspectedRows
      .map((item, index) => {
        const row = item.row;
        const cells = TASK_IMPORT_COLUMNS.map((col) => {
          const raw = String(row[col.key] || "—");
          const safe = escapeHtmlText(raw);
          return `<td title="${escapeHtmlAttr(raw)}">${safe}</td>`;
        }).join("");
        const badge = item.hasMissing
          ? `<div class="task-import-missing-badge">Нет в справочнике: ${escapeHtmlText(item.missing.join("; "))}</div>`
          : "";
        const rowClass = item.hasMissing ? "task-import-preview-row--missing" : "";
        return `<tr class="${rowClass}"><td><div>${index + 1}</div>${badge}</td>${cells}</tr>`;
      })
      .join("");
    previewWrap.innerHTML = `
      <table class="task-import-preview-table">
        <thead><tr><th>#</th>${head}</tr></thead>
        <tbody>${body}</tbody>
      </table>
      ${blockedRows > 0 && !allowCatalogSync ? `<p class="task-import-preview-note">Не будут импортированы: ${blockedRows} строк(и) — включите чекбокс выше, чтобы добавить новые значения в справочники.</p>` : ""}
    `;
    saveBtn.disabled = canImportRows.length === 0;
    saveSendBtn.disabled = canImportRows.length === 0;
  };

  const parseFromInput = () => {
    const res = parseTaskImportPayload(input?.value || "");
    if (!res.ok) {
      parsedRows = [];
      inspectedRows = [];
      renderPreview();
      return;
    }
    parsedRows = res.rows;
    const catalogs = {
      phases: getCatalogValueSet("phases"),
      sections: getCatalogValueSet("phaseSections"),
      subsections: getCatalogValueSet("phaseSubsections")
    };
    inspectedRows = parsedRows.map((row) => {
      const inspected = inspectImportedHierarchyValue(row, catalogs);
      return { row, ...inspected };
    });
    renderPreview();
  };

  const buildRows = ({ allowCatalogSync = false } = {}) => {
    ensureTaskIdCounter();
    const rows = [];
    inspectedRows.forEach((item) => {
      if (item.hasMissing && !allowCatalogSync) return;
      if (item.hasMissing && allowCatalogSync) {
        ensureHierarchyValuesInCatalogs(item.row);
        ensureResponsibleHierarchyLink(item.phase, item.phaseSection, item.phaseSubsection);
      }
      rows.push(createTaskRowFromImport(item.row, allocateNextTaskId()));
    });
    return rows;
  };

  const appendImportedRows = (rows, { markPending = true } = {}) => {
    if (!rows.length) return;
    rows.forEach((row) => {
      section.rows.push(row);
      if (markPending) markTaskAsPendingImported(row[TASK_COLUMNS.number], { save: false });
    });
    saveSectionsData();
    saveDisplaySettings();
  };

  const sendImportedRows = async (rows) => {
    const employeeSet = getEmployeeNameSet();
    let sentOk = 0;
    const notSent = [];
    for (const row of rows) {
      const taskId = normalizeTaskIdValue(row[TASK_COLUMNS.number]);
      const assignedName = normalizePersonName(row[TASK_COLUMNS.assignedResponsible]);
      if (!assignedName || !hasSystemEmployeeName(assignedName, employeeSet)) {
        notSent.push(`ID ${taskId || "—"}: не найден сотрудник «${assignedName || "—"}»`);
        markTaskAsPendingImported(taskId, { save: false });
        continue;
      }
      const result = await sendTaskRowTelegramNotification(row, { suppressAlerts: true });
      if (result.ok) {
        markTaskAsSentImported(taskId, { save: false });
        sentOk += 1;
      } else {
        notSent.push(`ID ${taskId || "—"}: ${result.message || result.reason || "ошибка отправки"}`);
        markTaskAsPendingImported(taskId, { save: false });
      }
    }
    saveDisplaySettings();
    return { sentOk, total: rows.length, notSent };
  };

  saveBtn?.addEventListener("click", () => {
    if (!parsedRows.length) return;
    const rows = buildRows({ allowCatalogSync: Boolean(syncCatalogsCheckbox?.checked) });
    if (!rows.length) return;
    appendImportedRows(rows, { markPending: true });
    close();
    resetTasksListPagingWindow();
    renderTablePreserveScroll();
    showStatusDialog({
      title: "Импорт задач",
      message: `Сохранено задач: ${rows.length}.\nОни добавлены во вкладку «Не отправленные».`,
      type: "success"
    });
  });

  saveSendBtn?.addEventListener("click", () => {
    if (!parsedRows.length) return;
    void (async () => {
      const rows = buildRows({ allowCatalogSync: Boolean(syncCatalogsCheckbox?.checked) });
      if (!rows.length) return;
      appendImportedRows(rows, { markPending: true });
      const result = await sendImportedRows(rows);
      close();
      resetTasksListPagingWindow();
      renderTablePreserveScroll();
      const message = result.notSent.length
        ? `Отправлено: ${result.sentOk} из ${result.total}.\n\nОстались в «Не отправленные»:\n${result.notSent.slice(0, 10).join("\n")}${result.notSent.length > 10 ? `\n… и ещё ${result.notSent.length - 10}` : ""}`
        : `Отправлено: ${result.sentOk} из ${result.total}.\nВсе задачи успешно доставлены.`;
      showStatusDialog({
        title: "Импорт задач",
        message,
        type: result.notSent.length ? "info" : "success"
      });
    })();
  });

  downloadTemplateBtn?.addEventListener("click", () => {
    const header = TASK_IMPORT_COLUMNS.map((col) => col.label);
    const rows = [
      ["ЖК REGNUM PLAZA", "Высокий", "17.04.2026", "Инициация", "Первичный анализ ЗУ", "СТУ", "Проверить комплект документов", "Эльбек Ризаев", "Фамилия Имя", "Комментарий к задаче 1", "24.04.2026"],
      ["Center one by MBC", "Средний", "17.04.2026", "Проработка", "Маркетинг", "Другое", "Подготовить витрину продаж", "Эльбек Ризаев", "Фамилия Имя", "Комментарий к задаче 2", "30.04.2026"],
      ["ЖК Saadiyat", "Критический", "17.04.2026", "Реализация", "СМР", "Надземные конструкции", "Проверить замечания подрядчика", "Эльбек Ризаев", "Фамилия Имя", "Комментарий к задаче 3", "05.05.2026"]
    ];
    const lines = [header, ...rows]
      .map((cols) => cols.map((v) => String(v || "").replace(/\r\n?/g, " ").trim()).join("\t"))
      .join("\n");
    const blob = new Blob([`\uFEFF${lines}`], { type: "text/tab-separated-values;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "template_task_import.xls";
    a.click();
    URL.revokeObjectURL(url);
  });

  input?.addEventListener("input", parseFromInput);
  syncCatalogsCheckbox?.addEventListener("change", renderPreview);
  cancelBtn?.addEventListener("click", close);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
  renderPreview();
  input?.focus();
}

function openTableSettingsModal(section) {
  ensureColumnDisplayState(section);
  const FIXED_COLUMN_INDEX = 0;
  const initialOrder = [...columnOrderBySection[section.id]];
  const initialVisibility = [...visibleColumnsBySection[section.id]];
  const initialHeaderNumbers = headerNumberingBySection[section.id] !== false;

  const draftOrder = [...initialOrder];
  const draftVisibility = [...initialVisibility];
  let draftHeaderNumbers = initialHeaderNumbers;
  let dragColIndex = null;

  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal table-settings-modal">
      <h3>${withIcon("settings", "Настройки таблицы")}</h3>
      <label class="settings-option table-settings-flag">
        <input type="checkbox" id="tableSettingsHeaderNumbersToggle" ${draftHeaderNumbers ? "checked" : ""} />
        <span>Отобразить нумерацию заголовка</span>
      </label>
      <div class="table-settings-dnd-list" id="tableSettingsDndList"></div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary table-settings-cancel-btn">Отмена</button>
        <button type="button" class="responsible-apply-btn table-settings-apply-btn">Сохранить</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  initLucideIcons();

  const listEl = overlay.querySelector("#tableSettingsDndList");
  const headerNumbersToggle = overlay.querySelector("#tableSettingsHeaderNumbersToggle");
  const cancelBtn = overlay.querySelector(".table-settings-cancel-btn");
  const applyBtn = overlay.querySelector(".table-settings-apply-btn");

  const renderList = () => {
    if (!listEl) return;
    listEl.innerHTML = draftOrder
      .map((colIndex, position) => {
        const title = String(section.columns[colIndex] || `Колонка ${colIndex + 1}`);
        const checked = draftVisibility[colIndex] !== false;
        const isFixed = colIndex === FIXED_COLUMN_INDEX;
        const cbId = `tableSettingsColVisible_${section.id}_${colIndex}`;
        return `
          <div class="table-settings-dnd-item${isFixed ? " is-fixed" : ""}" data-col-index="${colIndex}" draggable="${isFixed ? "false" : "true"}">
            <span class="table-settings-dnd-pos">${position + 1}</span>
            <input id="${escapeHtmlAttr(cbId)}" type="checkbox" class="table-settings-col-visible" data-col-index="${colIndex}" ${checked ? "checked" : ""} ${isFixed ? "disabled" : ""} />
            <label for="${escapeHtmlAttr(cbId)}" class="table-settings-dnd-title">${escapeHtmlText(title)}</label>
            <span class="table-settings-dnd-handle" title="${isFixed ? "Фиксированный столбец" : "Перетащите мышью"}">
              <i data-lucide="${isFixed ? "lock" : "grip-vertical"}" class="lucide-icon" aria-hidden="true"></i>
            </span>
          </div>
        `;
      })
      .join("");
    initLucideIcons();

    Array.from(listEl.querySelectorAll(".table-settings-col-visible")).forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        const colIndex = Number(checkbox.dataset.colIndex);
        if (!Number.isInteger(colIndex) || colIndex < 0) return;
        draftVisibility[colIndex] = checkbox.checked;
        draftVisibility[FIXED_COLUMN_INDEX] = true;
        const nonFixedVisible = draftOrder.some((idx) => idx !== FIXED_COLUMN_INDEX && draftVisibility[idx] !== false);
        if (!nonFixedVisible) {
          draftVisibility[colIndex] = true;
          checkbox.checked = true;
        }
      });
    });

    const items = Array.from(listEl.querySelectorAll(".table-settings-dnd-item"));
    items.forEach((item) => {
      const colIndex = Number(item.dataset.colIndex);
      const isFixed = colIndex === FIXED_COLUMN_INDEX;
      if (isFixed) return;

      item.addEventListener("dragstart", (event) => {
        dragColIndex = colIndex;
        item.classList.add("is-dragging");
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", String(colIndex));
        }
      });
      item.addEventListener("dragend", () => {
        dragColIndex = null;
        items.forEach((el) => el.classList.remove("is-dragging", "drop-target"));
      });
      item.addEventListener("dragover", (event) => {
        event.preventDefault();
        if (dragColIndex == null || dragColIndex === colIndex) return;
        items.forEach((el) => el.classList.remove("drop-target"));
        item.classList.add("drop-target");
      });
      item.addEventListener("dragleave", () => {
        item.classList.remove("drop-target");
      });
      item.addEventListener("drop", (event) => {
        event.preventDefault();
        item.classList.remove("drop-target");
        if (dragColIndex == null || dragColIndex === colIndex) return;
        const from = draftOrder.indexOf(dragColIndex);
        const to = draftOrder.indexOf(colIndex);
        if (from < 1 || to < 1) return;
        draftOrder.splice(from, 1);
        draftOrder.splice(to, 0, dragColIndex);
        renderList();
      });
    });
  };

  headerNumbersToggle?.addEventListener("change", () => {
    draftHeaderNumbers = Boolean(headerNumbersToggle.checked);
  });

  cancelBtn?.addEventListener("click", () => {
    overlay.remove();
  });
  applyBtn?.addEventListener("click", () => {
    visibleColumnsBySection[section.id] = [...draftVisibility];
    visibleColumnsBySection[section.id][FIXED_COLUMN_INDEX] = true;
    columnOrderBySection[section.id] = [...draftOrder];
    headerNumberingBySection[section.id] = draftHeaderNumbers;
    overlay.remove();
    renderTablePreserveScroll();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) overlay.remove();
  });

  renderList();
}

function attachFilterHandlers(section) {
  const filterPanel = document.querySelector(".filter-panel");
  if (!filterPanel) {
    return;
  }

  const bumpTasksPagingReset = () => {
    if (section.id === "tasks") resetTasksListPagingWindow();
  };

  const searchInput = document.getElementById("filterSearch");
  const resetButton = document.getElementById("filterResetBtn");
  const statusSelect = document.getElementById("filterStatus");
  const responsibleSelect = document.getElementById("filterResponsible");
  const objectSelect = document.getElementById("filterObject");
  const phaseSelect = document.getElementById("filterPhase");
  const sectionSelect = document.getElementById("filterSection");
  const subsectionSelect = document.getElementById("filterSubsection");
  const readStateSelect = document.getElementById("filterReadState");

  const ensureSectionFilters = () => {
    if (!filtersBySection[section.id]) {
      filtersBySection[section.id] = {};
    }
    return filtersBySection[section.id];
  };

  if (searchInput) {
    searchInput.addEventListener("input", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.search = searchInput.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (statusSelect) {
    statusSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.status = statusSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (responsibleSelect) {
    responsibleSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.responsible = responsibleSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (objectSelect) {
    objectSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.object = objectSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (phaseSelect) {
    phaseSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.phase = phaseSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (sectionSelect) {
    sectionSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.section = sectionSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (subsectionSelect) {
    subsectionSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.subsection = subsectionSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (readStateSelect) {
    readStateSelect.addEventListener("change", () => {
      const sectionFilters = ensureSectionFilters();
      sectionFilters.readState = readStateSelect.value;
      bumpTasksPagingReset();
      renderTable();
    });
  }

  if (resetButton) {
    resetButton.addEventListener("click", () => {
      filtersBySection[section.id] = {};
      bumpTasksPagingReset();
      renderTable();
    });
  }
}

function showApp(userName) {
  authExpiredNoticeShown = false;
  document.body.classList.remove("login-mode");
  document.body.classList.remove("shared-report-mode");
  sharedReportMode = false;
  reportShareRowsOverride = null;
  sharedReportExpiresAt = 0;
  if (currentUser) {
    currentUser.innerHTML = withIcon("user", `Пользователь: ${userName}`);
  }
  loginSection.classList.add("hidden");
  appSection.classList.remove("hidden");
  renderSidebarMenu();
  renderTable();
  startRemoteAutoPull();
  startOverdueTaskNotificationsScheduler();
  startSessionIdleWatcher();
  hideBootLoaderAfterRender();
}

function showLogin() {
  document.body.classList.add("login-mode");
  document.body.classList.remove("shared-report-mode");
  sharedReportMode = false;
  reportShareRowsOverride = null;
  sharedReportExpiresAt = 0;
  appSection.classList.add("hidden");
  loginSection.classList.remove("hidden");
  loginForm.reset();
  if (phoneInput) {
    phoneInput.value = DEFAULT_PHONE_PREFIX;
    updateLoginPhoneFlag();
  }
  if (passwordInput) {
    passwordInput.type = "password";
  }
  if (togglePasswordBtn) {
    togglePasswordBtn.title = "Показать пароль";
    togglePasswordBtn.setAttribute("aria-label", "Показать пароль");
    togglePasswordBtn.innerHTML = `<i data-lucide="eye" class="lucide-icon" aria-hidden="true"></i>`;
    initLucideIcons();
  }
  activeSectionId = "tasks";
  saveActiveSection("tasks");
  isSettingsOpen = false;
  stopRemoteAutoPull();
  stopOverdueTaskNotificationsScheduler();
  stopSessionIdleWatcher();
  currentAuthRole = "user";
  hideBootLoaderAfterRender();
}

function togglePasswordVisibility() {
  if (!passwordInput || !togglePasswordBtn) return;
  const nextVisible = passwordInput.type === "password";
  passwordInput.type = nextVisible ? "text" : "password";
  togglePasswordBtn.title = nextVisible ? "Скрыть пароль" : "Показать пароль";
  togglePasswordBtn.setAttribute("aria-label", nextVisible ? "Скрыть пароль" : "Показать пароль");
  togglePasswordBtn.innerHTML = `<i data-lucide="${nextVisible ? "eye-off" : "eye"}" class="lucide-icon" aria-hidden="true"></i>`;
  initLucideIcons();
}

function enforceUzPhonePrefix() {
  if (!phoneInput) return;
  const normalized = normalizeUzPhone(phoneInput.value);
  phoneInput.value = formatUzPhoneDisplay(normalized);
  updateLoginPhoneFlag();
}

function addEmptyRow(section) {
  const row = new Array(section.columns.length).fill("");
  const numericIds = section.rows
    .map((item) => Number(item[0]))
    .filter((value) => Number.isFinite(value));
  const nextId = numericIds.length ? Math.max(...numericIds) + 1 : section.rows.length + 1;
  row[0] = section.id === "tasks" ? allocateNextTaskId() : String(nextId);

  if (section.id === "tasks") {
    row[TASK_COLUMNS.status] = "Новый";
    row[TASK_COLUMNS.priority] = "Средний";
    row[TASK_COLUMNS.addedDate] = getTodayRuDate();
    row[TASK_COLUMNS.readState] = composeTaskReadState(false, "—");
    row[TASK_COLUMNS.lastSentAt] = "—";
  }
  if (section.id === "employees") {
    row[EMPLOYEE_COLUMNS.department] = String(getSectionById("departments")?.rows?.[0]?.[1] || "");
    row[EMPLOYEE_COLUMNS.phone] = DEFAULT_PHONE_PREFIX;
    row[EMPLOYEE_COLUMNS.telegram] = "Не подключен";
    row[EMPLOYEE_COLUMNS.chatId] = "";
    row[EMPLOYEE_COLUMNS.activity] = "Не активен";
  }
  if (section.id === "roles") {
    row[2] = "Не системный";
  }
  if (section.id === "departments") {
    row[3] = "Не системный";
  }

  section.rows.push(row);
}

function getTodayRuDate() {
  const parts = getCalendarDatePartsInTimeZone(new Date(), getServerTimezone());
  if (!parts) {
    const date = new Date();
    return formatDatePartsStorage(date.getDate(), date.getMonth() + 1, date.getFullYear());
  }
  return formatDatePartsStorage(parts.day, parts.month, parts.year);
}

function saveSectionsData() {
  normalizeTaskMultiStateStore();
  localStorage.setItem(DATA_STORAGE_KEY, JSON.stringify(sections));
  localStorage.setItem(TASK_MULTI_STATE_STORAGE_KEY, JSON.stringify(taskMultiState));
  scheduleServerSync();
}

function saveTaskMultiState(opts = {}) {
  normalizeTaskMultiStateStore();
  localStorage.setItem(TASK_MULTI_STATE_STORAGE_KEY, JSON.stringify(taskMultiState));
  if (!opts.skipServerSync) {
    scheduleServerSync();
  }
}

function restoreTaskMultiState() {
  const raw = localStorage.getItem(TASK_MULTI_STATE_STORAGE_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return;
    taskMultiState = parsed;
    normalizeTaskMultiStateStore();
  } catch (_) {
    // ignore broken storage
  }
}

function getTrashRows(sectionId) {
  if (!trashBySection[sectionId]) {
    trashBySection[sectionId] = [];
  }
  return trashBySection[sectionId];
}

function clearTasksTrashNow({ save = true } = {}) {
  const list = getTrashRows("tasks");
  if (!Array.isArray(list) || list.length === 0) return false;
  trashBySection.tasks = [];
  if (save) {
    saveTrashData();
  }
  return true;
}

function moveTaskToTrash(sectionId, rowIndex) {
  const section = sections.find((item) => item.id === sectionId);
  if (!section) return;
  const row = section.rows[rowIndex];
  if (!row) return;
  if (sectionId === "tasks") {
    markTaskAsSentImported(row[TASK_COLUMNS.number], { save: false });
  }
  section.rows.splice(rowIndex, 1);
  const now = Date.now();
  const expiresAt = now + (365 * 24 * 60 * 60 * 1000);
  getTrashRows(sectionId).push({ row, deletedAt: now, expiresAt });
}

function restoreTaskFromTrash(sectionId, trashIndex) {
  const section = sections.find((item) => item.id === sectionId);
  const trash = getTrashRows(sectionId);
  if (!section || !trash[trashIndex]) return;
  const [item] = trash.splice(trashIndex, 1);
  section.rows.push(item.row);
}

function cleanupExpiredTrash(sectionId) {
  const now = Date.now();
  const trash = getTrashRows(sectionId);
  const next = trash.filter((item) => Number(item.expiresAt || 0) > now);
  if (next.length !== trash.length) {
    trashBySection[sectionId] = next;
    saveTrashData();
  }
}

function saveTrashData() {
  localStorage.setItem(TRASH_STORAGE_KEY, JSON.stringify(trashBySection));
  scheduleServerSync();
}

function restoreTrashData() {
  const raw = localStorage.getItem(TRASH_STORAGE_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return;
    Object.keys(parsed).forEach((key) => {
      if (Array.isArray(parsed[key])) {
        trashBySection[key] = parsed[key];
      }
    });
  } catch (_) {
    // ignore broken storage
  }
}

function isTrashTab(sectionId) {
  return (statusTabBySection[sectionId] || "all") === "trash";
}

function getSelectionKey(sectionId) {
  return isTrashTab(sectionId) ? `${sectionId}::trash` : sectionId;
}

function confirmAction({ message, confirmLabel = "Да", onConfirm }) {
  const overlay = document.createElement("div");
  overlay.className = "unsaved-confirm";
  overlay.innerHTML = `
    <div class="unsaved-confirm-box">
      <p>${message}</p>
      <div class="unsaved-confirm-actions">
        <button type="button" class="confirm-btn confirm-cancel-btn">Нет</button>
        <button type="button" class="confirm-btn confirm-close-btn">${confirmLabel}</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  overlay.querySelector(".confirm-cancel-btn")?.addEventListener("click", () => overlay.remove());
  overlay.querySelector(".confirm-close-btn")?.addEventListener("click", () => {
    overlay.remove();
    onConfirm?.();
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) overlay.remove();
  });
}

function restoreSectionsData() {
  const raw = localStorage.getItem(DATA_STORAGE_KEY);
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return;
    const byId = new Map(parsed.map((item) => [item.id, item]));
    sections = sections.map((base) => {
      const saved = byId.get(base.id);
      if (!saved || !Array.isArray(saved.rows)) return base;
      if (
        saved.rows.length === 0 &&
        (base.id === "phases" || base.id === "phaseSections" || base.id === "phaseSubsections")
      ) {
        return base;
      }
      const migratedRows = saved.rows.map((row) => migrateRowForSection(base, row, saved.columns));
      return {
        ...base,
        rows: migratedRows
      };
    });
    ensureSystemRoles();
    ensureSystemDepartments();
    normalizePhaseAndSectionCatalogs();
    if (ensureNonEmptyPhaseCatalogs()) {
      saveSectionsData();
    }
    syncEmployeesDerivedFields();
    const employeeSet = getEmployeeNameSet();
    const taskSection = getSectionById("tasks");
    if (taskSection) {
      repairTaskIdsInRows(taskSection.rows);
      taskSection.rows.forEach((row) => {
        repairTaskRowCells(row, employeeSet);
      });
    }
    normalizeTaskMultiStateStore();
    ensureTaskIdCounter();
    saveDisplaySettings({ skipServerSync: true });
    // Фиксируем нормализованный порядок колонок задач в localStorage.
    saveSectionsData();
  } catch (_) {
    // broken storage ignored
  }
}

function applyObjectsSeedIfNeeded() {
  try {
    if (localStorage.getItem(OBJECTS_SEED_VERSION_KEY) === String(OBJECTS_SEED_VERSION)) return;
    const idx = sections.findIndex((s) => s.id === "objects");
    if (idx === -1) return;
    sections[idx] = {
      ...sections[idx],
      rows: JSON.parse(JSON.stringify(DEFAULT_OBJECTS_ROWS))
    };
    localStorage.setItem(OBJECTS_SEED_VERSION_KEY, String(OBJECTS_SEED_VERSION));
    saveSectionsData();
  } catch (_) {
    /* noop */
  }
}

function normalizeTaskColumnLabel(raw) {
  return String(raw || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[^a-zа-я0-9]+/giu, "");
}

function looksLikePersonName(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  if (text.length > 80) return false;
  const words = text.split(/\s+/).filter(Boolean);
  if (words.length < 2 || words.length > 4) return false;
  return words.every((w) => /^[A-Za-zА-Яа-яЁё-]+$/u.test(w));
}

function looksLikeTaskText(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  if (looksLikePersonName(text)) return false;
  if (text.length < 10) return false;
  const words = text.split(/\s+/).filter(Boolean);
  return words.length >= 2;
}

function repairTaskRowCells(row, employeeSet = null) {
  if (!Array.isArray(row)) return false;
  const knownEmployees = employeeSet || getEmployeeNameSet();
  const currentTask = String(row[TASK_COLUMNS.task] ?? "").trim();
  const currentResponsible = String(row[TASK_COLUMNS.responsible] ?? "").trim();
  const taskLooksLikeEmployee = currentTask && knownEmployees.has(normalizePersonName(currentTask));
  const responsibleLooksLikeEmployee = currentResponsible && knownEmployees.has(normalizePersonName(currentResponsible));
  const shouldSwapByEmployeeSet = taskLooksLikeEmployee && !responsibleLooksLikeEmployee && looksLikeTaskText(currentResponsible);
  const shouldSwapByHeuristic = looksLikePersonName(currentTask) && looksLikeTaskText(currentResponsible);
  if (!shouldSwapByEmployeeSet && !shouldSwapByHeuristic) return false;
  row[TASK_COLUMNS.task] = currentResponsible;
  row[TASK_COLUMNS.responsible] = currentTask;
  return true;
}

function repairTaskIdsInRows(rows) {
  if (!Array.isArray(rows)) return false;
  const seen = new Set();
  let changed = false;
  let nextId = ensureTaskIdCounter();
  rows.forEach((row) => {
    if (!Array.isArray(row)) return;
    const rawId = normalizeTaskIdValue(row[TASK_COLUMNS.number]);
    const parsed = readNumericTaskId(rawId);
    if (!parsed || seen.has(parsed)) {
      row[TASK_COLUMNS.number] = String(nextId);
      seen.add(nextId);
      nextId += 1;
      changed = true;
      return;
    }
    seen.add(parsed);
    if (parsed >= nextId) {
      nextId = parsed + 1;
    }
  });
  if (displaySettings.taskIdCounter !== nextId) {
    displaySettings.taskIdCounter = nextId;
    changed = true;
  }
  return changed;
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

function migrateRowForSection(baseSection, row, sourceColumns = null) {
  let source = Array.isArray(row) ? [...row] : [];

  if (baseSection.id === "tasks") {
    if (Array.isArray(sourceColumns) && sourceColumns.length) {
      source = remapTaskRowToCurrentOrder(source, sourceColumns);
    }

    source[TASK_COLUMNS.status] = normalizeTaskStatusValue(source[TASK_COLUMNS.status]);
    source[TASK_COLUMNS.priority] = normalizeTaskPriorityValue(source[TASK_COLUMNS.priority]);

    // Старый формат задач без колонки «Исполнитель».
    if (source.length === 17) {
      // Исторически в len=17 поля идут как:
      // ... Подраздел, Задача, Постановщик, Комментарий, План, Факт, ...
      // Добавляем «Ответственного» (assignedResponsible) после постановщика.
      source.splice(10, 0, "");
    }
    if (source.length === 18) {
      source.push("Не прочитано\n—");
    }
    if (source.length === 19) {
      source.push("—");
    }

    // Перестановка порядка колонок (старый -> новый):
    // old: [.., assignedResponsible, task, responsible, ...]
    // new: [.., task, responsible, assignedResponsible, ...]
    if (source.length >= 20) {
      const oldAssigned = source[8];
      const oldTask = source[9];
      const oldResponsible = source[10];
      const looksLikeOldOrder =
        looksLikePersonName(oldAssigned) &&
        looksLikeTaskText(oldTask) &&
        (String(oldResponsible || "").trim() === "" || looksLikePersonName(oldResponsible));
      if (looksLikeOldOrder) {
        source[8] = oldTask;
        source[9] = oldResponsible;
        source[10] = oldAssigned;
      }

      repairTaskRowCells(source);
    }

    if (source.length > TASK_COLUMNS.lastSentAt && !String(source[TASK_COLUMNS.lastSentAt] || "").trim()) {
      source[TASK_COLUMNS.lastSentAt] = "—";
    }
  }

  if (baseSection.id === "data") {
    // Старый формат данных без колонки "Ответственные".
    if (source.length === 5) {
      source.splice(4, 0, "");
    }
  }

  if (baseSection.id === "objects") {
    // Старый формат: ID, Наименование, Тип, Адрес, Статус.
    if (source.length === 5) {
      return [source[0] || "", source[1] || "", source[3] || "", source[4] || "", "", "", ""];
    }
    if (source.length >= 7) {
      const photoRaw = String(source[OBJECT_COLUMNS.photo] || "").trim();
      if (photoRaw === "-" || photoRaw === "—" || photoRaw.toLowerCase() === "null" || photoRaw.toLowerCase() === "undefined") {
        source[OBJECT_COLUMNS.photo] = "";
      }
    }
  }

  if (baseSection.id === "employees") {
    // Старый формат: ID, ФИО, Отдел, Роль, Активность.
    if (source.length === 5) {
      return [source[0] || "", source[1] || "", source[2] || "", source[3] || "", DEFAULT_PHONE_PREFIX, "Не подключен", "", source[4] || ""];
    }
  }

  if (baseSection.id === "roles") {
    if (source.length === 1) {
      source.unshift("", "Не системный");
    } else if (source.length === 2) {
      source.push(getRoleTypeLabel(source[1]));
    }
  }

  if (baseSection.id === "phases") {
    if (source.length === 1) {
      source.unshift("");
    }
  }

  if (baseSection.id === "phaseSections") {
    if (source.length === 3) {
      return [source[0] || "", source[2] || ""];
    }
    if (source.length === 2) {
      source.unshift("");
    }
  }

  if (baseSection.id === "phaseSubsections") {
    if (source.length === 4) {
      return [source[0] || "", source[3] || ""];
    }
    if (source.length === 3) {
      source.unshift("");
    }
  }

  if (baseSection.id === "departments") {
    if (source.length === 2) {
      source.push("", getDepartmentTypeLabel(source[1]));
    } else if (source.length === 3) {
      source.splice(2, 0, "");
    }
  }

  if (source.length < baseSection.columns.length) {
    while (source.length < baseSection.columns.length) {
      source.push("");
    }
  } else if (source.length > baseSection.columns.length) {
    source.length = baseSection.columns.length;
  }

  return source;
}

function isAppVisible() {
  return !appSection.classList.contains("hidden");
}

function triggerAddRowHotkey() {
  const addButton = document.getElementById("addRowBtn");
  if (addButton) {
    addButton.click();
  }
}

function openFilterAndFocusSearch() {
  if (filterPanelOpenBySection[activeSectionId] !== true) {
    filterPanelOpenBySection[activeSectionId] = true;
    renderTablePreserveScroll();
  }

  const searchInput = document.getElementById("filterSearch");
  if (searchInput) {
    searchInput.focus();
    searchInput.select?.();
  }
}

function toggleFilterWithHotkey() {
  const isOpen = filterPanelOpenBySection[activeSectionId] === true;
  if (isOpen) {
    filterPanelOpenBySection[activeSectionId] = false;
    renderTablePreserveScroll();
    return;
  }
  openFilterAndFocusSearch();
}

function toggleSidebarCollapse() {
  isSidebarCollapsed = !isSidebarCollapsed;
  document.body.classList.toggle("sidebar-collapsed", isSidebarCollapsed);
  if (sidebarBrandToggle) {
    sidebarBrandToggle.title = isSidebarCollapsed ? "Развернуть меню" : "Свернуть меню";
    sidebarBrandToggle.setAttribute("aria-expanded", isSidebarCollapsed ? "false" : "true");
  }
  renderSidebarMenu();
  initLucideIcons();
}

function registerHotkeys() {
  document.addEventListener("keydown", (event) => {
    if (!isAppVisible()) return;
    if (event.key === "Insert") {
      event.preventDefault();
      event.stopPropagation();
      triggerAddRowHotkey();
      return;
    }
    if (!(event.ctrlKey || event.metaKey)) return;

    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      triggerAddRowHotkey();
      return;
    }

    // code=KeyF не зависит от раскладки клавиатуры.
    if (event.code === "KeyF" || event.key.toLowerCase() === "f") {
      event.preventDefault();
      event.stopPropagation();
      toggleFilterWithHotkey();
      return;
    }

    if (event.code === "KeyB" || event.key.toLowerCase() === "b") {
      event.preventDefault();
      event.stopPropagation();
      toggleSidebarCollapse();
    }
  }, true);
}

function saveSession(userName) {
  localStorage.setItem(SESSION_STORAGE_KEY, userName);
}

function clearSession() {
  localStorage.removeItem(SESSION_STORAGE_KEY);
  setAuthToken("");
  clearLastSessionActivityTs();
  stopRemoteAutoPull();
  hasUnsyncedLocalChanges = false;
  serverPushInFlight = false;
  stopSessionIdleWatcher();
}

function restoreSession() {
  return localStorage.getItem(SESSION_STORAGE_KEY);
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(loginForm);
  const phone = normalizeUzPhone(String(formData.get("phone") || ""));
  const password = String(formData.get("password") || "");
  if (!isPhoneLengthValid(phone)) {
    loginError.textContent = `Введите корректный номер (${getPhoneLengthHint(phone)} цифр после + для выбранной страны).`;
    loginError.classList.remove("hidden");
    return;
  }
  loginError.textContent = "Неверный логин или пароль";

  if (isHostedRuntime()) {
    try {
      const r = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ phone, password })
      });
      if (r.ok) {
        const j = await r.json();
        setAuthToken(j.token);
        loginError.classList.add("hidden");
        const me = await refreshAuthMeProfile();
        const userName = String(me?.displayName || j.displayName || "").trim() || "Пользователь";
        saveSession(userName);
        await pullRemoteAppState();
        showApp(userName);
        scheduleServerSync();
        return;
      }
    } catch (_) {
      /* offline or server error — пробуем локальный режим ниже */
    }
  }

  const isDefaultCredentials = phone === normalizeUzPhone(AUTH_PHONE) && password === AUTH_PASSWORD;

  if (isDefaultCredentials) {
    loginError.classList.add("hidden");
    const userName = "Пользователь";
    currentAuthRole = "admin";
    saveSession(userName);
    showApp(userName);
    return;
  }

  loginError.classList.remove("hidden");
});

logoutBtn.addEventListener("click", () => {
  confirmAction({
    message: "Выйти из системы?",
    confirmLabel: "Да",
    onConfirm: () => {
      clearSession();
      showLogin();
    }
  });
});

togglePasswordBtn?.addEventListener("click", togglePasswordVisibility);
phoneInput?.addEventListener("focus", () => {
  if (!phoneInput.value.trim()) {
    phoneInput.value = DEFAULT_PHONE_PREFIX;
  }
  updateLoginPhoneFlag();
  setCaretAfterDialCode(phoneInput);
});
phoneInput?.addEventListener("blur", enforceUzPhonePrefix);
if (phoneInput instanceof HTMLInputElement) {
  attachStrictPhoneInputBehavior(
    phoneInput,
    updateLoginPhoneFlag,
    () => {
      if (passwordInput instanceof HTMLInputElement) {
        requestAnimationFrame(() => passwordInput.focus());
      }
    }
  );
}
loginPhoneCountryBtn?.addEventListener("click", openLoginCountryPickerModal);

loginBtn.innerHTML = withLucideIcon("log-in", "Войти");
logoutBtn.innerHTML = withLucideIcon("log-out", "Выйти");
restoreDisplaySettings();
restoreSectionsData();
restoreTaskMultiState();
cleanupPendingImportedTaskIds({ save: false });
applyObjectsSeedIfNeeded();
loadObjectPhotoThumbsFromStorage();
ensureSystemRoles();
ensureSystemDepartments();
normalizePhaseAndSectionCatalogs();
restoreTrashData();
clearTasksTrashNow({ save: true });
registerHotkeys();
activeSectionId = restoreActiveSection();
isSettingsOpen = activeSectionId !== "tasks" && activeSectionId !== "report";
sidebarBrandToggle?.addEventListener("click", () => toggleSidebarCollapse());
sidebarBrandToggle?.setAttribute("aria-expanded", String(!isSidebarCollapsed));
initLucideIcons();
document.body.classList.add("login-mode");
updateLoginPhoneFlag();
window.addEventListener("resize", () => {
  updateTableStickyHeaderOffsets();
});

const initialShareIdBoot = new URLSearchParams(location.search).get("share");
if (initialShareIdBoot) {
  loginSection.classList.add("hidden");
  appSection.classList.add("hidden");
  document.body.classList.remove("login-mode");
  mountReportShareGate(initialShareIdBoot);
} else {
  (async () => {
    const savedUser = restoreSession();
    const hasToken = Boolean(getAuthToken());
    let userName = String(savedUser || "").trim();
    if (savedUser && hasToken && isHostedRuntime()) {
      try {
        const me = await refreshAuthMeProfile();
        if (me?.displayName) {
          userName = me.displayName;
          saveSession(userName);
        }
        await pullRemoteAppState();
      } catch (_) {
        /* noop */
      }
    }
    if (savedUser && hasToken) {
      showApp(userName || savedUser);
      if (getAuthToken() && isHostedRuntime()) {
        scheduleServerSync();
      }
    } else {
      showLogin();
    }
  })();
}

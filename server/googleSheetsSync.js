import crypto from "crypto";

const TASK_COLUMNS = {
  object: 1
};

let autoSyncTimer = null;
let syncInFlight = false;

function getEnv(name) {
  return String(process.env[name] || "").trim();
}

function normalizePrivateKey(raw) {
  return String(raw || "").replace(/\\n/g, "\n").trim();
}

function clampInt(v, min, max, fallback) {
  const n = Number(v);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, Math.floor(n)));
}

function getGoogleSheetsEnvConfig() {
  const clientEmail = getEnv("GOOGLE_SHEETS_CLIENT_EMAIL");
  const privateKey = normalizePrivateKey(process.env.GOOGLE_SHEETS_PRIVATE_KEY || "");
  return {
    clientEmail,
    privateKey,
    ready: Boolean(clientEmail && privateKey)
  };
}

function base64Url(input) {
  return Buffer.from(input)
    .toString("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

async function getGoogleAccessToken({ clientEmail, privateKey }) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600
  };
  const encodedHeader = base64Url(JSON.stringify(header));
  const encodedPayload = base64Url(JSON.stringify(payload));
  const unsigned = `${encodedHeader}.${encodedPayload}`;
  const signature = crypto.createSign("RSA-SHA256").update(unsigned).sign(privateKey);
  const assertion = `${unsigned}.${base64Url(signature)}`;

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion
  });
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) {
    throw new Error(String(j.error_description || j.error || "google_token_failed"));
  }
  return String(j.access_token);
}

function normalizeSheetTitle(raw, fallback = "Без объекта") {
  const cleaned = String(raw || "")
    .replace(/[\[\]\*\/\\\?\:]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const base = cleaned || fallback;
  return base.slice(0, 100).trim() || fallback;
}

function normalizeSheetTitleKey(raw) {
  return normalizeSheetTitle(raw)
    .normalize("NFKC")
    .toLocaleLowerCase("ru")
    .trim();
}

function makeUniqueSheetTitle(baseTitle, used) {
  const base = normalizeSheetTitle(baseTitle);
  if (!used.has(base)) {
    used.add(base);
    return base;
  }
  let idx = 2;
  while (idx < 9999) {
    const suffix = ` (${idx})`;
    const stem = base.slice(0, Math.max(1, 100 - suffix.length)).trim();
    const next = `${stem}${suffix}`;
    if (!used.has(next)) {
      used.add(next);
      return next;
    }
    idx += 1;
  }
  const fallback = `${base.slice(0, 90)} ${Date.now()}`.slice(0, 100);
  used.add(fallback);
  return fallback;
}

function sheetRangeA1(title) {
  const escaped = String(title || "").replace(/'/g, "''");
  return `'${escaped}'!A1`;
}

function normalizeCell(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\r\n?/g, "\n");
}

function buildSheetsPayload(payload, settings) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  const tasks = sections.find((s) => s?.id === "tasks");
  const columns = Array.isArray(tasks?.columns) ? tasks.columns.map((c) => normalizeCell(c)) : [];
  const rows = Array.isArray(tasks?.rows) ? tasks.rows : [];
  const summaryRows = [columns, ...rows.map((r) => (Array.isArray(r) ? r.map(normalizeCell) : []))];

  const includeByObject = settings.includeObjectSheets;
  const objectSheets = [];
  if (includeByObject) {
    const grouped = new Map();
    for (const row of rows) {
      const obj = normalizeCell(Array.isArray(row) ? row[TASK_COLUMNS.object] : "").trim() || "Без объекта";
      if (!grouped.has(obj)) grouped.set(obj, []);
      grouped.get(obj).push(Array.isArray(row) ? row.map(normalizeCell) : []);
    }
    const usedTitles = new Set([settings.summarySheetName]);
    for (const [objectName, objectRows] of grouped.entries()) {
      const sheetTitle = makeUniqueSheetTitle(objectName, usedTitles);
      objectSheets.push({
        title: sheetTitle,
        values: [columns, ...objectRows]
      });
    }
  }

  return {
    summary: {
      title: settings.summarySheetName,
      values: summaryRows
    },
    objectSheets
  };
}

async function sheetsApi(spreadsheetId, accessToken, path = "", options = {}) {
  const method = options.method || "GET";
  const query = options.query || "";
  const body = options.body;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}${path}${query ? `?${query}` : ""}`;
  const r = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(body ? { "Content-Type": "application/json" } : {})
    },
    body: body ? JSON.stringify(body) : undefined
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) {
    const m = j?.error?.message || j?.error_description || j?.error || r.statusText || "google_sheets_api_failed";
    throw new Error(String(m));
  }
  return j;
}

async function getSpreadsheetSheetTitles(spreadsheetId, accessToken) {
  const data = await sheetsApi(spreadsheetId, accessToken, "", {
    query: "fields=sheets.properties.title"
  });
  const sheets = Array.isArray(data?.sheets) ? data.sheets : [];
  return sheets.map((s) => String(s?.properties?.title || "").trim()).filter(Boolean);
}

async function ensureSheetsExist(spreadsheetId, accessToken, titles) {
  if (!titles.length) return;
  const existing = await getSpreadsheetSheetTitles(spreadsheetId, accessToken);
  const existingKeySet = new Set(existing.map((t) => normalizeSheetTitleKey(t)).filter(Boolean));
  const queue = [];
  const queuedKeySet = new Set();
  for (const title of titles) {
    const key = normalizeSheetTitleKey(title);
    if (!key || existingKeySet.has(key) || queuedKeySet.has(key)) continue;
    queue.push(normalizeSheetTitle(title));
    queuedKeySet.add(key);
  }
  const missing = queue;
  if (!missing.length) return;
  try {
    await sheetsApi(spreadsheetId, accessToken, ":batchUpdate", {
      method: "POST",
      body: {
        requests: missing.map((title) => ({ addSheet: { properties: { title } } }))
      }
    });
  } catch (error) {
    const msg = String(error?.message || error || "");
    if (!/already exists/i.test(msg)) throw error;
    for (const title of missing) {
      try {
        await sheetsApi(spreadsheetId, accessToken, ":batchUpdate", {
          method: "POST",
          body: {
            requests: [{ addSheet: { properties: { title } } }]
          }
        });
      } catch (singleError) {
        const singleMsg = String(singleError?.message || singleError || "");
        if (!/already exists/i.test(singleMsg)) throw singleError;
      }
    }
  }
}

async function writeSheetValues(spreadsheetId, accessToken, title, values) {
  const range = sheetRangeA1(title);
  await sheetsApi(spreadsheetId, accessToken, `/values/${encodeURIComponent(range)}:clear`, {
    method: "POST",
    body: {}
  });
  await sheetsApi(spreadsheetId, accessToken, `/values/${encodeURIComponent(range)}`, {
    method: "PUT",
    query: "valueInputOption=RAW",
    body: {
      range,
      majorDimension: "ROWS",
      values: Array.isArray(values) && values.length ? values : [[""]]
    }
  });
}

async function loadPayload(pool) {
  const { rows } = await pool.query("SELECT payload FROM app_state WHERE id = 1");
  const raw = rows[0]?.payload;
  return raw && typeof raw === "object" ? JSON.parse(JSON.stringify(raw)) : {};
}

async function saveSyncStateOnly(pool, patch = {}) {
  const status = String(patch.status || "").trim();
  const at = String(patch.at || "").trim();
  const atMs = Number(patch.atMs) || 0;
  const message = String(patch.message || "").trim();
  const rows = Number(patch.rows) || 0;
  const mode = String(patch.mode || "").trim();
  await pool.query(
    `INSERT INTO app_state (id, payload, updated_at)
     VALUES (
       1,
       jsonb_build_object(
         'displaySettings',
         jsonb_build_object(
           'googleSheetsLastSyncStatus', $1::text,
           'googleSheetsLastSyncAt', $2::text,
           'googleSheetsLastSyncAtMs', $3::bigint,
           'googleSheetsLastSyncMessage', $4::text,
           'googleSheetsLastSyncRows', $5::int,
           'googleSheetsLastSyncMode', $6::text
         )
       ),
       NOW()
     )
     ON CONFLICT (id) DO UPDATE SET
       payload = jsonb_set(
         COALESCE(app_state.payload, '{}'::jsonb),
         '{displaySettings}',
         COALESCE(app_state.payload->'displaySettings', '{}'::jsonb) || jsonb_build_object(
           'googleSheetsLastSyncStatus', $1::text,
           'googleSheetsLastSyncAt', $2::text,
           'googleSheetsLastSyncAtMs', $3::bigint,
           'googleSheetsLastSyncMessage', $4::text,
           'googleSheetsLastSyncRows', $5::int,
           'googleSheetsLastSyncMode', $6::text
         ),
         true
       ),
       updated_at = NOW()`,
    [status, at, atMs, message, rows, mode]
  );
}

function readSyncSettings(payload) {
  const ds = payload?.displaySettings || {};
  const spreadsheetId = String(ds.googleSheetsSpreadsheetId || "").trim();
  const summarySheetName = normalizeSheetTitle(ds.googleSheetsSummarySheetName || "Сводная", "Сводная");
  const enabled = Boolean(ds.googleSheetsEnabled);
  const autoEnabled = Boolean(ds.googleSheetsAutoSyncEnabled);
  const includeObjectSheets = ds.googleSheetsIncludeObjectSheets !== false;
  const intervalMinutes = clampInt(ds.googleSheetsSyncIntervalMinutes, 1, 1440, 30);
  return {
    enabled,
    autoEnabled,
    includeObjectSheets,
    spreadsheetId,
    summarySheetName,
    intervalMinutes
  };
}

function setSyncState(payload, patch = {}) {
  if (!payload.displaySettings || typeof payload.displaySettings !== "object") payload.displaySettings = {};
  payload.displaySettings.googleSheetsLastSyncStatus = String(patch.status || "").trim();
  payload.displaySettings.googleSheetsLastSyncAt = String(patch.at || "").trim();
  payload.displaySettings.googleSheetsLastSyncAtMs = Number(patch.atMs) || 0;
  payload.displaySettings.googleSheetsLastSyncMessage = String(patch.message || "").trim();
  payload.displaySettings.googleSheetsLastSyncRows = Number(patch.rows) || 0;
  payload.displaySettings.googleSheetsLastSyncMode = String(patch.mode || "").trim();
}

function canRunAuto(settings, payload, nowMs) {
  if (!settings.enabled || !settings.autoEnabled || !settings.spreadsheetId) return false;
  const ds = payload?.displaySettings || {};
  const lastAtMs = Number(ds.googleSheetsLastSyncAtMs) || 0;
  if (!lastAtMs) return true;
  const diff = nowMs - lastAtMs;
  return diff >= settings.intervalMinutes * 60 * 1000;
}

export async function runGoogleSheetsSync(pool, options = {}) {
  const mode = String(options.mode || "manual");
  if (syncInFlight) {
    return { ok: false, busy: true, error: "sync_in_progress" };
  }
  syncInFlight = true;
  const started = Date.now();
  let payload = null;
  try {
    payload = await loadPayload(pool);
    const settings = readSyncSettings(payload);
    if (!settings.enabled) {
      throw new Error("Синхронизация Google Sheets отключена в настройках.");
    }
    if (!settings.spreadsheetId) {
      throw new Error("Не заполнен Spreadsheet ID.");
    }
    const env = getGoogleSheetsEnvConfig();
    if (!env.ready) {
      throw new Error("Не заполнены GOOGLE_SHEETS_CLIENT_EMAIL / GOOGLE_SHEETS_PRIVATE_KEY.");
    }

    const accessToken = await getGoogleAccessToken(env);
    const data = buildSheetsPayload(payload, settings);
    const allTitles = [data.summary.title, ...data.objectSheets.map((s) => s.title)];
    await ensureSheetsExist(settings.spreadsheetId, accessToken, allTitles);
    await writeSheetValues(settings.spreadsheetId, accessToken, data.summary.title, data.summary.values);
    for (const sh of data.objectSheets) {
      await writeSheetValues(settings.spreadsheetId, accessToken, sh.title, sh.values);
    }

    const rowsCount = Math.max(0, data.summary.values.length - 1);
    const at = new Date();
    const syncPatch = {
      status: "ok",
      at: at.toISOString(),
      atMs: at.getTime(),
      message: `Успешно: ${rowsCount} задач, листов по объектам: ${data.objectSheets.length}.`,
      rows: rowsCount,
      mode
    };
    setSyncState(payload, syncPatch);
    await saveSyncStateOnly(pool, syncPatch);
    return {
      ok: true,
      rows: rowsCount,
      objectSheets: data.objectSheets.length,
      durationMs: Date.now() - started,
      message: payload.displaySettings.googleSheetsLastSyncMessage
    };
  } catch (e) {
    const at = new Date();
    const msg = String(e?.message || e || "google_sync_failed");
    if (payload) {
      const syncPatch = {
        status: "error",
        at: at.toISOString(),
        atMs: at.getTime(),
        message: msg,
        rows: 0,
        mode
      };
      setSyncState(payload, syncPatch);
      try {
        await saveSyncStateOnly(pool, syncPatch);
      } catch (_) {
        /* noop */
      }
    }
    return { ok: false, error: msg, durationMs: Date.now() - started };
  } finally {
    syncInFlight = false;
  }
}

export function startGoogleSheetsAutoSync(pool) {
  if (autoSyncTimer) clearInterval(autoSyncTimer);
  autoSyncTimer = setInterval(async () => {
    if (syncInFlight) return;
    try {
      const payload = await loadPayload(pool);
      const settings = readSyncSettings(payload);
      if (!canRunAuto(settings, payload, Date.now())) return;
      await runGoogleSheetsSync(pool, { mode: "auto" });
    } catch (_) {
      /* noop */
    }
  }, 30000);
}

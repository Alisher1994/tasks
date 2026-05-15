function getPayloadSection(payload, sectionId) {
  const sections = Array.isArray(payload?.sections) ? payload.sections : [];
  return sections.find((section) => section && String(section.id || "") === sectionId) || null;
}

function sectionRowsAsObjects(section) {
  const columns = Array.isArray(section?.columns) ? section.columns.map((x) => String(x || "")) : [];
  const rows = Array.isArray(section?.rows) ? section.rows : [];
  return rows.map((row) => {
    const out = {};
    columns.forEach((column, index) => {
      out[column || `col_${index}`] = Array.isArray(row) ? row[index] ?? "" : "";
    });
    return out;
  });
}

export function buildExportSectionPayload(payload, sectionId) {
  const section = getPayloadSection(payload, sectionId);
  if (!section) return null;
  return {
    id: String(section.id || sectionId),
    title: String(section.title || sectionId),
    columns: Array.isArray(section.columns) ? section.columns : [],
    rows: Array.isArray(section.rows) ? section.rows : [],
    items: sectionRowsAsObjects(section)
  };
}

export function buildCatalogsExportPayload(payload) {
  const ids = ["data", "phases", "phaseSections", "phaseSubsections", "delayReasons", "roles", "departments"];
  return Object.fromEntries(
    ids
      .map((id) => [id, buildExportSectionPayload(payload, id)])
      .filter(([, value]) => Boolean(value))
  );
}

export function sanitizeDisplaySettingsForExport(settings) {
  const out = settings && typeof settings === "object" && !Array.isArray(settings)
    ? JSON.parse(JSON.stringify(settings))
    : {};
  [
    "telegramBotToken",
    "smsGatewayPassword",
    "smsGatewayApiKey",
    "googleSheetsSpreadsheetId"
  ].forEach((key) => {
    if (key in out) out[key] = "";
  });
  return out;
}

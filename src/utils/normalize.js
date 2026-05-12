export function normalizeDateDisplayFormatId(value) {
  const id = String(value || "");
  return ["DMY_DOT", "ISO", "DMY_SLASH", "MDY_SLASH"].includes(id) ? id : "DMY_DOT";
}

export function normalizeTimeDisplayFormatId(value) {
  return String(value) === "12" ? "12" : "24";
}

export function normalizeGoogleSheetsInterval(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 30;
  return Math.min(1440, Math.max(1, Math.floor(n)));
}

export function normalizeGoogleSheetsWriteMode(value) {
  return String(value || "").trim() === "update" ? "update" : "rewrite";
}

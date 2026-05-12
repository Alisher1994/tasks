export function isPlainObjectRecord(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

export function cloneJsonSafe(value) {
  return JSON.parse(JSON.stringify(value));
}

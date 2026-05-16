export function parseRuDateStringToParts(value) {
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

export function formatDatePartsStorage(day, month, year) {
  return `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`;
}

export function calendarDiffDays(fromParts, toParts) {
  const a = Date.UTC(fromParts.year, fromParts.month - 1, fromParts.day);
  const b = Date.UTC(toParts.year, toParts.month - 1, toParts.day);
  return Math.round((b - a) / 86400000);
}

export function parseRuDate(value) {
  const parts = parseRuDateStringToParts(value);
  if (!parts) return null;
  const date = new Date(parts.year, parts.month - 1, parts.day);
  return Number.isNaN(date.getTime()) ? null : date;
}

export function getCalendarDatePartsInTimeZone(date, timeZone) {
  const d = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(d.getTime())) return null;
  const tz = timeZone || "UTC";
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

/**
 * Разбор даты-времени в русском формате «ДД.ММ.ГГГГ[ ЧЧ:ММ[:СС]]» в ms.
 * @returns {number} ms либо NaN, если не удалось разобрать.
 */
export function parseRuDateTimeToMs(value) {
  const raw = String(value || "").trim();
  if (!raw) return NaN;
  const m = raw.match(/(\d{1,2}\.\d{1,2}\.\d{2,4})(?:[ ,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return NaN;
  const parts = parseRuDateStringToParts(m[1]);
  if (!parts) return NaN;
  const hh = m[2] != null ? Number(m[2]) : 0;
  const mm = m[3] != null ? Number(m[3]) : 0;
  const ss = m[4] != null ? Number(m[4]) : 0;
  if (hh > 23 || mm > 59 || ss > 59) return NaN;
  const ms = new Date(parts.year, parts.month - 1, parts.day, hh, mm, ss).getTime();
  return Number.isFinite(ms) ? ms : NaN;
}

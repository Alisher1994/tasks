import { escapeHtmlText } from "./escape.js";

export function shortenHistorySnippet(s, maxLen = 100) {
  const t = String(s ?? "");
  if (t.length <= maxLen) return t;
  return `${t.slice(0, maxLen)}…`;
}

export function normalizePersonName(s) {
  return String(s ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

export function isGenericSystemUserName(value) {
  const name = String(value || "").trim().toLowerCase();
  return name === "пользователь" || name === "администратор";
}

export function formatTelegramPreviewHtml(text) {
  const t = String(text ?? "").trim() || "—";
  return escapeHtmlText(t).replace(/\r\n|\r|\n/g, "<br />");
}

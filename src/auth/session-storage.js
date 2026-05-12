import { SESSION_STORAGE_KEY, SESSION_PHONE_STORAGE_KEY } from "../constants/storage-keys.js";
import { normalizeUzPhone } from "../utils/phone.js";

export function getSessionUserDisplayName() {
  return String(localStorage.getItem(SESSION_STORAGE_KEY) || "").trim() || "Пользователь";
}

export function saveSessionPhone(phone) {
  try {
    const normalized = normalizeUzPhone(String(phone || ""));
    if (!normalized) {
      localStorage.removeItem(SESSION_PHONE_STORAGE_KEY);
      return;
    }
    localStorage.setItem(SESSION_PHONE_STORAGE_KEY, normalized);
  } catch (_) {
    /* noop */
  }
}

export function getSessionPhone() {
  try {
    return normalizeUzPhone(String(localStorage.getItem(SESSION_PHONE_STORAGE_KEY) || ""));
  } catch (_) {
    return "";
  }
}

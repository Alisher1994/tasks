import { promises as fsp } from "fs";

const MEDIA_UPLOAD_MAX_BYTES = 18 * 1024 * 1024;

export async function ensureMediaStorageDir(mediaStoragePath) {
  await fsp.mkdir(mediaStoragePath, { recursive: true });
}

function mimeToExt(mime) {
  const m = String(mime || "").toLowerCase();
  if (m === "image/jpeg") return "jpg";
  if (m === "image/png") return "png";
  if (m === "image/webp") return "webp";
  if (m === "image/gif") return "gif";
  if (m === "image/bmp") return "bmp";
  if (m === "video/mp4") return "mp4";
  if (m === "video/webm") return "webm";
  if (m === "video/ogg") return "ogv";
  if (m === "video/quicktime") return "mov";
  if (m === "video/x-m4v") return "m4v";
  if (m === "application/pdf") return "pdf";
  return "bin";
}

function extFromFileName(fileName) {
  const src = String(fileName || "").trim();
  if (!src) return "";
  const clean = src.replace(/[?#].*$/, "");
  const dot = clean.lastIndexOf(".");
  if (dot <= 0 || dot >= clean.length - 1) return "";
  const ext = clean.slice(dot + 1).toLowerCase().replace(/[^a-z0-9]/g, "");
  if (!ext) return "";
  const allow = new Set([
    "jpg", "jpeg", "png", "webp", "gif", "bmp",
    "mp4", "webm", "ogv", "mov", "m4v",
    "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx",
    "txt", "csv", "zip", "rar", "7z"
  ]);
  return allow.has(ext) ? ext : "";
}

function parseDataUrl(dataUrl) {
  const m = /^data:([^;]+);base64,(.+)$/s.exec(String(dataUrl || "").trim());
  if (!m) return null;
  return { mime: String(m[1] || "").trim(), base64: String(m[2] || "").trim() };
}

function decodeUploadBase64(base64) {
  const compact = String(base64 || "").replace(/\s+/g, "");
  if (!compact || !/^[A-Za-z0-9+/]+={0,2}$/.test(compact) || compact.length % 4 !== 0) {
    return null;
  }
  return Buffer.from(compact, "base64");
}

export function validateMediaUpload({ dataUrl, fileName }) {
  const parsed = parseDataUrl(dataUrl);
  if (!parsed) {
    return { ok: false, status: 400, error: "Неверный формат файла (ожидается data URL)." };
  }
  if (String(parsed.mime || "").toLowerCase() === "image/svg+xml") {
    return { ok: false, status: 415, error: "SVG-файлы не принимаются по соображениям безопасности." };
  }
  if (parsed.base64.length > Math.ceil(MEDIA_UPLOAD_MAX_BYTES * 4 / 3) + 4) {
    return { ok: false, status: 413, error: "Файл слишком большой (максимум ~18MB)." };
  }
  const buf = decodeUploadBase64(parsed.base64);
  if (!buf) {
    return { ok: false, status: 400, error: "Файл повреждён или не является корректным base64." };
  }
  if (buf.length > MEDIA_UPLOAD_MAX_BYTES) {
    return { ok: false, status: 413, error: "Файл слишком большой (максимум ~18MB)." };
  }
  const ext = extFromFileName(fileName) || mimeToExt(parsed.mime);
  if (!ext || ext === "bin") {
    return { ok: false, status: 415, error: "Тип файла не поддерживается." };
  }
  return { ok: true, buf, ext, mime: parsed.mime };
}

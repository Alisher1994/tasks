/**
 * Миниатюрная аватарка по имени: кружок с инициалами и детерминированным
 * цветом (один и тот же человек → всегда один цвет). Реального фото в данных
 * нет, поэтому генерируем из имени. Чистый модуль без зависимостей.
 *
 * Стили инлайновые — функция используется в самых разных местах рендера
 * (ячейки таблицы, аналитика, аккаунт-блок), и инлайн исключает зависимость
 * от того, подключён ли нужный CSS-класс в конкретном контексте.
 */

const PALETTE = [
  "#5b8def", "#e8704f", "#3fa372", "#9b6dd6", "#d9a441",
  "#4aa3c7", "#d0608f", "#6a8c3a", "#c2553b", "#7b6cc4"
];

function hashString(str) {
  let h = 0;
  for (let i = 0; i < str.length; i += 1) {
    h = (h << 5) - h + str.charCodeAt(i);
    h |= 0;
  }
  return Math.abs(h);
}

function escapeAttr(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
  ));
}

/** Инициалы: до двух букв из первых двух значимых слов имени. */
function initialsOf(name) {
  const words = String(name || "")
    .trim()
    .split(/\s+/)
    .filter((w) => /[A-Za-zА-Яа-яЁё0-9]/.test(w));
  if (!words.length) return "—";
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return (words[0][0] + words[1][0]).toUpperCase();
}

/**
 * HTML миниатюрной аватарки.
 * @param {string} name полное имя
 * @param {{ size?: number }} [opts] размер кружка в px (по умолчанию 18)
 * @returns {string} безопасный HTML (имя экранируется)
 */
export function avatarHtml(name, opts = {}) {
  const size = Number(opts.size) > 0 ? Math.floor(Number(opts.size)) : 18;
  const clean = String(name || "").trim();
  const initials = initialsOf(clean);
  const bg = clean ? PALETTE[hashString(clean) % PALETTE.length] : "#9aa3ad";
  const fontSize = Math.max(8, Math.round(size * 0.42));
  const style = [
    `display:inline-flex`,
    `align-items:center`,
    `justify-content:center`,
    `width:${size}px`,
    `height:${size}px`,
    `min-width:${size}px`,
    `border-radius:50%`,
    `background:${bg}`,
    `color:#fff`,
    `font-size:${fontSize}px`,
    `font-weight:600`,
    `line-height:1`,
    `vertical-align:middle`,
    `user-select:none`,
    `overflow:hidden`
  ].join(";");
  return `<span class="mbc-avatar" title="${escapeAttr(clean)}" aria-hidden="true" style="${style}">${escapeAttr(initials)}</span>`;
}

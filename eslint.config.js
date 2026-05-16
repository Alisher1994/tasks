// Минимальная, высокосигнальная конфигурация ESLint (flat config, ESLint 9).
// Цель сейчас — не «причесать стиль», а ловить РЕАЛЬНЫЕ баги, особенно класс
// «сломанная ссылка при рефакторинге» (no-undef). Шумные стилевые правила
// намеренно ослаблены, чтобы вывод был полезным, а не стеной из тысяч замечаний.

const browserGlobals = {
  window: "readonly",
  document: "readonly",
  navigator: "readonly",
  localStorage: "readonly",
  sessionStorage: "readonly",
  location: "readonly",
  fetch: "readonly",
  setTimeout: "readonly",
  clearTimeout: "readonly",
  setInterval: "readonly",
  clearInterval: "readonly",
  requestAnimationFrame: "readonly",
  cancelAnimationFrame: "readonly",
  console: "readonly",
  alert: "readonly",
  confirm: "readonly",
  prompt: "readonly",
  FormData: "readonly",
  Blob: "readonly",
  File: "readonly",
  FileReader: "readonly",
  URL: "readonly",
  URLSearchParams: "readonly",
  WebSocket: "readonly",
  Image: "readonly",
  CustomEvent: "readonly",
  Event: "readonly",
  Element: "readonly",
  HTMLElement: "readonly",
  HTMLInputElement: "readonly",
  HTMLTextAreaElement: "readonly",
  HTMLSelectElement: "readonly",
  HTMLButtonElement: "readonly",
  HTMLCanvasElement: "readonly",
  Node: "readonly",
  AbortController: "readonly",
  IntersectionObserver: "readonly",
  ResizeObserver: "readonly",
  MutationObserver: "readonly",
  crypto: "readonly",
  history: "readonly",
  CSS: "readonly",
  structuredClone: "readonly",
  caches: "readonly",
  self: "readonly",
  // Внешние библиотеки, подключаемые через <script> (CDN):
  Chart: "readonly",
  ChartDataLabels: "readonly"
};

const nodeGlobals = {
  process: "readonly",
  console: "readonly",
  Buffer: "readonly",
  __dirname: "readonly",
  __filename: "readonly",
  URL: "readonly",
  URLSearchParams: "readonly",
  setTimeout: "readonly",
  clearTimeout: "readonly",
  setInterval: "readonly",
  clearInterval: "readonly",
  fetch: "readonly",
  globalThis: "readonly",
  Blob: "readonly",
  FormData: "readonly",
  AbortController: "readonly",
  Headers: "readonly",
  Response: "readonly",
  Request: "readonly",
  structuredClone: "readonly"
};

// Правила: только то, что реально предотвращает баги при рефакторинге.
const rules = {
  "no-undef": "error",          // главное: ловит несуществующие/опечатанные имена
  "no-redeclare": "error",
  "no-dupe-keys": "error",
  "no-dupe-args": "error",
  "no-unreachable": "error",
  "no-unsafe-negation": "error",
  "use-isnan": "error",
  "valid-typeof": "error",
  "no-unused-vars": ["warn", { args: "none", ignoreRestSiblings: true }],
  // Шум в большом легаси-коде — отключаем, чтобы вывод оставался полезным:
  "no-empty": "off",
  "no-constant-condition": "off",
  "no-cond-assign": "off",
  "no-control-regex": "off"
};

export default [
  { ignores: ["node_modules/**", "dist/**", "public/sw.js"] },
  {
    files: ["src/**/*.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: browserGlobals
    },
    rules
  },
  {
    files: ["server/**/*.js", "vite.config.js", "eslint.config.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: nodeGlobals
    },
    rules
  }
];

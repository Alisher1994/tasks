export const DEFAULT_PHONE_PREFIX = "+";
export const PHONE_MIN_DIGITS = 8;
export const PHONE_MAX_DIGITS = 15;
export const PHONE_MAX_LENGTH = PHONE_MAX_DIGITS + 1;

export function normalizeUzPhone(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw) return DEFAULT_PHONE_PREFIX;
  let s = raw.replace(/[^\d+]/g, "");
  if (s.startsWith("00")) s = `+${s.slice(2)}`;
  if (!s.startsWith("+")) s = `+${s.replace(/\D/g, "")}`;
  const digits = s.slice(1).replace(/\D/g, "").slice(0, PHONE_MAX_DIGITS);
  return `+${digits}`;
}

export const DIAL_TO_ISO = [
  ["998", "UZ"], ["7", "RU"], ["380", "UA"], ["375", "BY"], ["1", "US"],
  ["90", "TR"], ["971", "AE"], ["966", "SA"], ["44", "GB"], ["49", "DE"],
  ["33", "FR"], ["39", "IT"], ["34", "ES"], ["48", "PL"], ["995", "GE"],
  ["994", "AZ"], ["996", "KG"], ["992", "TJ"], ["993", "TM"], ["20", "EG"],
  ["91", "IN"], ["92", "PK"], ["86", "CN"], ["81", "JP"], ["82", "KR"],
  ["61", "AU"], ["55", "BR"], ["52", "MX"], ["62", "ID"], ["63", "PH"],
  ["65", "SG"], ["60", "MY"], ["66", "TH"], ["84", "VN"], ["98", "IR"]
];

export const COUNTRY_NAME_BY_ISO = {
  UZ: "Узбекистан", RU: "Россия", UA: "Украина", BY: "Беларусь", US: "США",
  TR: "Турция", AE: "ОАЭ", SA: "Саудовская Аравия", GB: "Великобритания", DE: "Германия",
  FR: "Франция", IT: "Италия", ES: "Испания", PL: "Польша", GE: "Грузия",
  AZ: "Азербайджан", KG: "Кыргызстан", TJ: "Таджикистан", TM: "Туркменистан", EG: "Египет",
  IN: "Индия", PK: "Пакистан", CN: "Китай", JP: "Япония", KR: "Южная Корея",
  AU: "Австралия", BR: "Бразилия", MX: "Мексика", ID: "Индонезия", PH: "Филиппины",
  SG: "Сингапур", MY: "Малайзия", TH: "Таиланд", VN: "Вьетнам", IR: "Иран"
};

/** Ограничения по длине в формате E.164 (общее количество цифр без "+"). */
export const PHONE_TOTAL_DIGITS_BY_DIAL = {
  "998": { min: 12, max: 12 },
  "7": { min: 11, max: 11 },
  "380": { min: 12, max: 12 },
  "375": { min: 12, max: 12 },
  "1": { min: 11, max: 11 },
  "90": { min: 12, max: 12 },
  "971": { min: 12, max: 12 },
  "966": { min: 12, max: 12 },
  "44": { min: 12, max: 12 },
  "49": { min: 11, max: 13 },
  "33": { min: 11, max: 11 },
  "39": { min: 11, max: 13 },
  "34": { min: 11, max: 11 },
  "48": { min: 11, max: 11 },
  "995": { min: 12, max: 12 },
  "994": { min: 12, max: 12 },
  "996": { min: 12, max: 12 },
  "992": { min: 12, max: 12 },
  "993": { min: 11, max: 11 },
  "20": { min: 12, max: 12 },
  "91": { min: 12, max: 12 },
  "92": { min: 12, max: 12 },
  "86": { min: 13, max: 13 },
  "81": { min: 11, max: 12 },
  "82": { min: 11, max: 12 },
  "61": { min: 11, max: 11 },
  "55": { min: 12, max: 13 },
  "52": { min: 12, max: 12 },
  "62": { min: 10, max: 13 },
  "63": { min: 12, max: 12 },
  "65": { min: 10, max: 10 },
  "60": { min: 10, max: 12 },
  "66": { min: 11, max: 11 },
  "84": { min: 11, max: 11 },
  "98": { min: 12, max: 12 }
};

export function flagEmojiFromIso2(iso2) {
  const code = String(iso2 || "").trim().toUpperCase();
  if (!/^[A-Z]{2}$/.test(code)) return "🌐";
  return String.fromCodePoint(...Array.from(code).map((c) => 127397 + c.charCodeAt(0)));
}

export function flagSvgUrlByIso(iso2) {
  const code = String(iso2 || "").trim().toLowerCase();
  if (!/^[a-z]{2}$/.test(code)) return "";
  return `https://flagcdn.com/${code}.svg`;
}

export function globeSvgDataUrl() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="14" viewBox="0 0 18 14"><rect width="18" height="14" rx="2" fill="#eef3f9"/><circle cx="9" cy="7" r="4.2" fill="none" stroke="#6b7c93" stroke-width="1"/><path d="M4.8 7h8.4M9 2.8c1.6 1.1 1.6 7.3 0 8.4M9 2.8c-1.6 1.1-1.6 7.3 0 8.4" stroke="#6b7c93" stroke-width="1" fill="none"/></svg>`;
  return `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;
}

export function detectCountryIsoByPhone(rawPhone) {
  const digits = normalizeUzPhone(rawPhone).replace(/\D/g, "");
  if (!digits) return "";
  let best = "";
  let bestIso = "";
  for (const [dial, iso] of DIAL_TO_ISO) {
    if (digits.startsWith(dial) && dial.length > best.length) {
      best = dial;
      bestIso = iso;
    }
  }
  return bestIso;
}

export function phoneFlagByValue(rawPhone) {
  const iso = detectCountryIsoByPhone(rawPhone);
  return flagEmojiFromIso2(iso);
}

export function getPhoneDigitsCount(rawPhone) {
  return normalizeUzPhone(rawPhone).replace(/\D/g, "").length;
}

export function detectDialCodeByPhone(rawPhone) {
  const digits = normalizeUzPhone(rawPhone).replace(/\D/g, "");
  if (!digits) return "";
  let best = "";
  for (const [dial] of DIAL_TO_ISO) {
    if (digits.startsWith(dial) && dial.length > best.length) {
      best = dial;
    }
  }
  return best;
}

export function getPhoneRuleByDial(dial) {
  const d = String(dial || "").trim();
  return PHONE_TOTAL_DIGITS_BY_DIAL[d] || { min: PHONE_MIN_DIGITS, max: PHONE_MAX_DIGITS };
}

export function getPhoneLengthHint(rawPhone) {
  const dial = detectDialCodeByPhone(rawPhone);
  const rule = getPhoneRuleByDial(dial);
  return rule.min === rule.max ? `${rule.max}` : `${rule.min}-${rule.max}`;
}

export function isPhoneLengthValid(rawPhone) {
  const normalized = normalizeUzPhone(rawPhone);
  const len = normalized.replace(/\D/g, "").length;
  const dial = detectDialCodeByPhone(normalized);
  const rule = getPhoneRuleByDial(dial);
  return len >= rule.min && len <= rule.max;
}

export function buildCountryPhoneOptions() {
  return DIAL_TO_ISO
    .map(([dial, iso]) => ({
      dial,
      iso,
      flagUrl: flagSvgUrlByIso(iso),
      name: COUNTRY_NAME_BY_ISO[iso] || iso
    }))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
}

export function applyCountryDialToPhone(rawPhone, selectedDial) {
  const normalized = normalizeUzPhone(rawPhone);
  const digits = normalized.replace(/\D/g, "");
  const prevDial = detectDialCodeByPhone(normalized);
  const national = prevDial && digits.startsWith(prevDial) ? digits.slice(prevDial.length) : digits;
  return normalizeUzPhone(`+${selectedDial}${national}`);
}

export function sanitizePhoneInputValue(rawValue) {
  const n = normalizeUzPhone(rawValue);
  if (n === "+") return DEFAULT_PHONE_PREFIX;
  const digits = n.replace(/\D/g, "");
  const dial = detectDialCodeByPhone(n);
  const rule = getPhoneRuleByDial(dial);
  return `+${digits.slice(0, rule.max)}`;
}

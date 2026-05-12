export function formatSmsInviteHistoryStatusLabel(entry) {
  const status = String(entry?.status || "").trim();
  if (status === "sent" || entry?.ok === true) return "Отправлено";
  return "Ошибка";
}

export function shortenSmsGatewayId(rawId) {
  const id = String(rawId || "").trim();
  if (!id) return "";
  if (id.length <= 12) return id;
  return `${id.slice(0, 6)}...${id.slice(-4)}`;
}

export function mapSmsGatewayStateLabel(rawState) {
  const state = String(rawState || "").trim().toLowerCase();
  if (!state) return "";
  if (state === "pending" || state === "queued") return "в очереди";
  if (state === "sent") return "отправлено";
  if (state === "delivered") return "доставлено";
  if (state === "failed" || state === "error" || state === "rejected") return "ошибка";
  return String(rawState || "").trim();
}

export function summarizeSmsGatewayResponse(rawValue) {
  const full = String(rawValue || "").trim();
  if (!full) {
    return { shortText: "—", fullText: "—" };
  }
  if (/unauthorized|401/i.test(full)) {
    return { shortText: "Ошибка авторизации", fullText: full };
  }
  if (/Принято SMS Gate:/i.test(full)) {
    const stateMatch = full.match(/state=([^,\s]+)/i);
    const recipientsMatch = full.match(/recipients=([0-9]+)/i);
    const idMatch = full.match(/id=([^,\s]+)/i);
    const stateLabel = mapSmsGatewayStateLabel(stateMatch?.[1] || "");
    const recipients = recipientsMatch?.[1] ? Number(recipientsMatch[1]) : 0;
    const shortId = shortenSmsGatewayId(idMatch?.[1] || "");
    const shortText = [
      stateLabel ? `Принято: ${stateLabel}` : "Принято gateway",
      recipients > 0 ? `получателей: ${recipients}` : "",
      shortId ? `ID: ${shortId}` : ""
    ].filter(Boolean).join(" | ");
    return { shortText, fullText: full };
  }
  if (full.startsWith("{") || full.startsWith("[")) {
    try {
      const parsed = JSON.parse(full);
      const state = String(parsed?.state || parsed?.states?.[0]?.state || "").trim();
      const recipientsCount = Array.isArray(parsed?.recipients)
        ? parsed.recipients.length
        : Number(parsed?.recipientsCount || 0);
      const shortId = shortenSmsGatewayId(parsed?.id || "");
      const stateLabel = mapSmsGatewayStateLabel(state);
      const shortText = [
        stateLabel ? `Gateway: ${stateLabel}` : "Принято gateway",
        recipientsCount > 0 ? `получателей: ${recipientsCount}` : "",
        shortId ? `ID: ${shortId}` : ""
      ].filter(Boolean).join(" | ");
      return { shortText: shortText || "Принято gateway", fullText: full };
    } catch (_) {
      // fallback below
    }
  }
  if (full.length > 140) {
    return { shortText: `${full.slice(0, 137)}...`, fullText: full };
  }
  return { shortText: full, fullText: full };
}

export function summarizeSmsTextForHistory(rawValue) {
  const full = String(rawValue || "").trim();
  if (!full) {
    return { shortText: "—", fullText: "—" };
  }
  const oneLine = full.replace(/\s+/g, " ").trim();
  if (/регистрац/i.test(oneLine) && /t\.me\//i.test(oneLine)) {
    return {
      shortText: "Приглашение в Telegram-бот (ссылка на регистрацию).",
      fullText: full
    };
  }
  if (/задач[аеи]\s*№/i.test(oneLine) && /t\.me\//i.test(oneLine)) {
    return {
      shortText: "Уведомление по задаче (ссылка на Telegram-бот).",
      fullText: full
    };
  }
  if (oneLine.length > 160) {
    return { shortText: `${oneLine.slice(0, 157)}...`, fullText: full };
  }
  return { shortText: oneLine, fullText: full };
}

export function getSmsHistoryGatewayCategory(responseText, statusText = "") {
  const response = String(responseText || "").trim();
  const status = String(statusText || "").trim().toLowerCase();
  if (/unauthorized|401/i.test(response)) return "auth_error";
  if (status === "ошибка") return "error";
  if (status === "отправлено") return "success";
  if (/error|failed|rejected/i.test(response)) return "error";
  return "other";
}

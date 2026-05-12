import { copyTextToClipboard } from "../utils/clipboard.js";
import { initLucideIcons } from "./icons.js";

export function openLoginSupportModal() {
  const overlay = document.createElement("div");
  overlay.className = "responsible-modal-overlay login-support-modal-overlay";
  overlay.innerHTML = `
    <div class="responsible-modal login-support-modal" role="dialog" aria-modal="true" aria-label="Служба поддержки">
      <h3>Служба поддержки</h3>
      <div class="login-support-contact-card">
        <span class="login-support-contact-icon"><i data-lucide="user-round-cog" class="lucide-icon" aria-hidden="true"></i></span>
        <div>
          <div class="login-support-contact-title">Администратор системы</div>
          <strong>Алишер</strong>
          <div class="login-support-contact-row">
            <i data-lucide="phone" class="lucide-icon" aria-hidden="true"></i>
            <a href="tel:+998994067406">+998 99 406 74 06</a>
            <button type="button" class="login-support-copy-btn" data-copy-support-phone="+998994067406" title="Скопировать телефон" aria-label="Скопировать телефон">
              <i data-lucide="copy" class="lucide-icon" aria-hidden="true"></i>
            </button>
          </div>
          <div class="login-support-contact-row">
            <i data-lucide="send" class="lucide-icon" aria-hidden="true"></i>
            <a href="https://t.me/alishermusayev94" target="_blank" rel="noopener">Telegram: @alishermusayev94</a>
          </div>
        </div>
      </div>
      <div class="responsible-modal-actions">
        <button type="button" class="secondary login-support-close">Закрыть</button>
      </div>
    </div>
  `;
  const close = () => overlay.remove();
  overlay.querySelector(".login-support-close")?.addEventListener("click", close);
  overlay.querySelector("[data-copy-support-phone]")?.addEventListener("click", (event) => {
    copyTextToClipboard(event.currentTarget?.getAttribute("data-copy-support-phone") || "+998994067406");
    event.currentTarget?.classList.add("is-copied");
    window.setTimeout(() => event.currentTarget?.classList.remove("is-copied"), 900);
  });
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
  document.addEventListener("keydown", function onKey(event) {
    if (event.key !== "Escape") return;
    document.removeEventListener("keydown", onKey);
    close();
  });
  document.body.appendChild(overlay);
  initLucideIcons();
}

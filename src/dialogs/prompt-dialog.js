/* global Office */
import "./prompt-dialog.css";
import {initI18n, t} from "../i18n.js";

const ENV = (typeof window !== "undefined" && window.__ENV) || {};
const ENV_APP_NAME = ENV.APP_NAME || "AI Assistant";

const DEFAULT_PROMPT =
  "You are a professional email reply assistant. Based on the quoted email content provided, draft ONLY the reply body text. Do NOT include any headers (From, To, Date, Subject), greetings (Dear/Caro), sign-offs (Atenciosamente/Best regards), or signatures — those are added automatically by the email client. Just write the core reply content.";

function applyOfficeTheme() {
  // 1. Check URL query parameter passed from parent context
  const params = new URLSearchParams(window.location.search);
  const themeParam = params.get("theme");
  if (themeParam === "dark" || themeParam === "light") {
    document.documentElement.classList.toggle(
      "dark-theme",
      themeParam === "dark",
    );
    return;
  }

  // 2. Try Office.context.officeTheme (may not be available in dialogs)
  try {
    const theme = Office.context.officeTheme;
    if (theme && theme.bodyBackgroundColor) {
      const hex = theme.bodyBackgroundColor.replace("#", "");
      const r = parseInt(hex.substring(0, 2), 16);
      const g = parseInt(hex.substring(2, 4), 16);
      const b = parseInt(hex.substring(4, 6), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      document.documentElement.classList.toggle("dark-theme", luminance < 0.5);
      return;
    }
  } catch {
    // officeTheme not available in dialog context
  }

  // 3. Fall back to OS-level preference (prefers-color-scheme)
  if (
    window.matchMedia &&
    window.matchMedia("(prefers-color-scheme: dark)").matches
  ) {
    document.documentElement.classList.add("dark-theme");
  }
}

Office.onReady(() => {
  applyOfficeTheme();
  initI18n(ENV_APP_NAME);

  // Apply translations to data-i18n elements
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    el.textContent = t(el.dataset.i18n);
  });

  document.getElementById("prompt-input").value = DEFAULT_PROMPT;

  document.getElementById("submit-btn").addEventListener("click", () => {
    const prompt = document.getElementById("prompt-input").value.trim();
    if (!prompt) return;
    Office.context.ui.messageParent(JSON.stringify({action: "submit", prompt}));
  });

  document.getElementById("cancel-btn").addEventListener("click", () => {
    Office.context.ui.messageParent(JSON.stringify({action: "cancel"}));
  });
});

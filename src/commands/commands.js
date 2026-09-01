/* global Office */
import {initI18n, t, getLocale} from "../i18n.js";
import {getAccessToken} from "../auth.js";
import DOMPurify from "dompurify";
import {marked} from "marked";

const ENV = (typeof window !== "undefined" && window.__ENV) || {};
const ENV_API_URL = ENV.LIBRECHAT_API_URL || "";
const ENV_AGENT_ID = ENV.LIBRECHAT_AGENT_ID || "";
const ENV_APP_NAME = ENV.APP_NAME || "AI Assistant";

// The LibreChat agent owns its persona and instructions server-side
// (identity, security analysis, summary + suggested-reply contract).
// We intentionally do NOT prepend a local system prompt.

const REPLY_MARKER_START = "<!--ai-reply-start-->";
const REPLY_MARKER_END = "<!--ai-reply-end-->";
// Invisible Unicode markers that survive Outlook HTML-to-text conversion.
// Start: ZWSP + ZWNJ + ZWSP. End: ZWSP + ZWJ + ZWSP. Distinct sequences so
// we can reliably delimit the AI block even in plain-text body reads.
const REPLY_TEXT_MARKER_START = "\u200B\u200C\u200B";
const REPLY_TEXT_MARKER_END = "\u200B\u200D\u200B";

function wrapReplyHtml(htmlContent) {
  return `${REPLY_MARKER_START}<div class="ai-reply" data-ai-reply="1" style="font-family:inherit;font-size:inherit;color:inherit;">${REPLY_TEXT_MARKER_START}${htmlContent}${REPLY_TEXT_MARKER_END}</div>${REPLY_MARKER_END}`;
}

// Renders the agent's markdown reply into sanitized HTML for Outlook.
function sanitizeResponseHtml(response) {
  return DOMPurify.sanitize(marked.parse(response || ""));
}

/**
 * Convert an HTML fragment to plain text, preserving paragraph/line breaks.
 * Used to feed the compose body (after stripping prior AI replies) to the LLM.
 */
function htmlToPlainText(html) {
  if (!html) return "";
  // Use a detached element so scripts/styles are not executed.
  const container = document.createElement("div");
  container.innerHTML = DOMPurify.sanitize(html, {
    USE_PROFILES: {html: true},
    FORBID_TAGS: ["style", "script"],
  });
  // Replace common block/line separators with newlines before extracting text.
  container
    .querySelectorAll("br")
    .forEach((el) => el.replaceWith(document.createTextNode("\n")));
  container
    .querySelectorAll("p, div, li, tr, h1, h2, h3, h4, h5, h6, blockquote")
    .forEach((el) => el.append(document.createTextNode("\n")));
  return (container.textContent || "").replace(/\n{3,}/g, "\n\n");
}

// --- Agent response helpers ---

function isPhishingResponse(md) {
  return (
    /ATENÇÃO\s*—?\s*POSSÍVEL AMEAÇA/i.test(md) || /⛔\s*NÃO RESPONDAS/i.test(md)
  );
}

function isNoReplyNeeded(md) {
  return /não requer resposta|does not require a reply/i.test(md);
}

function extractSuggestedReply(md) {
  if (!md) return "";
  const re = /^##\s*✍️?\s*(?:Resposta Sugerida|Suggested Reply)\s*$/im;
  const parts = md.split(re);
  if (parts.length > 1) {
    return parts[1].replace(/^\s*---\s*$/m, "").trim();
  }
  return md.trim();
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Removes any previously-inserted AI reply block from the compose body HTML.
 * Tries the HTML-comment markers first (most reliable) and falls back to the
 * wrapping div.ai-reply (in case Outlook sanitization stripped comments).
 * Returns { cleaned, hadPrevious }.
 */
function stripPreviousReply(html) {
  if (!html) return {cleaned: html || "", hadPrevious: false};

  let hadPrevious = false;
  let cleaned = html;

  // 1) HTML-comment markers (most reliable when Outlook preserves comments).
  const commentRe = new RegExp(
    escapeRegex(REPLY_MARKER_START) +
      "[\\s\\S]*?" +
      escapeRegex(REPLY_MARKER_END),
    "g",
  );
  cleaned = cleaned.replace(commentRe, () => {
    hadPrevious = true;
    return "";
  });

  // 2) data-ai-reply attribute wrapper (Outlook preserves data-* attrs).
  const dataAttrRe = /<div[^>]*data-ai-reply=["']1["'][^>]*>[\s\S]*?<\/div>/gi;
  cleaned = cleaned.replace(dataAttrRe, () => {
    hadPrevious = true;
    return "";
  });

  // 3) Class wrapper fallback.
  const divRe =
    /<div[^>]*class=["'][^"']*\bai-reply\b[^"']*["'][^>]*>[\s\S]*?<\/div>/gi;
  cleaned = cleaned.replace(divRe, () => {
    hadPrevious = true;
    return "";
  });

  // 4) Invisible Unicode markers fallback (survives text coercion too).
  const textMarkerRe = new RegExp(
    escapeRegex(REPLY_TEXT_MARKER_START) +
      "[\\s\\S]*?" +
      escapeRegex(REPLY_TEXT_MARKER_END),
    "g",
  );
  cleaned = cleaned.replace(textMarkerRe, () => {
    hadPrevious = true;
    return "";
  });

  return {cleaned, hadPrevious};
}

/**
 * Strip any previously-inserted AI reply from a plain-text body. Relies on the
 * invisible Unicode start/end markers that survive HTML→Text conversion.
 */
function stripPreviousReplyText(text) {
  if (!text) return {cleaned: text || "", hadPrevious: false};
  const re = new RegExp(
    escapeRegex(REPLY_TEXT_MARKER_START) +
      "[\\s\\S]*?" +
      escapeRegex(REPLY_TEXT_MARKER_END),
    "g",
  );
  let hadPrevious = false;
  const cleaned = text.replace(re, () => {
    hadPrevious = true;
    return "";
  });
  return {cleaned, hadPrevious};
}

/**
 * Insert an AI reply into the compose body at the current cursor / selection.
 *
 * We use setSelectedDataAsync because it is the only Office.js body-write
 * method that preserves the user's scroll position on Outlook Windows
 * (prependAsync also preserves it, but does not allow replacement; setAsync
 * always jumps to the bottom).
 *
 * Trade-off: the reply lands wherever the caret currently is. If the user
 * has a previous AI reply they want to replace, they are expected to select
 * it (or position the cursor where they want the new reply) before clicking
 * the button. This is an intentional UX choice to avoid the scroll-jump.
 */
function insertReplyIntoBody(item, wrappedReply, onError, onDone) {
  console.log(" insertReplyIntoBody called (setSelectedDataAsync)");

  if (typeof item.body.setSelectedDataAsync !== "function") {
    // Very old host — fall back to prependAsync (no scroll jump but appends).
    console.warn(
      " setSelectedDataAsync unavailable, falling back to prependAsync",
    );
    item.body.prependAsync(
      wrappedReply,
      {coercionType: Office.CoercionType.Html},
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(" prependAsync failed:", result.error);
          onError(result.error?.message || "Failed to insert reply");
        }
        onDone();
      },
    );
    return;
  }

  item.body.setSelectedDataAsync(
    wrappedReply,
    {coercionType: Office.CoercionType.Html},
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(" setSelectedDataAsync failed:", result.error);
        onError(result.error?.message || "Failed to insert reply");
      } else {
        console.log(" setSelectedDataAsync succeeded");
      }
      onDone();
    },
  );
}

// Ribbon commands run with no UI surface to host a login prompt, so they are
// silent-only: they ride on the MSAL localStorage cache the task pane
// populated (taskpane.html and commands.html share an origin). When the cache
// is cold we surface a notification and stop — never acquireTokenPopup here.

// Acquires a token silently and returns the Authorization header value.
// Throws (InteractionRequiredAuthError et al.) when the cache is cold; the
// caller turns that into a notification and aborts the command.
async function getAuthHeader() {
  const token = await getAccessToken({allowInteractive: false});
  return `Bearer ${token}`;
}

// Silent-token guard for every command entry point. Returns the token when the
// cache is warm; when it is cold, shows the "open the task pane to sign in"
// notification, completes the event, and returns null so the caller stops.
async function requireSilentToken(event) {
  try {
    return await getAccessToken({allowInteractive: false});
  } catch {
    showNotificationError("signInRequired", t("notify.signInRequired"));
    event.completed();
    return null;
  }
}

function detectTheme() {
  try {
    const theme = Office.context.officeTheme;
    if (theme && theme.bodyBackgroundColor) {
      const hex = theme.bodyBackgroundColor.replace("#", "");
      const r = parseInt(hex.substring(0, 2), 16);
      const g = parseInt(hex.substring(2, 4), 16);
      const b = parseInt(hex.substring(4, 6), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      return luminance < 0.5 ? "dark" : "light";
    }
  } catch {
    // officeTheme not available
  }
  return "";
}

function buildDialogUrl() {
  const params = new URLSearchParams();
  const theme = detectTheme();
  if (theme) params.set("theme", theme);
  params.set("lang", getLocale());
  return `${location.protocol}//${location.host}/prompt-dialog.html?${params.toString()}`;
}

function getSettings() {
  const storedAgentId = Office.context.roamingSettings.get("agentId") || "";
  return {
    apiUrl: ENV_API_URL,
    model: storedAgentId || ENV_AGENT_ID,
  };
}

// --- Safe notification helpers ---
// OWA may not fully implement notificationMessages (throws internally).
// Guard every call by checking the method exists before invoking it.

function _getNotifications() {
  try {
    const nm = Office.context.mailbox.item?.notificationMessages;
    return nm && typeof nm.addAsync === "function" ? nm : null;
  } catch {
    return null;
  }
}

function safeNotifyAdd(key, type, message) {
  const nm = _getNotifications();
  const truncated =
    message.length > 150 ? message.slice(0, 147) + "..." : message;
  if (nm) nm.addAsync(key, {type, message: truncated});
}

function safeNotifyReplace(key, type, message) {
  const nm = _getNotifications();
  const truncated =
    message.length > 150 ? message.slice(0, 147) + "..." : message;
  if (nm && typeof nm.replaceAsync === "function") {
    nm.replaceAsync(key, {type, message: truncated});
  }
}

function safeNotifyRemove(key) {
  const nm = _getNotifications();
  if (nm && typeof nm.removeAsync === "function") {
    nm.removeAsync(key);
  }
}

function showNotificationError(key, message) {
  safeNotifyAdd(
    key,
    Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message,
  );
}

function showProgressNotification(key, message) {
  safeNotifyReplace(
    key,
    Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
    message,
  );
}

function removeNotification(key) {
  safeNotifyRemove(key);
}

/**
 * Opens a reply-all form and only calls event.completed() after the form is open.
 * Uses displayReplyAllFormAsync (Mailbox 1.9+) when available so the JS context
 * stays alive until the native reply window is ready. Falls back to the
 * synchronous displayReplyAllForm for older hosts.
 *
 * On Windows "new Outlook" the callback can fire before the reply window is
 * fully rendered, so we add a short delay before completing the event to
 * prevent the runtime from tearing down the context too early.
 */
function openReplyForm(item, htmlBody, event) {
  const completeWithDelay = () => {
    setTimeout(() => event.completed(), 1500);
  };

  try {
    if (typeof item.displayReplyAllFormAsync === "function") {
      item.displayReplyAllFormAsync({htmlBody}, (asyncResult) => {
        if (
          asyncResult &&
          asyncResult.status === Office.AsyncResultStatus.Failed
        ) {
          showNotificationError(
            "replyFormError",
            "Failed to open reply form: " +
              (asyncResult.error?.message || "unknown error"),
          );
        }
        completeWithDelay();
      });
    } else {
      item.displayReplyAllForm({htmlBody});
      completeWithDelay();
    }
  } catch (err) {
    showNotificationError(
      "replyFormError",
      "Failed to open reply form: " + (err.message || "unknown error"),
    );
    completeWithDelay();
  }
}

function getEmailData() {
  return new Promise((resolve, reject) => {
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error(t("error.noEmail")));
        return;
      }

      const subject = item.subject || t("error.noSubject");
      const from = item.from
        ? item.from.emailAddress || item.from.displayName || "Unknown"
        : item.sender
          ? item.sender.emailAddress || item.sender.displayName || "Unknown"
          : "Unknown";
      const to = item.to
        ? item.to.map((r) => r.emailAddress || r.displayName).join(", ")
        : "";
      const cc = item.cc
        ? item.cc.map((r) => r.emailAddress || r.displayName).join(", ")
        : "";
      const dateReceived = item.dateTimeCreated
        ? new Date(item.dateTimeCreated).toLocaleString()
        : "";

      if (item.body && typeof item.body.getAsync === "function") {
        item.body.getAsync(Office.CoercionType.Text, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve({
              subject,
              from,
              to,
              cc,
              date: dateReceived,
              body: result.value,
            });
          } else {
            resolve({
              subject,
              from,
              to,
              cc,
              date: dateReceived,
              body: item.preview || t("error.noBody"),
            });
          }
        });
      } else {
        resolve({
          subject,
          from,
          to,
          cc,
          date: dateReceived,
          body: item.preview || t("error.noBody"),
        });
      }
    } catch (err) {
      reject(err);
    }
  });
}

function formatEmailForPrompt(email) {
  return [
    `Subject: ${email.subject}`,
    `Sender: ${email.from}`,
    `To: ${email.to}`,
    `CC: ${email.cc || ""}`,
    `Datetime: ${email.date}`,
    ``,
    `Body:`,
    email.body,
  ].join("\n");
}

async function replyWithAI(event) {
  try {
    const settings = getSettings();

    const token = await requireSilentToken(event);
    if (!token) return;

    showProgressNotification("progress", t("notify.generating"));

    const email = await getEmailData();
    const emailText = formatEmailForPrompt(email);
    const response = await callLibreChatWithPrompt(settings, emailText, "");

    removeNotification("progress");

    if (isPhishingResponse(response)) {
      showNotificationError("phishing", t("notify.phishingBlocked"));
      event.completed();
      return;
    }
    if (isNoReplyNeeded(response)) {
      showNotificationError("noReply", t("notify.noReplyNeeded"));
      event.completed();
      return;
    }

    // Open reply form with only the Suggested Reply section, rendered.
    const htmlBody = sanitizeResponseHtml(extractSuggestedReply(response));
    openReplyForm(Office.context.mailbox.item, wrapReplyHtml(htmlBody), event);
    return;
  } catch (err) {
    removeNotification("progress");
    showNotificationError(
      "error",
      t("notify.replyFailed") + (err.message || t("error.unexpected")),
    );
    event.completed();
  }
}

// --- Compose-mode email extraction ---

function getComposeEmailData() {
  return new Promise((resolve, reject) => {
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error(t("error.noCompose")));
        return;
      }

      const getData = {};
      const promises = [];

      promises.push(
        new Promise((res) => {
          if (item.subject && typeof item.subject.getAsync === "function") {
            item.subject.getAsync((result) => {
              getData.subject =
                result.status === Office.AsyncResultStatus.Succeeded
                  ? result.value
                  : t("error.noSubject");
              res();
            });
          } else {
            getData.subject = item.subject || t("error.noSubject");
            res();
          }
        }),
      );

      promises.push(
        new Promise((res) => {
          if (item.to && typeof item.to.getAsync === "function") {
            item.to.getAsync((result) => {
              getData.to =
                result.status === Office.AsyncResultStatus.Succeeded
                  ? result.value
                      .map((r) => r.emailAddress || r.displayName)
                      .join(", ")
                  : "";
              res();
            });
          } else {
            getData.to = "";
            res();
          }
        }),
      );

      promises.push(
        new Promise((res) => {
          if (item.cc && typeof item.cc.getAsync === "function") {
            item.cc.getAsync((result) => {
              getData.cc =
                result.status === Office.AsyncResultStatus.Succeeded
                  ? result.value
                      .map((r) => r.emailAddress || r.displayName)
                      .join(", ")
                  : "";
              res();
            });
          } else {
            getData.cc = "";
            res();
          }
        }),
      );

      promises.push(
        new Promise((res) => {
          if (item.body && typeof item.body.getAsync === "function") {
            // Read as HTML so we can reliably strip any previously-inserted
            // AI reply block before sending the body to the LLM. Without this
            // the agent re-analyses its own output and may trigger the
            // security-check / phishing branch on a plain reply request.
            item.body.getAsync(Office.CoercionType.Html, (result) => {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                getData.body = t("error.noComposeBody");
                res();
                return;
              }
              const {cleaned, hadPrevious} = stripPreviousReply(
                result.value || "",
              );
              if (hadPrevious) {
                console.log(
                  " previous AI reply stripped from compose body before LLM call",
                );
              }
              // Convert the cleaned HTML to plain text for the prompt.
              const textBody = htmlToPlainText(cleaned);
              const {cleaned: finalText} = stripPreviousReplyText(textBody);
              getData.body = finalText.trim() || t("error.noComposeBody");
              res();
            });
          } else {
            getData.body = t("error.noComposeBody");
            res();
          }
        }),
      );

      Promise.all(promises).then(() => {
        resolve({
          subject: getData.subject,
          from: Office.context.mailbox.userProfile.emailAddress || "Me",
          to: getData.to,
          cc: getData.cc || "",
          date: new Date().toLocaleString(),
          body: getData.body,
        });
      });
    } catch (err) {
      reject(err);
    }
  });
}

// Directive prefixes signalled to the agent so it knows which action the
// user triggered. Keep in sync with the agent's instructions.
const ACTION_DIRECTIVES = {
  summarize:
    "Action: SUMMARIZE. Run the full security check and return the complete Summary + Suggested Reply contract as defined in your instructions.",
  reply:
    "Action: REPLY. Run the security check. If the email is safe and a reply is needed, return ONLY the '## ✍️ Resposta Sugerida' section as defined in your instructions (no summary). If it is phishing or does not require a reply, return the corresponding block.",
};

async function callLibreChatWithPrompt(
  settings,
  emailText,
  customPrompt,
  action = "reply",
) {
  const base = settings.apiUrl.replace(/\/+$/, "");
  const url = base + "/api/agents/v1/chat/completions";

  const directive = ACTION_DIRECTIVES[action] || ACTION_DIRECTIVES.reply;
  // If the user provided extra instructions, treat them as additional context
  // on top of the action directive (never as a replacement for the agent's
  // persona/output contract, which lives on the server).
  const extra = customPrompt
    ? `\n\nAdditional user instructions (apply within your standard output contract):\n${customPrompt}`
    : "";
  const content = `${directive}${extra}\n\nHere is the email:\n\n${emailText}`;
  const messages = [{role: "user", content}];

  const headers = {
    "Content-Type": "application/json",
    Authorization: await getAuthHeader(),
  };

  const payload = {model: settings.model, messages, stream: false};
  const response = await fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify(payload),
  });

  const responseText = await response.text();

  if (!response.ok) {
    let detail = responseText;
    try {
      const errBody = JSON.parse(responseText);
      detail = errBody.error?.message || responseText;
    } catch {
      /* use responseText */
    }
    throw new Error(`API error ${response.status}: ${detail}`);
  }

  const data = JSON.parse(responseText);
  if (data.choices && data.choices.length > 0) {
    return data.choices[0].message.content || "";
  }
  throw new Error("No response content from API.");
}

async function composeReplyCustom(event) {
  let settings;
  try {
    settings = getSettings();
  } catch {
    showNotificationError("settingsError", t("notify.settingsError"));
    event.completed();
    return;
  }

  const token = await requireSilentToken(event);
  if (!token) return;

  const dialogUrl = buildDialogUrl();

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {height: 50, width: 40, displayInIframe: true},
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showNotificationError(
          "dialogError",
          t("notify.dialogError") + (result.error?.message || ""),
        );
        event.completed();
        return;
      }

      const dialog = result.value;
      let messageReceived = false;

      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        async (arg) => {
          messageReceived = true;
          dialog.close();

          let msg;
          try {
            msg = JSON.parse(arg.message);
          } catch {
            event.completed();
            return;
          }

          if (msg.action !== "submit" || !msg.prompt) {
            event.completed();
            return;
          }

          try {
            showProgressNotification("progress", t("notify.generating"));

            const email = await getComposeEmailData();
            const emailText = formatEmailForPrompt(email);
            const response = await callLibreChatWithPrompt(
              settings,
              emailText,
              msg.prompt,
            );

            removeNotification("progress");

            if (isPhishingResponse(response)) {
              showNotificationError("phishing", t("notify.phishingBlocked"));
              event.completed();
              return;
            }
            if (isNoReplyNeeded(response)) {
              showNotificationError("noReply", t("notify.noReplyNeeded"));
              event.completed();
              return;
            }

            // Insert reply into compose body
            const item = Office.context.mailbox.item;
            const htmlReply = sanitizeResponseHtml(
              extractSuggestedReply(response),
            );
            const wrappedReply = wrapReplyHtml(htmlReply);

            insertReplyIntoBody(
              item,
              wrappedReply,
              (errMsg) =>
                showNotificationError("error", t("error.insertReply") + errMsg),
              () => event.completed(),
            );
          } catch (err) {
            removeNotification("progress");
            showNotificationError(
              "error",
              t("notify.replyFailed") + (err.message || t("error.unexpected")),
            );
            event.completed();
          }
        },
      );

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (messageReceived) return;
        switch (arg.error) {
          case 12006:
            // Dialog closed by user — normal, just complete.
            break;
          case 12002:
            showNotificationError(
              "dialogError",
              t("notify.dialogError") + "Page not found.",
            );
            break;
          case 12003:
            showNotificationError(
              "dialogError",
              t("notify.dialogError") + "HTTPS required.",
            );
            break;
          default:
            showNotificationError(
              "dialogError",
              t("notify.dialogError") + (arg.message || ""),
            );
            break;
        }
        event.completed();
      });
    },
  );
}

async function composeReplyDefault(event) {
  console.log(" composeReplyDefault called");
  try {
    const settings = getSettings();

    const token = await requireSilentToken(event);
    if (!token) return;

    showProgressNotification("progress", t("notify.generating"));

    const email = await getComposeEmailData();
    console.log(" email data retrieved, subject:", email.subject);
    const emailText = formatEmailForPrompt(email);
    const response = await callLibreChatWithPrompt(settings, emailText, "");
    console.log(" API response received, length:", response.length);

    removeNotification("progress");

    if (isPhishingResponse(response)) {
      showNotificationError("phishing", t("notify.phishingBlocked"));
      event.completed();
      return;
    }
    if (isNoReplyNeeded(response)) {
      showNotificationError("noReply", t("notify.noReplyNeeded"));
      event.completed();
      return;
    }

    // Insert only the Suggested Reply section, rendered as HTML.
    const item = Office.context.mailbox.item;
    const htmlReply = sanitizeResponseHtml(extractSuggestedReply(response));
    const wrappedReply = wrapReplyHtml(htmlReply);

    insertReplyIntoBody(
      item,
      wrappedReply,
      (errMsg) => {
        console.error(" insertReplyIntoBody error:", errMsg);
        showNotificationError("error", t("error.insertReply") + errMsg);
      },
      () => {
        console.log(" composeReplyDefault completed");
        event.completed();
      },
    );
  } catch (err) {
    console.error(" composeReplyDefault exception:", err);
    removeNotification("progress");
    showNotificationError(
      "error",
      t("notify.replyFailed") + (err.message || t("error.unexpected")),
    );
    event.completed();
  }
}

async function readReplyCustom(event) {
  let settings;
  try {
    settings = getSettings();
  } catch {
    showNotificationError("settingsError", t("notify.settingsError"));
    event.completed();
    return;
  }

  const token = await requireSilentToken(event);
  if (!token) return;

  const dialogUrl = buildDialogUrl();

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {height: 50, width: 40, displayInIframe: true},
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showNotificationError(
          "dialogError",
          t("notify.dialogError") + (result.error?.message || ""),
        );
        event.completed();
        return;
      }

      const dialog = result.value;
      let messageReceived = false;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        messageReceived = true;
        dialog.close();

        let msg;
        try {
          msg = JSON.parse(arg.message);
        } catch {
          event.completed();
          return;
        }

        if (msg.action !== "submit" || !msg.prompt) {
          event.completed();
          return;
        }

        // Delay processing to allow the dialog window to fully close on
        // Windows before we open the reply form (only one popup at a time).
        setTimeout(async () => {
          try {
            showProgressNotification("progress", t("notify.generating"));

            const email = await getEmailData();
            const emailText = formatEmailForPrompt(email);
            const response = await callLibreChatWithPrompt(
              settings,
              emailText,
              msg.prompt,
            );

            removeNotification("progress");

            if (isPhishingResponse(response)) {
              showNotificationError("phishing", t("notify.phishingBlocked"));
              event.completed();
              return;
            }
            if (isNoReplyNeeded(response)) {
              showNotificationError("noReply", t("notify.noReplyNeeded"));
              event.completed();
              return;
            }

            const htmlBody = sanitizeResponseHtml(
              extractSuggestedReply(response),
            );
            openReplyForm(
              Office.context.mailbox.item,
              wrapReplyHtml(htmlBody),
              event,
            );
          } catch (err) {
            removeNotification("progress");
            showNotificationError(
              "error",
              t("notify.replyFailed") + (err.message || t("error.unexpected")),
            );
            event.completed();
          }
        }, 500);
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        if (messageReceived) return;
        switch (arg.error) {
          case 12006:
            break;
          case 12002:
            showNotificationError(
              "dialogError",
              t("notify.dialogError") + "Page not found.",
            );
            break;
          case 12003:
            showNotificationError(
              "dialogError",
              t("notify.dialogError") + "HTTPS required.",
            );
            break;
          default:
            showNotificationError(
              "dialogError",
              t("notify.dialogError") + (arg.message || ""),
            );
            break;
        }
        event.completed();
      });
    },
  );
}

// Global safety net — catch any truly unexpected errors and show a notification
window.addEventListener("unhandledrejection", (e) => {
  const msg = e.reason?.message || String(e.reason || "Unknown error");
  showNotificationError("unexpectedError", t("notify.unexpectedError") + msg);
});

window.addEventListener("error", (e) => {
  const msg = e.message || "Unknown error";
  showNotificationError("unexpectedError", t("notify.unexpectedError") + msg);
});

Office.onReady(() => {
  initI18n(ENV_APP_NAME);
  Office.actions.associate("replyWithAI", replyWithAI);
  Office.actions.associate("readReplyCustom", readReplyCustom);
  Office.actions.associate("composeReplyDefault", composeReplyDefault);
  Office.actions.associate("composeReplyCustom", composeReplyCustom);
});

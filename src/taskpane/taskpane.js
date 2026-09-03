/* global Office, __ADDIN_VERSION__ */
import {marked} from "marked";
import DOMPurify from "dompurify";
import {initI18n, t} from "../i18n.js";
import {
  isNaaSupported,
  getAccessToken,
  getSignedInUser,
  signIn,
  signOut,
} from "../auth.js";
import {getManifestVersion} from "./manifestVersion.js";
import {
  recordAdoption,
  recordActionStarted,
  recordActionFinished,
  recordError,
} from "../telemetry.js";

const ENV = (typeof window !== "undefined" && window.__ENV) || {};
const ENV_API_URL = ENV.LIBRECHAT_API_URL || "";
const ENV_AGENT_ID = ENV.LIBRECHAT_AGENT_ID || "";
const ENV_APP_NAME = ENV.APP_NAME || "AI Assistant";
const ENV_APP_LOGO_URL = ENV.APP_LOGO_URL || "";
const ADDIN_VERSION =
  typeof __ADDIN_VERSION__ === "string" ? __ADDIN_VERSION__ : "dev";
const MANIFEST_VERSION = getManifestVersion();

// The LibreChat agent owns its persona and instructions server-side
// (identity, security analysis, summary + suggested-reply contract).
// We intentionally do NOT prepend a local system prompt — doing so
// would conflict with the agent's own instructions.

// --- Office Theme Detection ---

function applyOfficeTheme() {
  try {
    const theme = Office.context.officeTheme;
    if (theme && theme.bodyBackgroundColor) {
      const hex = theme.bodyBackgroundColor.replace("#", "");
      const r = parseInt(hex.substring(0, 2), 16);
      const g = parseInt(hex.substring(2, 4), 16);
      const b = parseInt(hex.substring(4, 6), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      document.documentElement.classList.toggle("dark-theme", luminance < 0.5);
    }
  } catch {
    // officeTheme not available — keep default light theme
  }
}

// --- Settings persistence (Office.context.roamingSettings) ---

function loadSettings() {
  const rs = Office.context.roamingSettings;
  const agentId = rs.get("agentId") || "";

  return {apiUrl: ENV_API_URL, model: agentId || ENV_AGENT_ID};
}

function getSettings() {
  const agentSelect = document.getElementById("agent-select");
  const selectedAgent = agentSelect ? agentSelect.value : "";
  const rs = Office.context.roamingSettings;
  const storedAgent = rs.get("agentId") || "";

  return {
    apiUrl: ENV_API_URL,
    model: selectedAgent || storedAgent || ENV_AGENT_ID,
  };
}

// Remove the legacy per-user API key from roaming settings. The field is gone
// from the UI; leaving the value behind would be dead credential material.
// Non-blocking by design: a purge failure must never stop the add-in.
async function purgeLegacyApiKey() {
  try {
    const rs = Office.context.roamingSettings;
    if (!rs.get("apiKey")) return;
    rs.remove("apiKey");
    await new Promise((resolve) => {
      rs.saveAsync(() => resolve());
    });
  } catch (err) {
    console.warn("Could not purge legacy apiKey from roamingSettings:", err);
  }
}

// --- Agent listing ---

async function listAgents() {
  const base = ENV_API_URL.replace(/\/+$/, "");
  const url = base + "/api/agents/v1/models";

  const token = await getAccessToken({allowInteractive: true});
  const headers = {Authorization: `Bearer ${token}`};

  const response = await fetch(url, {method: "GET", headers});

  if (!response.ok) {
    const err = new Error(`${response.status} ${response.statusText}`);
    err.status = response.status;
    throw err;
  }

  const data = await response.json();
  const models = data.data || [];
  return models.map((m) => ({id: m.id, name: m.name || m.id}));
}

function populateAgentDropdown(agents, selectedId) {
  const select = document.getElementById("agent-select");
  // Clear existing options except the placeholder
  while (select.options.length > 1) {
    select.remove(1);
  }

  agents.forEach((agent) => {
    const option = document.createElement("option");
    option.value = agent.id;
    option.textContent = agent.name;
    select.appendChild(option);
  });

  select.disabled = agents.length === 0;

  if (selectedId) {
    const exists = agents.some((a) => a.id === selectedId);
    if (exists) {
      select.value = selectedId;
    } else if (agents.length > 0) {
      select.value = agents[0].id;
    }
  } else if (agents.length > 0) {
    select.value = agents[0].id;
  }
}

function showAgentStatus(type, message) {
  const el = document.getElementById("agent-status");
  el.className = `connection-status connection-status-${type}`;
  el.textContent = message;
  el.classList.remove("hidden");
}

function hideAgentStatus() {
  document.getElementById("agent-status").classList.add("hidden");
}

// Loads the agent list once, after sign-in, and populates the dropdown.
async function loadAgents() {
  hideAgentStatus();
  const select = document.getElementById("agent-select");
  try {
    const agents = await listAgents();
    if (agents.length === 0) {
      showAgentStatus("warning", t("settings.noAgentsFound"));
      select.disabled = true;
      return;
    }
    const rs = Office.context.roamingSettings;
    const storedAgentId = rs.get("agentId") || ENV_AGENT_ID;
    populateAgentDropdown(agents, storedAgentId);
  } catch (err) {
    select.disabled = true;
    if (err.status === 401) {
      // Token was valid, but LibreChat could not match it to a user.
      showAgentStatus("error", t("auth.identityMismatch"));
    } else {
      showAgentStatus("error", err.message || "" || t("error.unexpected"));
    }
  }
}

// --- Email extraction using Office.js ---

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

      // Try body.getAsync first (requires Mailbox 1.3+)
      // Fall back to item.body plain text preview if not available
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
          body: item.preview || t("error.noBodyUpgrade"),
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

// --- Agent response helpers ---

// Detects the phishing-alert block produced by the agent's security step.
function isPhishingResponse(md) {
  return (
    /ATENÇÃO\s*—?\s*POSSÍVEL AMEAÇA/i.test(md) || /⛔\s*NÃO RESPONDAS/i.test(md)
  );
}

// Detects the "no reply needed" marker.
function isNoReplyNeeded(md) {
  return /não requer resposta|does not require a reply/i.test(md);
}

// Extracts the "Resposta Sugerida" / "Suggested Reply" section from the
// agent's markdown output so we only insert the reply body — not the
// security analysis or the summary — into the Outlook compose window.
function extractSuggestedReply(md) {
  if (!md) return "";
  const re = /^##\s*✍️?\s*(?:Resposta Sugerida|Suggested Reply)\s*$/im;
  const parts = md.split(re);
  if (parts.length > 1) {
    // Strip an optional horizontal rule left over from the template.
    return parts[1].replace(/^\s*---\s*$/m, "").trim();
  }
  return md.trim();
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

      // In compose mode, subject and recipients are async
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
            item.body.getAsync(Office.CoercionType.Text, (result) => {
              getData.body =
                result.status === Office.AsyncResultStatus.Succeeded
                  ? result.value
                  : t("error.noComposeBody");
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

// --- LibreChat API (OpenAI-compatible) ---

// Directive prefixes signalled to the agent so it knows whether to produce
// the full contract (security + summary + suggested reply) or only the reply.
const ACTION_DIRECTIVES = {
  summarize:
    "Action: SUMMARIZE. Run the full security check and return the complete Summary + Suggested Reply contract as defined in your instructions.",
  reply:
    "Action: REPLY. Run the security check. If the email is safe and a reply is needed, return ONLY the '## ✍️ Resposta Sugerida' section as defined in your instructions (no summary). If it is phishing or does not require a reply, return the corresponding block.",
};

async function callLibreChat(
  settings,
  emailText,
  action = "summarize",
  directiveOverride = null,
) {
  const base = settings.apiUrl.replace(/\/+$/, "");
  const url = base + "/api/agents/v1/chat/completions";

  const directive =
    directiveOverride ||
    ACTION_DIRECTIVES[action] ||
    ACTION_DIRECTIVES.summarize;
  // The agent's persona and output contract live on the LibreChat server.
  // We only tell it which action the user picked and provide the text.
  const label =
    action === "tone" ? "Here is the text to rewrite:" : "Here is the email:";
  const content = `${directive}\n\n${label}\n\n${emailText}`;

  const messages = [{role: "user", content}];

  const headers = {
    "Content-Type": "application/json",
    Authorization: `Bearer ${await getAccessToken({allowInteractive: true})}`,
  };

  const payload = {
    model: settings.model,
    messages,
    stream: false,
  };

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
      // already have responseText as detail
    }
    throw new Error(`API error ${response.status}: ${detail}`);
  }

  const data = JSON.parse(responseText);
  if (data.choices && data.choices.length > 0) {
    return data.choices[0].message.content;
  }
  throw new Error(t("error.noApiResponse"));
}

// --- UI Helpers ---

function showError(msg) {
  const errorBox = document.getElementById("error-box");
  const errorMsg = document.getElementById("error-message");
  errorMsg.textContent = msg;
  errorBox.classList.remove("hidden");
}

function hideError() {
  document.getElementById("error-box").classList.add("hidden");
}

function showResponse(text) {
  hideLoading();
  const box = document.getElementById("response-box");
  const content = document.getElementById("response-content");
  content.innerHTML = DOMPurify.sanitize(marked.parse(text));
  box.classList.remove("hidden");
}

function hideResponse() {
  document.getElementById("response-box").classList.add("hidden");
}

function showLoading() {
  hideResponse();
  hideError();
  document.getElementById("loading-box").classList.remove("hidden");
}

function hideLoading() {
  document.getElementById("loading-box").classList.add("hidden");
}

function showEmailPreview(email) {
  const preview = document.getElementById("email-preview");
  const bodySnippet =
    email.body.length > 200 ? email.body.substring(0, 200) + "…" : email.body;
  preview.innerHTML = `
    <div class="email-meta">
      <strong>${escapeHtml(email.subject)}</strong><br/>
      From: ${escapeHtml(email.from)} &middot; ${escapeHtml(email.date)}
    </div>
    <div>${escapeHtml(bodySnippet)}</div>
  `;
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.appendChild(document.createTextNode(text));
  return div.innerHTML;
}

// --- Compose body insertion ---

// Invisible Unicode marker that survives Outlook HTML sanitization.
const REPLY_TEXT_MARKER = "\u200B\u200C\u200B";

/**
 * Insert an AI reply into the compose body.
 *
 * Uses prependAsync exclusively — this is the ONLY Office.js body-write
 * method that does not reset the scroll position in Outlook on Windows.
 * setAsync always scrolls to the bottom, so we never call it.
 *
 * Trade-off: if the user generates a second reply, the previous one
 * remains in the body below the new one.  This is acceptable because
 * the alternative (setAsync) causes a disruptive scroll-jump.
 */
function insertReplyIntoBody(item, wrappedReply, onError) {
  item.body.prependAsync(
    wrappedReply,
    {coercionType: Office.CoercionType.Html},
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        onError(result.error?.message || "Failed to prepend reply");
      }
    },
  );
}

// --- View switching ---

function showView(view) {
  const mainView = document.getElementById("main-view");
  const settingsView = document.getElementById("settings-view");
  const unsupportedView = document.getElementById("auth-unsupported-view");
  const signInView = document.getElementById("auth-signin-view");
  const settingsBtn = document.getElementById("settings-btn");
  const backBtn = document.getElementById("back-btn");

  // Hide everything, then reveal the requested view.
  [mainView, settingsView, unsupportedView, signInView].forEach((el) =>
    el.classList.add("hidden"),
  );
  settingsBtn.classList.add("hidden");
  backBtn.classList.add("hidden");

  if (view === "settings") {
    settingsView.classList.remove("hidden");
    backBtn.classList.remove("hidden");
  } else if (view === "auth-unsupported") {
    unsupportedView.classList.remove("hidden");
  } else if (view === "auth-signin") {
    signInView.classList.remove("hidden");
  } else {
    mainView.classList.remove("hidden");
    settingsBtn.classList.remove("hidden");
  }
}

// --- Auth gate ---

function setAccountStatus(text) {
  const el = document.getElementById("account-status");
  if (el) el.textContent = text;
}

function setText(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

// Signs the user in interactively (the ONLY place consent can be granted) and
// returns the token. Throws typed errors that the caller maps to a screen.
async function signInInteractively() {
  setAccountStatus(t("auth.signingIn"));
  return getAccessToken({allowInteractive: true});
}

// The startup gate. Returns true when the add-in may proceed to the requested
// action; false when it rendered a blocking/sign-in screen and must stop.
async function ensureSignedIn() {
  if (!isNaaSupported()) {
    showView("auth-unsupported");
    return false;
  }
  try {
    await signInInteractively();
  } catch (err) {
    // Popup dismissed/blocked or another interactive failure: show a
    // retryable screen with a Sign in button.
    showAuthSignIn(err);
    return false;
  }
  await purgeLegacyApiKey();
  await refreshAccountCard();
  return true;
}

function showAuthSignIn(err) {
  const detail = document.getElementById("auth-error-detail");
  if (detail) {
    const msg = err && err.message ? err.message : "";
    detail.textContent = msg ? t("auth.failed") + msg : "";
    detail.classList.toggle("hidden", !msg);
  }
  showView("auth-signin");
}

async function refreshAccountCard() {
  try {
    const user = await getSignedInUser();
    setText("account-email", user || t("auth.signingIn"));
    setAccountStatus(user ? t("settings.connected") : t("auth.signingIn"));
    setText("deployment-version", ADDIN_VERSION);
    setText("manifest-version", MANIFEST_VERSION);
    setText("instance-host", getInstanceHost());
  } catch {
    setAccountStatus(t("auth.signingIn"));
  }
}

function getInstanceHost() {
  try {
    return (
      new URL(ENV_API_URL).host || ENV_API_URL || t("settings.notConfigured")
    );
  } catch {
    return ENV_API_URL || t("settings.notConfigured");
  }
}

async function handleSignOut() {
  const button = document.getElementById("sign-out-btn");
  const errorBox = document.getElementById("settings-error-box");
  const errorMessage = document.getElementById("settings-error-message");
  if (button) button.disabled = true;
  if (errorBox) errorBox.classList.add("hidden");
  try {
    await signOut();
    showView("auth-signin");
  } catch (err) {
    if (errorMessage) {
      errorMessage.textContent = err.message || t("error.signOutFailed");
    }
    if (errorBox) errorBox.classList.remove("hidden");
    if (button) button.disabled = false;
  }
}

// Retry entry point for the Sign in button on the auth-signin view.
async function retrySignIn() {
  const btn = document.getElementById("auth-signin-btn");
  if (btn) btn.disabled = true;
  try {
    await signIn();
    await purgeLegacyApiKey();
    await refreshAccountCard();
    // Signed in — reload so the pending action runs cleanly.
    window.location.reload();
  } catch (err) {
    showAuthSignIn(err);
  } finally {
    if (btn) btn.disabled = false;
  }
}

// --- Main flow ---

// With SSO the startup gate has already signed the user in, so by the time an
// action runs there is nothing to validate. Kept as a thin wrapper so the
// action handlers read the same way as before.
function validateSettings() {
  return getSettings();
}

async function summarizeEmail() {
  hideError();
  hideResponse();

  const settings = validateSettings();
  if (!settings) return;

  showLoading();
  recordActionStarted("summarize");
  const startedAt = Date.now();
  const phaseDurations = {};
  const readingStartedAt = Date.now();

  try {
    const email = await getEmailData();
    phaseDurations.reading_email = Date.now() - readingStartedAt;
    showEmailPreview(email);

    const emailText = formatEmailForPrompt(email);
    const response = await callLibreChat(settings, emailText, "summarize");
    phaseDurations.thinking =
      Date.now() - readingStartedAt - phaseDurations.reading_email;

    showResponse(response);
    recordActionFinished("summarize", "completed", {
      durationMs: Date.now() - startedAt,
      phaseDurations,
    });
  } catch (err) {
    hideLoading();
    recordError("summarize", err, Date.now() - startedAt, phaseDurations);
    showError(err.message || t("error.unexpected"));
  }
}

async function replyWithAI() {
  hideError();
  hideResponse();

  const settings = validateSettings();
  if (!settings) return;

  showLoading();
  recordActionStarted("reply");
  const startedAt = Date.now();
  const phaseDurations = {};
  const readingStartedAt = Date.now();

  try {
    const email = await getEmailData();
    phaseDurations.reading_email = Date.now() - readingStartedAt;
    showEmailPreview(email);

    const emailText = formatEmailForPrompt(email);
    const response = await callLibreChat(settings, emailText, "reply");
    phaseDurations.thinking =
      Date.now() - readingStartedAt - phaseDurations.reading_email;

    showResponse(response);

    if (isPhishingResponse(response)) {
      showError(t("error.phishingBlocked"));
      recordActionFinished("reply", "cancelled", {
        durationMs: Date.now() - startedAt,
        phaseDurations,
        errorKind: "invalid_response",
      });
      return;
    }
    if (isNoReplyNeeded(response)) {
      showError(t("error.noReplyNeeded"));
      recordActionFinished("reply", "cancelled", {
        durationMs: Date.now() - startedAt,
        phaseDurations,
        errorKind: "invalid_response",
      });
      return;
    }

    // Open reply form with the Suggested Reply section only, rendered as HTML.
    const replyMd = extractSuggestedReply(response);
    const item = Office.context.mailbox.item;
    const replyHtml = DOMPurify.sanitize(marked.parse(replyMd));
    // Wrap so the inserted reply inherits the user's default compose font.
    const htmlBody = `<div style="font-family:inherit;font-size:inherit;color:inherit;">${replyHtml}</div>`;
    item.displayReplyAllForm({
      htmlBody: htmlBody,
    });

    recordActionFinished("reply", "completed", {
      durationMs: Date.now() - startedAt,
      phaseDurations,
    });
  } catch (err) {
    hideLoading();
    recordError("reply", err, Date.now() - startedAt, phaseDurations);
    showError(err.message || t("error.unexpected"));
  }
}

async function composeReplyWithAI() {
  hideError();
  hideResponse();

  const settings = validateSettings();
  if (!settings) return;

  showLoading();

  try {
    const email = await getComposeEmailData();
    showEmailPreview(email);

    const emailText = formatEmailForPrompt(email);
    const response = await callLibreChat(settings, emailText, "reply");

    showResponse(response);

    if (isPhishingResponse(response)) {
      showError(t("error.phishingBlocked"));
      return;
    }
    if (isNoReplyNeeded(response)) {
      showError(t("error.noReplyNeeded"));
      return;
    }

    // Insert only the Suggested Reply section, rendered as HTML.
    const replyMd = extractSuggestedReply(response);
    const item = Office.context.mailbox.item;
    const htmlReply = DOMPurify.sanitize(marked.parse(replyMd));
    const wrappedReply = `<div class="ai-reply">${REPLY_TEXT_MARKER}${htmlReply}</div>`;

    insertReplyIntoBody(item, wrappedReply, (errMsg) =>
      showError(t("error.insertReply") + errMsg),
    );
  } catch (err) {
    hideLoading();
    showError(err.message || t("error.unexpected"));
  }
}

async function composeReplyWithCustomInstructions() {
  hideError();
  hideResponse();

  const settings = validateSettings();
  if (!settings) return;

  const customPrompt = document.getElementById("custom-prompt").value.trim();
  if (!customPrompt) {
    showError(t("error.enterPrompt"));
    return;
  }

  const btn = document.getElementById("custom-reply-btn");
  const btnText = document.getElementById("custom-reply-btn-text");
  const spinner = document.getElementById("custom-reply-btn-spinner");
  btn.disabled = true;
  btnText.textContent = t("taskpane.generatingReply");
  spinner.classList.remove("hidden");
  showLoading();

  try {
    const email = await getComposeEmailData();
    showEmailPreview(email);

    const emailText = formatEmailForPrompt(email);
    // Reply action + user's additional instructions on top of the agent contract.
    const base = settings.apiUrl.replace(/\/+$/, "");
    const url = base + "/api/agents/v1/chat/completions";

    const directive = ACTION_DIRECTIVES.reply;
    const content = `${directive}\n\nAdditional user instructions (apply within your standard output contract):\n${customPrompt}\n\nHere is the email:\n\n${emailText}`;
    const messages = [{role: "user", content}];

    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${await getAccessToken({allowInteractive: true})}`,
    };

    const payload = {model: settings.model, messages, stream: false};
    const resp = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });
    const responseText = await resp.text();

    if (!resp.ok) {
      let detail = responseText;
      try {
        const errBody = JSON.parse(responseText);
        detail = errBody.error?.message || responseText;
      } catch {
        /* use responseText */
      }
      throw new Error(`API error ${resp.status}: ${detail}`);
    }

    const data = JSON.parse(responseText);
    if (!data.choices || data.choices.length === 0) {
      throw new Error("No response content from API.");
    }
    const response = data.choices[0].message.content;

    showResponse(response);

    if (isPhishingResponse(response)) {
      showError(t("error.phishingBlocked"));
      return;
    }
    if (isNoReplyNeeded(response)) {
      showError(t("error.noReplyNeeded"));
      return;
    }

    const replyMd = extractSuggestedReply(response);
    const item = Office.context.mailbox.item;
    const htmlReply = DOMPurify.sanitize(marked.parse(replyMd));
    const wrappedReply = `<div class="ai-reply">${REPLY_TEXT_MARKER}${htmlReply}</div>`;

    insertReplyIntoBody(item, wrappedReply, (errMsg) =>
      showError(t("error.insertReply") + errMsg),
    );
  } catch (err) {
    hideLoading();
    showError(err.message || t("error.unexpected"));
  } finally {
    btn.disabled = false;
    btnText.textContent = t("taskpane.generateReply");
    spinner.classList.add("hidden");
  }
}

// --- Tone rewrite ---

// Maps each tone preset to the phrasing sent to the agent. The agent keeps the
// input language and only adjusts tone; descriptions are intentionally in
// English (the directive language) but never override the email's language.
const TONE_DESCRIPTIONS = {
  formal: "formal and professional",
  friendly: "warm, friendly and approachable",
  concise:
    "concise and to the point, removing redundancy while keeping all key information",
  assertive: "confident and assertive",
};

function buildToneDirective(tone) {
  const desc = TONE_DESCRIPTIONS[tone] || TONE_DESCRIPTIONS.formal;
  return (
    `Action: REWRITE_TONE. Rewrite the user's text so it sounds more ${desc}. ` +
    `Keep the SAME language as the input. Preserve all facts, names, numbers and intent. ` +
    `Do not add greetings, signatures, or any commentary that is not in the original. ` +
    `Do not run a security analysis and do not produce a summary. ` +
    `Return ONLY the rewritten text as plain prose (no markdown headings).`
  );
}

// Holds the latest tone-rewrite state so the Apply button can write it back.
const toneState = {tone: "formal", rewritten: ""};

// Reads the user's current selection from the compose body.
function getSelectedText() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.getSelectedDataAsync !== "function") {
      reject(new Error(t("tone.noSelectionApi")));
      return;
    }
    item.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve((result.value && result.value.data) || "");
      } else {
        reject(new Error(result.error?.message || t("error.unexpected")));
      }
    });
  });
}

// Replaces the current selection with the rewritten text. setSelectedDataAsync
// preserves the scroll position (unlike body.setAsync) and targets the same
// selection the user had when they triggered the rewrite.
function applyToneRewrite(html) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.setSelectedDataAsync !== "function") {
      reject(new Error(t("tone.noSelectionApi")));
      return;
    }
    item.setSelectedDataAsync(
      html,
      {coercionType: Office.CoercionType.Html},
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || t("error.unexpected")));
        }
      },
    );
  });
}

function plainTextToHtml(text) {
  return escapeHtml(text).replace(/\n/g, "<br/>");
}

async function runToneRewrite(tone) {
  hideError();
  hideResponse();
  document.getElementById("tone-actions").classList.add("hidden");
  document.getElementById("tone-applied-status").classList.add("hidden");

  const settings = validateSettings();
  if (!settings) return;

  toneState.tone = tone;
  document.querySelectorAll(".tone-chip").forEach((chip) => {
    chip.classList.toggle("active", chip.dataset.tone === tone);
  });

  showLoading();
  recordActionStarted("tone");
  const startedAt = Date.now();
  const phaseDurations = {};
  const readingStartedAt = Date.now();

  try {
    // Re-read the selection each time so the user can change the selected text
    // between attempts (or switch tones) without reopening the task pane.
    const selection = (await getSelectedText()).trim();
    phaseDurations.reading_email = Date.now() - readingStartedAt;
    if (!selection) {
      hideLoading();
      recordActionFinished("tone", "cancelled", {
        durationMs: Date.now() - startedAt,
        phaseDurations,
        errorKind: "invalid_response",
      });
      showError(t("tone.noSelection"));
      return;
    }

    const directive = buildToneDirective(tone);
    const rewritten = await callLibreChat(
      settings,
      selection,
      "tone",
      directive,
    );
    phaseDurations.thinking =
      Date.now() - readingStartedAt - phaseDurations.reading_email;
    toneState.rewritten = (rewritten || "").trim();

    showResponse(toneState.rewritten);
    document.getElementById("tone-actions").classList.remove("hidden");
    recordActionFinished("tone", "completed", {
      durationMs: Date.now() - startedAt,
      phaseDurations,
    });
  } catch (err) {
    hideLoading();
    recordError("tone", err, Date.now() - startedAt, phaseDurations);
    showError(err.message || t("error.unexpected"));
  }
}

async function applyTone() {
  if (!toneState.rewritten) return;

  const btn = document.getElementById("tone-apply-btn");
  const btnText = document.getElementById("tone-apply-btn-text");
  const spinner = document.getElementById("tone-apply-btn-spinner");
  const statusEl = document.getElementById("tone-applied-status");

  btn.disabled = true;
  spinner.classList.remove("hidden");
  statusEl.classList.add("hidden");

  try {
    const html = `<span style="font-family:inherit;font-size:inherit;color:inherit;">${plainTextToHtml(
      toneState.rewritten,
    )}</span>`;
    await applyToneRewrite(html);
    statusEl.textContent = t("tone.applied");
    statusEl.classList.remove("hidden");
  } catch (err) {
    showError(t("tone.applyFailed") + (err.message || ""));
  } finally {
    btn.disabled = false;
    spinner.classList.add("hidden");
    btnText.textContent = t("tone.apply");
  }
}

function setupToneView(initialTone) {
  document.getElementById("tone-section").classList.remove("hidden");

  document.querySelectorAll(".tone-chip").forEach((chip) => {
    chip.addEventListener("click", () => runToneRewrite(chip.dataset.tone));
  });
  document
    .getElementById("tone-apply-btn")
    .addEventListener("click", applyTone);

  // Only auto-run when the caller pre-selected a tone (deep link). When the
  // user opens the generic "Tone" menu item, show the options and wait for a
  // pick so they can choose the tone they want.
  if (TONE_DESCRIPTIONS[initialTone]) {
    runToneRewrite(initialTone);
  }
}

// --- Init ---

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    applyOfficeTheme();
    initI18n(ENV_APP_NAME);

    // Apply configurable branding (app name + logo) to the header/title.
    document.title = ENV_APP_NAME;
    const brandName = document.getElementById("brand-name");
    if (brandName) brandName.textContent = ENV_APP_NAME;
    const headerLogo = document.getElementById("header-logo");
    if (headerLogo) {
      headerLogo.alt = ENV_APP_NAME;
      if (ENV_APP_LOGO_URL) headerLogo.src = ENV_APP_LOGO_URL;
    }

    // Apply translations to data-i18n elements
    document.querySelectorAll("[data-i18n]").forEach((el) => {
      el.textContent = t(el.dataset.i18n);
    });
    document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
      el.placeholder = t(el.dataset.i18nPlaceholder);
    });
    document.querySelectorAll("[data-i18n-title]").forEach((el) => {
      el.title = t(el.dataset.i18nTitle);
    });

    loadSettings();

    // The Sign in button lives on the auth-signin screen; wire it once.
    document
      .getElementById("auth-signin-btn")
      .addEventListener("click", retrySignIn);

    // Startup gate: host must support NAA and the user must be signed in
    // before anything else runs. This is the only place consent can be granted.
    const signedIn = await ensureSignedIn();
    if (!signedIn) return;

    recordAdoption();

    // Agents load once, right after sign-in — the dropdown is the only setting.
    loadAgents();

    const action = new URLSearchParams(window.location.search).get("action");

    // Wire up settings gear button and back button
    document.getElementById("settings-btn").addEventListener("click", () => {
      showView("settings");
    });
    document.getElementById("back-btn").addEventListener("click", () => {
      showView("main");
    });
    document
      .getElementById("sign-out-btn")
      .addEventListener("click", handleSignOut);
    document
      .getElementById("refresh-agents-btn")
      .addEventListener("click", loadAgents);

    // Wire up agent selection persistence
    document.getElementById("agent-select").addEventListener("change", () => {
      const rs = Office.context.roamingSettings;
      rs.set("agentId", document.getElementById("agent-select").value);
      rs.saveAsync(() => {});
    });

    document.getElementById("copy-btn").addEventListener("click", () => {
      const content = document.getElementById("response-content").textContent;
      navigator.clipboard.writeText(content).catch(() => {
        // Fallback: select and copy
        const range = document.createRange();
        range.selectNodeContents(document.getElementById("response-content"));
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        document.execCommand("copy");
        sel.removeAllRanges();
      });
    });

    // Signed in: first-run users go straight to the requested action (no
    // Settings detour). Settings is only shown when explicitly requested.
    if (action === "settings") {
      showView("settings");
    } else if (action === "tone") {
      showView("main");
      // No tone param => show the four options and let the user pick.
      const tone =
        new URLSearchParams(window.location.search).get("tone") || "";
      setupToneView(tone);
    } else {
      showView("main");
      // Auto-trigger action: default to summarize if no specific action
      const currentAction = action || "summarize";
      if (currentAction === "reply") replyWithAI();
      if (currentAction === "summarize") summarizeEmail();
      if (currentAction === "compose-reply") composeReplyWithAI();
    }

    // Show custom instructions UI for compose-reply-custom action
    if (action === "compose-reply-custom") {
      document
        .getElementById("custom-instructions-section")
        .classList.remove("hidden");
      document.getElementById("custom-prompt").value =
        ACTION_DIRECTIVES["reply"];

      document
        .getElementById("custom-reply-btn")
        .addEventListener("click", composeReplyWithCustomInstructions);
    }
  }
});

// Global safety net — catch unexpected errors and display them in the taskpane
window.addEventListener("unhandledrejection", (e) => {
  const msg = e.reason?.message || String(e.reason || "Unknown error");
  showError(t("notify.unexpectedError") + msg);
});

window.addEventListener("error", (e) => {
  const msg = e.message || "Unknown error";
  showError(t("notify.unexpectedError") + msg);
});
